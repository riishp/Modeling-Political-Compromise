"""Build organized Excel workbooks from the BehaviorSpace CSV exports.

usage: build_excel.py <raw_csv_dir> <out_dir> <model_commit>
"""
import csv, os, statistics as st, sys
from collections import defaultdict, OrderedDict
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, LineChart, PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

RAW, OUT, COMMIT = sys.argv[1], sys.argv[2], sys.argv[3]
TRIALS = 200
WINDOW = 10
HI, LO = 85.0, 15.0
MODES = ["Imitation", "Reinforcement", "Combined"]
BUCKETS = ["0-20%", "20-40%", "40-60%", "60-80%", "80-100%"]

HEAD = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="4B3B8F")
TITLE = Font(bold=True, size=14)
BOLD = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")


def load(name):
    path = os.path.join(RAW, name + ".csv")
    rows = list(csv.reader(open(path)))
    i = next(k for k, r in enumerate(rows) if r and r[0] == "[run number]")
    h = rows[i]
    runs = defaultdict(list)
    for r in rows[i + 1:]:
        d = dict(zip(h, r))
        runs[int(d["[run number]"])].append(d)
    for k in runs:
        runs[k].sort(key=lambda d: int(d["[step]"]))
    return h, runs


def unq(v):
    return v.strip('"')


def num(v):
    return float(v)


def converge(rates):
    """rates: list of per-trial compromise rate (%) for trials 1..N.
    Returns (outcome, trial_converged, final_ma)."""
    ma = {}
    for t in range(WINDOW, len(rates) + 1):
        ma[t] = sum(rates[t - WINDOW:t]) / WINDOW
    if not ma:
        return "Did not converge", None, None
    last = max(ma)
    final = ma[last]
    for band, test in (("Converged: compromise", lambda x: x >= HI),
                       ("Converged: no compromise", lambda x: x <= LO)):
        if test(final):
            t = last
            while t - 1 in ma and test(ma[t - 1]):
                t -= 1
            return band, t, final
    return "Did not converge", None, final


def header(ws, row, cols, widths=None):
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=j, value=c)
        cell.font, cell.fill = HEAD, HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 45
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = f"A{row + 1}"   # string, so no empty cell/row is created


def run_records(runs, setting_col, setting_fmt=lambda v: v):
    recs = []
    for rn, steps in sorted(runs.items()):
        d0 = steps[0]
        rates = [num(d["compromise-rate-this-trial"]) for d in steps if int(d["[step]"]) >= 1]
        outcome, t, final = converge(rates)
        recs.append(OrderedDict([
            ("Run #", rn),
            ("Learning mode", unq(d0["learning-mode"])),
            ("Setting", setting_fmt(d0[setting_col])),
            ("Avg extremism of all agents (%)", round(num(d0["mean-extremism"]), 2)),
            ("Outcome", outcome),
            ("Trials to convergence", t),
            ("Final 10-trial avg rate of compromise (%)", round(final, 2) if final is not None else None),
            ("Mean rate of compromise over all 200 trials (%)", round(st.mean(rates), 2)),
            ("_rates", rates),
        ]))
    return recs


def write_runs(wb, title, recs, setting_label):
    ws = wb.create_sheet(title)
    cols = [k for k in recs[0] if not k.startswith("_")]
    labels = [setting_label if c == "Setting" else c for c in cols]
    header(ws, 1, labels, [8, 15, 16, 18, 26, 16, 20, 20])
    for r in recs:
        ws.append([r[c] for c in cols])
    return ws


def summarize(recs, setting_order):
    groups = OrderedDict()
    for m in MODES:
        for s in setting_order:
            g = [r for r in recs if r["Learning mode"] == m and r["Setting"] == s]
            if g:
                groups[(m, s)] = g
    out = []
    for (m, s), g in groups.items():
        conv = [r["Trials to convergence"] for r in g if r["Trials to convergence"] is not None]
        up = sum(r["Outcome"] == "Converged: compromise" for r in g)
        down = sum(r["Outcome"] == "Converged: no compromise" for r in g)
        out.append(OrderedDict([
            ("Learning mode", m),
            ("Setting", s),
            ("Avg extremism of all agents (%) - mean of runs", round(st.mean(r["Avg extremism of all agents (%)"] for r in g), 2)),
            ("Runs", len(g)),
            ("Converged: compromise (runs)", up),
            ("Converged: no compromise (runs)", down),
            ("Did not converge (runs)", len(g) - up - down),
            ("% of runs converged (either direction)", round(100 * (up + down) / len(g), 1)),
            ("% of runs converged to compromise", round(100 * up / len(g), 1)),
            ("Mean trials to convergence (converged runs only)", round(st.mean(conv), 1) if conv else None),
            ("Median trials to convergence", st.median(conv) if conv else None),
            ("Min trials to convergence", min(conv) if conv else None),
            ("Max trials to convergence", max(conv) if conv else None),
            ("Mean rate of compromise over all trials (%)", round(st.mean(r["Mean rate of compromise over all 200 trials (%)"] for r in g), 2)),
        ]))
    return out


def scatter(ws, title, xcol, ycol, rows_by_mode, xlabel, ylabel, anchor, ymax=None):
    ch = ScatterChart()
    ch.title, ch.style = title, 13
    ch.x_axis.title, ch.y_axis.title = xlabel, ylabel
    ch.height, ch.width = 9, 16
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if ymax is not None:
        ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, ymax
    for mode, (r1, r2) in rows_by_mode.items():
        xs = Reference(ws, min_col=xcol, min_row=r1, max_row=r2)
        ys = Reference(ws, min_col=ycol, min_row=r1, max_row=r2)
        s = Series(ys, xs, title=mode)
        s.marker.symbol = "circle"
        ch.series.append(s)
    ws.add_chart(ch, anchor)


def write_summary(wb, title, heading, notes, summ, setting_label, x_is_extremism=False):
    ws = wb.create_sheet(title)
    ws["A1"] = heading
    ws["A1"].font = TITLE
    for i, n in enumerate(notes, 2):
        ws.cell(row=i, column=1, value=n).alignment = Alignment(wrap_text=False)
    top = len(notes) + 3
    cols = list(summ[0].keys())
    labels = [setting_label if c == "Setting" else c for c in cols]
    header(ws, top, labels, [15, 14, 18, 7, 14, 14, 13, 14, 14, 18, 13, 12, 12, 18])
    rows_by_mode = OrderedDict()
    for k, s in enumerate(summ):
        r = top + 1 + k
        for j, c in enumerate(cols, 1):
            ws.cell(row=r, column=j, value=s[c])
        m = s["Learning mode"]
        a, _ = rows_by_mode.get(m, (r, r))
        rows_by_mode[m] = (a, r)
    xcol = 3 if x_is_extremism else 2
    xlab = "Average % extremism of all agents" if x_is_extremism else setting_label
    anchor_col = get_column_letter(len(cols) + 2)
    scatter(ws, "Trials until convergence (mean of converged runs)", xcol, 10, rows_by_mode,
            xlab, "# trials to convergence", f"{anchor_col}{top}", ymax=200)
    scatter(ws, "% of runs that converged to compromise", xcol, 9, rows_by_mode,
            xlab, "% of runs", f"{anchor_col}{top + 20}", ymax=100)
    return ws


def write_by_trial(wb, title, recs, setting_order, setting_label, chart_mode):
    ws = wb.create_sheet(title)
    ws["A1"] = (f"Average rate of compromise (% purple votes) in each trial, across the 30 runs of each setting. "
                f"Caution: an average mixes runs that converge in opposite directions; see the 'Runs' sheet or raw data for individual runs.")
    combos = [(m, s) for m in MODES for s in setting_order
              if any(r["Learning mode"] == m and r["Setting"] == s for r in recs)]
    cols = ["Trial"] + [f"{m} | {setting_label} = {s}" for m, s in combos]
    top = 3
    header(ws, top, cols, [8] + [16] * len(combos))
    for t in range(1, TRIALS + 1):
        row = [t]
        for m, s in combos:
            g = [r["_rates"][t - 1] for r in recs if r["Learning mode"] == m and r["Setting"] == s and len(r["_rates"]) >= t]
            row.append(round(st.mean(g), 2) if g else None)
        ws.append(row)
    ch = LineChart()
    ch.title = f"Graph 1 style: avg rate of compromise by trial ({chart_mode} mode)"
    ch.y_axis.title, ch.x_axis.title = "% purple votes", "Trial"
    ch.height, ch.width = 10, 22
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 100
    for j, (m, s) in enumerate(combos, 2):
        if m == chart_mode:
            ch.add_data(Reference(ws, min_col=j, min_row=top, max_row=top + TRIALS), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=top + 1, max_row=top + TRIALS))
    for s in ch.series:
        s.smooth = False
    ws.add_chart(ch, f"{get_column_letter(len(cols) + 2)}{top}")


def graph1_examples(wb, recs, label):
    """Three individual runs illustrating graph 1 with its moving average."""
    ws = wb.create_sheet("Graph 1 Examples")
    ws["A1"] = f"Graph 1 examples: individual runs ({label}). Rate of compromise per trial plus its 10-trial moving average (the 'line of best fit')."
    picks = []
    for want in ("Converged: compromise", "Converged: no compromise", "Did not converge"):
        r = next((r for r in recs if r["Outcome"] == want), None)
        if r:
            picks.append(r)
    cols = ["Trial"]
    for r in picks:
        cols += [f"Run {r['Run #']} rate (%) [{r['Outcome']}]", f"Run {r['Run #']} 10-trial moving avg (%)"]
    header(ws, 3, cols, [8] + [22] * (len(cols) - 1))
    for t in range(1, TRIALS + 1):
        row = [t]
        for r in picks:
            rates = r["_rates"]
            row.append(round(rates[t - 1], 2) if t <= len(rates) else None)
            row.append(round(sum(rates[t - WINDOW:t]) / WINDOW, 2) if WINDOW <= t <= len(rates) else None)
        ws.append(row)
    for k, r in enumerate(picks):
        ch = LineChart()
        tc = r["Trials to convergence"]
        ch.title = f"Run {r['Run #']}: {r['Outcome']}" + (f" at trial {tc}" if tc else "")
        ch.y_axis.title, ch.x_axis.title = "% purple votes", "Trial"
        ch.height, ch.width = 8, 18
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 100
        c0 = 2 + 2 * k
        ch.add_data(Reference(ws, min_col=c0, max_col=c0 + 1, min_row=3, max_row=3 + TRIALS), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + TRIALS))
        ws.add_chart(ch, f"{get_column_letter(len(cols) + 2)}{3 + 17 * k}")


def analysis4(wb, runs):
    ws = wb.create_sheet("4 Voters by Rule")
    ws["A1"] = "Analysis 4 (Level 1 / Combined mode only): who voted for each bill, by the learning rule that set the probability they voted with."
    ws["A2"] = ("Values are averages across the 30 runs. 'initial' = trial 1, when probabilities still come straight from extremism. "
                "'% imitation' = imitation voters / (imitation + reinforcement voters) for that bill, pooled across runs.")
    bills, rules = ["purple", "red", "blue"], ["imitation", "reinforcement", "initial"]
    cols = ["Trial"] + [f"{b} bill: {r} voters (avg)" for b in bills for r in rules] + \
           [f"{b} bill: % imitation among its voters" for b in bills] + ["Rate of compromise (%) (avg)"]
    top = 4
    header(ws, top, cols, [8] + [13] * (len(cols) - 1))
    totals = {(b, r): 0.0 for b in bills for r in rules}
    for t in range(1, TRIALS + 1):
        ds = [s for steps in runs.values() for s in steps if int(s["[step]"]) == t]
        row = [t]
        for b in bills:
            for r in rules:
                v = [num(d[f'voters "{b}" "{r}"']) for d in ds]
                row.append(round(st.mean(v), 2) if v else None)
                if t >= 2:
                    totals[(b, r)] += sum(v)
        for b in bills:
            im = sum(num(d[f'voters "{b}" "imitation"']) for d in ds)
            re_ = sum(num(d[f'voters "{b}" "reinforcement"']) for d in ds)
            row.append(round(100 * im / (im + re_), 1) if (im + re_) > 0 else None)
        row.append(round(st.mean(num(d["compromise-rate-this-trial"]) for d in ds), 2) if ds else None)
        ws.append(row)
    # line chart: % imitation among each bill's voters
    ch = LineChart()
    ch.title = "% of each bill's voters who used imitation"
    ch.y_axis.title, ch.x_axis.title = "% imitation", "Trial"
    ch.height, ch.width = 9, 20
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 100
    c0 = 2 + len(bills) * len(rules)
    ch.add_data(Reference(ws, min_col=c0, max_col=c0 + 2, min_row=top, max_row=top + TRIALS), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=top + 1, max_row=top + TRIALS))
    anchor = get_column_letter(len(cols) + 2)
    ws.add_chart(ch, f"{anchor}{top}")
    # pie totals table
    pr = top + TRIALS + 3
    ws.cell(row=pr, column=1, value="Totals across trials 2-200 and all 30 runs (for the pie charts)").font = BOLD
    ws.cell(row=pr + 1, column=1, value="Learning rule").font = BOLD
    for j, b in enumerate(bills, 2):
        ws.cell(row=pr + 1, column=j, value=f"{b} bill voters").font = BOLD
    for i, r in enumerate(["imitation", "reinforcement"], pr + 2):
        ws.cell(row=i, column=1, value=r)
        for j, b in enumerate(bills, 2):
            ws.cell(row=i, column=j, value=round(totals[(b, r)]))
    for k, b in enumerate(bills):
        pie = PieChart()
        pie.title = f"{b.capitalize()} bill voters by learning rule"
        pie.add_data(Reference(ws, min_col=2 + k, min_row=pr + 1, max_row=pr + 3), titles_from_data=True)
        pie.set_categories(Reference(ws, min_col=1, min_row=pr + 2, max_row=pr + 3))
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.height, pie.width = 7, 9
        ws.add_chart(pie, f"{anchor}{top + 20 + 15 * k}")


def raw_sheet(wbr, title, h, runs, keep):
    ws = wbr.create_sheet(title)
    names = {"[run number]": "Run #", "learning-mode": "Learning mode", "agents-per-party": "Agents per party",
             "lambda": "Lambda", "initial-propensity": "Initial propensity", "red-group-1": "Extremism bucket (all groups)",
             "[step]": "Trial", "compromise-rate-this-trial": "Rate of compromise (% purple votes)",
             "mean-extremism": "Avg extremism of all agents (%)"}
    cols = [c for c in keep if c in h] + [c for c in h if c.startswith("voters ")]
    ws.append([names.get(c, c.replace('voters ', 'Voters: ').replace('"', '')) for c in cols])
    for rn in sorted(runs):
        for d in runs[rn]:
            if int(d["[step]"]) < 1:
                continue
            row = []
            for c in cols:
                v = d[c]
                try:
                    v = float(v)
                    if v.is_integer():
                        v = int(v)
                except ValueError:
                    v = unq(v)
                row.append(v)
            ws.append(row)


def main():
    os.makedirs(OUT, exist_ok=True)
    wb = Workbook()
    readme = wb.active
    readme.title = "README"
    wbr = Workbook(write_only=True)

    specs = [
        ("analysis-3a-extremism", "3a", "Extremism", "red-group-1", "Extremism bucket (all 10 groups)", unq, BUCKETS,
         "Trials until convergence vs. average % extremism of all agents", True),
        ("analysis-3b-lambda", "3b", "Lambda", "lambda", "Lambda", lambda v: num(v),
         [round(x * 0.1, 1) for x in range(11)], "Trials until convergence vs. lambda (imitation dynamics)", False),
        ("analysis-3c-propensity", "3c", "Propensity", "initial-propensity", "Initial propensity", lambda v: int(num(v)),
         [1, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50], "Trials until convergence vs. initial propensity (reinforcement learning)", False),
        ("analysis-3d-population", "3d", "Population", "agents-per-party", "Agents per party", lambda v: int(num(v)),
         [3, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50], "Trials until convergence vs. population size", False),
    ]
    keep_common = ["[run number]", "learning-mode", "agents-per-party", "lambda", "initial-propensity", "red-group-1",
                   "[step]", "compromise-rate-this-trial", "mean-extremism"]
    example_recs = None
    sheet_list = []
    for name, code, short, col, label, fmt, order, heading, is_ext in specs:
        h, runs = load(name)
        recs = run_records(runs, col, fmt)
        modes_here = [m for m in MODES if any(r["Learning mode"] == m for r in recs)]
        notes = [f"Each row = one setting, summarizing its 30 runs of 200 trials. Modes run: {', '.join(modes_here)}.",
                 "Converged = 10-trial moving average of the rate of compromise enters 85-100% (or 0-15%) and stays there through trial 200.",
                 "Trials to convergence = the trial where it entered that band for good (earliest possible: 10). Runs that never settle count as 'did not converge'."]
        summ = summarize(recs, order)
        write_summary(wb, f"{code} {short} Summary", f"Analysis {code}: {heading}", notes, summ, label, is_ext)
        write_runs(wb, f"{code} {short} Runs", recs, label)
        write_by_trial(wb, f"{code} {short} By Trial", recs, order, label, "Combined")
        raw_sheet(wbr, f"{code} {short}", h, runs, keep_common)
        sheet_list += [(f"{code} {short} Summary", f"{heading}: one row per learning mode x setting, with charts."),
                       (f"{code} {short} Runs", "One row per run: outcome, trials to convergence, final moving average."),
                       (f"{code} {short} By Trial", "Average rate of compromise in each trial, per setting (graph 1 style, with chart).")]
        if name == "analysis-3b-lambda":
            example_recs = [r for r in recs if r["Learning mode"] == "Combined" and r["Setting"] == 0.5]
    graph1_examples(wb, example_recs, "Combined mode, all parameters at their constants")
    sheet_list.append(("Graph 1 Examples", "Three individual runs (converged up, converged down, did not converge) with their 10-trial moving average."))
    h4, runs4 = load("analysis-4-voters-by-rule")
    analysis4(wb, runs4)
    raw_sheet(wbr, "4 Voters by Rule", h4, runs4, keep_common)
    sheet_list.append(("4 Voters by Rule", "Analysis 4 (Level 1): voters for each bill by learning rule, by trial, with line and pie charts."))

    # README
    R = readme
    R.column_dimensions["A"].width = 34
    R.column_dimensions["B"].width = 110
    lines = [
        ("Political Compromise Model: Planned Analysis Results", None),
        (None, None),
        ("Model", f"Political_Compromise.nlogox, git commit {COMMIT} (github.com/riishp/Modeling-Political-Compromise). Run headless in NetLogo 7.0.4 BehaviorSpace."),
        ("Runs", "30 runs per setting, 200 trials (bills) per run. Data recorded every trial."),
        ("Held constant (unless swept)", "15 agents per party; all 10 extremism buckets 40-60%; initial propensity 25; lambda 0.5; keep-own-tendency? off. "
                                          "At 3 agents per party (population sweep only) the six exact sliders are set to 50%."),
        ("Modes per analysis", "Each sweep only runs in the modes where its parameter matters: 3a extremism and 3d population in all three; "
                               "3b lambda in Imitation and Combined; 3c initial propensity in Reinforcement and Combined; 4 in Combined only."),
        (None, None),
        ("KEY DEFINITIONS", None),
        ("Rate of compromise (%)", "Purple votes / total votes in ONE trial x 100. This is the model's 'Purple Vote % per Trial' plot "
                                    "(NOT the 'compromise success rate' monitor, which is purple wins / trials)."),
        ("Line of best fit", "10-trial moving average of the rate of compromise (average of the current trial and the 9 before it)."),
        ("Converged", "The moving average enters 85-100% ('Converged: compromise') or 0-15% ('Converged: no compromise') and STAYS in that band through trial 200."),
        ("Trials to convergence", "The trial where the moving average entered its final band for good. Earliest possible value is 10. Only defined for converged runs."),
        ("Did not converge", "The moving average was not inside either band for the whole end of the run (e.g. still between 15% and 85% at trial 200, or it left the band)."),
        ("Avg extremism of all agents (%)", "Mean of every agent's actual constituency extremism (random within its bucket), as in the model's monitor."),
        ("Learning rule (analysis 4)", "The rule that produced the probability an agent voted with: imitation or reinforcement from the previous trial; 'initial' on trial 1."),
        (None, None),
        ("SHEETS IN THIS WORKBOOK", None),
    ] + sheet_list + [
        (None, None),
        ("RAW DATA", "Planned_Analysis_Raw_Data.xlsx has every trial of every run (one sheet per experiment), and raw_csv/ has the original BehaviorSpace exports."),
        ("Charts", "Charts are basic previews to show what each graph could look like. They are built from the tables, so you can restyle or rebuild them."),
        ("Caution: 'did not converge'", "Includes runs that are perfectly stable but settle BETWEEN 15% and 85%. Imitation-mode runs typically level off around 70-80% "
                                         "(the level is set by lambda), so they rarely reach the 85% band. Check the 'Final 10-trial avg' column on the Runs sheets."),
        ("Caution: late convergence", "Some runs enter the band only near trial 200, so they 'stay' for just a few trials. Trials to convergence of 180+ is weak evidence of convergence."),
        ("Caution on averages", "Runs often lock in to opposite outcomes (near 0% or near 100%). Averages across runs (By Trial sheets) can sit in the middle even "
                                "though no single run did; check the Runs sheets and Graph 1 Examples."),
    ]
    for k, (a, b) in enumerate(lines, 1):
        ca = R.cell(row=k, column=1, value=a)
        cb = R.cell(row=k, column=2, value=b)
        cb.alignment = WRAP
        if a and b is None:
            ca.font = TITLE if k == 1 else BOLD
        elif a:
            ca.font = BOLD
    res = os.path.join(OUT, "Planned_Analysis_Results.xlsx")
    raw = os.path.join(OUT, "Planned_Analysis_Raw_Data.xlsx")
    wb.save(res)
    wbr.save(raw)
    print("wrote", res, os.path.getsize(res) // 1024, "KB")
    print("wrote", raw, os.path.getsize(raw) // 1024, "KB")


main()
