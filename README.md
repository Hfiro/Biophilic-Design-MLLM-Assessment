Author: Hoda Alem
Date: November 2025
Context: Part of PhD research for a PhD in Environmental Design and Planning
at Myers-Lawson School of Construction, Virginia Tech, Blacksburg, VA.

Biophilic Design Evaluation â€“ Workspace Guide

This repository contains a small toolkit to evaluate and analyze biophilic design attributes with GPTâ€‘Vision, aggregate the results, and generate plots and summary tables. It supports a dynamic number of attributes, prompt iterations, and built spaces.

Developed by Hoda Alem for PhD research in Environmental Design and Planning at Myers-Lawson School of Construction, Virginia Tech, Blacksburg, VA.

**Prerequisites**
- Python 3.10+ recommended; create a virtualenv and install deps: `pip install -r requirements.txt`
- To use `biophilic_eval.py`, set `OPENAI_API_KEY` (e.g., in `.env` or exported)
- If Matplotlib warns about a non-writable config dir, run with `MPLBACKEND=Agg` and optionally set `MPLCONFIGDIR=.cache-mpl`

**Workbook: BD GPT Prompts.xlsx**
- Prompts sheet
  - Column B (from row 3): Attribute names (any count)
  - Columns Câ†’: Prompt text per iteration (one column per iteration)
- Spaces sheet
- Column A (from row 2): Built space names (any count). Each space must have a folder with exactly five JPG/JPEG images
- Expert sheet
  - Row 1: Space names repeated in pairs (0â€“2 then 1â€“10) â€” two adjacent columns per space
  - Row 2: Tags for those columns (ignored by code)
  - Column B (from row 3): Attribute names
  - Data rows (from row 3): Expert values for 0â€“2 and 1â€“10 in each space pair

Ordering
- Some outputs use the Expert sheet order (attributes/spaces) where applicable
- All tools work with dynamic counts of attributes, spaces, and iterations

Example folder layout (when generating new results)
```
project-root/
  BD GPT Prompts.xlsx
  requirements.txt
  Color/                      # attribute sheets generated later
  results_Theater Entrance_iterAll.xlsx   # example precomputed result (optional)
  Goodwin Entrance/           # images for this space (names must match Spaces sheet)
    img1.jpg
    img2.jpg
    img3.jpg
    img4.jpg
    img5.jpg
  Dining Area/
    a.jpg
    b.jpg
    c.jpg
    d.jpg
    e.jpg
  ...
```

**Typical Workflow (no new model runs)**
If you already have perâ€‘space result workbooks (files named like `results_<space>_iter*.xlsx`), you can start here. If you do not have those yet, see â€œCollect model outputs first (costs money)â€ below.
1) Consolidate results
   - `python results_analyzer.py --consolidated-xlsx combined_summary.xlsx --xlsx "BD GPT Prompts.xlsx" --glob "results_*.xlsx"`
   - Output: `combined_summary.xlsx` (one sheet per attribute; rows per space and Expert)
   - Columns per iteration: `itN_median_rating_0_2`, `itN_q1_rating_0_2`, `itN_q3_rating_0_2`, `itN_iqr_rating_0_2`, `itN_mean_score_1_10`, `itN_std_score_1_10`

2) Summary plots (attributes Ã— spaces grid)
   - `MPLBACKEND=Agg python plot_summary.py --xlsx combined_summary.xlsx --grid --out grid.png`
   - Single plot: `MPLBACKEND=Agg python plot_summary.py --xlsx combined_summary.xlsx --attribute "Color" --space "Dining Area" --out color_dining.png`
   - 1â€“10: mean Â± std; 0â€“2: median with Q1â€“Q3 whiskers (fallback to Â±IQR/2)
   - Iterations (x-axis) detected dynamically from headers
   - You can override the auto-detected lists:
     - `--attr-list "Color" "Natural Light" ...`
     - `--space-list "Goodwin Entrance" "Dining Area" ...`

3) Aggregate performance trends
   - Accuracy/MSE progression: `MPLBACKEND=Agg python plot_accuracy_progression.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out accuracy_progression.png`
     - Top: 0â€“1â€“2 exact-match accuracy (grouped bars per attribute + average line)
     - Bottom: 1â€“10 MSE vs Expert (grouped bars per attribute + average line)
   - Dispersion progression: `MPLBACKEND=Agg python plot_dispersion_progression.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out dispersion_progression.png`
     - Top: average group MAD (0â€“1â€“2)
     - Bottom: pooled standard deviation (1â€“10)

4) Accuracy matrices (by iteration)
   - 0â€“2 exact match: `python aggregate_accuracy.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out accuracy_by_iteration.xlsx`
   - 1â€“10 within Â±2 (configurable): `python aggregate_accuracy_10scale.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out accuracy_by_iteration_1to10.xlsx --tolerance 2.0`
   - Each workbook contains one sheet per iteration with a matrix across attributesÃ—spaces plus row/column/grand aggregates

5) Final report table (last iteration)
   - `python build_report_table.py --summary combined_summary.xlsx --expert "BD GPT Prompts.xlsx" --out final_report_table.xlsx`
   - 0â€“2: `median [Q1, Q3]` (expert appended in parentheses if mismatch)
   - 1â€“10: `mean Â± std` (expert appended in parentheses if mismatch)

Optional exports
- Table figure: `MPLBACKEND=Agg python build_report_table.py --summary combined_summary.xlsx --expert "BD GPT Prompts.xlsx" --out final_report_table.xlsx --fig-out final_report_table.svg`
- LaTeX table: `python build_report_table.py --summary combined_summary.xlsx --expert "BD GPT Prompts.xlsx" --out final_report_table.xlsx --latex-out final_report_table.tex`

**Collect model outputs first (costs money â€” calls OpenAI API)**
This repository is shared without images and without precomputed ChatGPT results. To generate your own results you must supply the input images and call the OpenAI API (incurs cost and takes time). Only `biophilic_eval.py` calls the API; all other scripts run locally.

Input images layout
- Create one folder per built space. The folder name must exactly match the space name listed in the â€œSpacesâ€ sheet of `BD GPT Prompts.xlsx` (e.g., `Goodwin Entrance`).
- Put exactly five image files in each space folder:
  - Allowed extensions: `.jpg`/`.jpeg` (caseâ€‘insensitive). Other files are ignored.
  - Filenames are arbitrary; only the count and type matter.
  - Place space folders at the project root alongside the scripts (or provide absolute paths when prompted).

Run the evaluator
- `python biophilic_eval.py --xlsx "BD GPT Prompts.xlsx"`
- You will be prompted for iteration (or â€œAllâ€), number of runs, and space (or â€œAllâ€).
- Outputs: perâ€‘space workbooks `results_<space>_iter*.xlsx`.

Resume the main workflow
- Proceed to â€œConsolidate resultsâ€ above and continue through plotting and reporting.

**Statistical conventions**
- 0â€“2 (ordinal): center = median; spread = Q1â€“Q3 for plots; average MAD for aggregate dispersion
- 1â€“10 (scalar): center = mean; spread = standard deviation; pooled std for aggregate dispersion

**Performance and reproducibility tips**
- Consolidation and plotting run locally and are fast; API calls are the slow/paid step.
- If you edit perâ€‘run cells in `results_*.xlsx`, all downstream consolidations, tables, and plots will use those corrected values.
- When using formulas in Excel, save the workbook so cached values are updated (openpyxl reads cached results in readâ€‘only mode).

**Known limitations**
- The Expert sheet layout is assumed (paired columns per space). If your structure differs, update the sheet reader helpers accordingly.
- Image inputs must be JPEG/JPG and exactly five per space folder.

If your workbook differs from the described structure, adjust paths and parameters accordingly.

---

What's New (utilities and commands)

- repair_missing_responses.py
  - Fix only connection/API error rows in `results_*.xlsx`.
  - List only: `python repair_missing_responses.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --dry-run --errors-only`
  - Repair in place: `python repair_missing_responses.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --errors-only`

- reparse_with_llm.py
  - Reparse raw text for one results workbook (text-only) to fill `rating_0_2` and `score_1_10`.
  - In-place: `python reparse_with_llm.py --file "results_<space>_iter-1.xlsx" --in-place`
  - Inspect only (no writes): add `--print-only` with `--sheets/--iterations/--runs` filters.

- refresh_iteration_summaries.py
  - Recompute per-iteration summaries in place (0â€“2: median, MAD; 1â€“10: mean, std).
  - `python refresh_iteration_summaries.py --glob "results_*.xlsx"`

- plot_scatter_all.py
  - Per-run scatter across all files. x=0/1/2 (normal-jittered), y=1â€“10 (normal-jittered), color by space; faceting by space/attribute.
  - Single figure: `python plot_scatter_all.py --glob "results_*.xlsx" --out scatter_all.png --jitter 0.1 --y-jitter 0.2 --alpha 0.5`
  - Facet by space: `python plot_scatter_all.py --glob "results_*.xlsx" --out scatter_by_space.png --facet space --ncols 3`
  - Facet by attribute: `python plot_scatter_all.py --glob "results_*.xlsx" --out scatter_by_attr.png --facet attribute --ncols 3`
  - Optional CSV of suspicious points: add `--suspicious-csv suspicious_points.csv`

Notes
- plot_summary.py y-axis now shows only the attribute name (no "Rating" prefix).
- For large grids, pass explicit `--attr-list` and `--space-list` (e.g., paginate spaces 5 per image, keep Expert order with "Prospect and Refuge" last).
- Current Spaces order (from the Spaces sheet, using display names via --names-xlsx) split into three 5-space panels:
  - `python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" --space-list "Goodwin Hall Entrance" "Lavery Hall Dining Area" "Theater Entrance" "Residential Common Area" "Administrative Service Area" --out grid_part1.png`
  - `python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" --space-list "Ambler Hall Common Area" "New Classroom Building" "Recreational Food Court" "Center for the Arts" "Football Locker Room" --out grid_part2.png`
  - `python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" --space-list "VetMed Building" "Sandy Hall Common Area" "Davidson Hall Entrance" "Liberal Arts Entrance" "BioSciences Common Area" --out grid_part3.png`

Validation plots

- Dual Expert vs Model (per-run scatter):
  - `python plot_expert_vs_model_dual.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out expert_vs_model.png --color-by iteration`
- Dual Expert vs Model (aggregated, Q1â€“Q3/Â±std):
  - `python plot_expert_vs_model_dual.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out expert_vs_model_agg.png --color-by iteration --aggregate --rating-error q1q3 --iter-filter all`
- Per-attribute panels (per-run scatter):
  - `python plot_expert_vs_model_facet_attr.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --color-by iteration --out02 expert_vs_model_by_attr_0to2_scatter.png --out10 expert_vs_model_by_attr_1to10_scatter.png`
- Per-attribute panels (aggregated, Q1â€“Q3/Â±std):
  - `python plot_expert_vs_model_facet_attr.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --aggregate --iter-filter all --color-by iteration`

Example outputs (filenames)

- Consolidation
  - `combined_summary.xlsx`

- Trends and matrices
  - `accuracy_progression.png`
  - `dispersion_progression.png`
  - `accuracy_by_iteration.xlsx`
  - `accuracy_by_iteration_1to10.xlsx`

- Final table
  - `final_report_table.xlsx`

- Scatter (all per-run points)
  - `scatter_all.png`
  - `scatter_by_space.png`
  - `scatter_by_attr.png`
  - Optional QA CSV: `suspicious_points.csv`

- Expert vs Model (dual panels)
  - Per-run scatter: `expert_vs_model.png`
  - Aggregated (Q1â€“Q3 / Â±std): `expert_vs_model_agg.png`

- Expert vs Model (per-attribute panels)
  - Per-run scatter:
    - `expert_vs_model_by_attr_0to2_scatter.png`
    - `expert_vs_model_by_attr_1to10_scatter.png`
  - Aggregated (Q1â€“Q3 / Â±std):
    - `expert_vs_model_by_attr_0to2.png`
    - `expert_vs_model_by_attr_1to10.png`


Spaces Display Names (Column B) and --names-xlsx

- You can add a second column to the Spaces sheet with anonymized display names. The scripts below will continue to match by Spaces Column A internally, but will show Column B in panel titles and legends and will enforce the order shown in the Spaces sheet.

- Supported scripts and examples:
  - plot_scatter_all.py
    - Single figure: `python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --out scatter_all.png --jitter 0.1 --y-jitter 0.2 --alpha 0.5`
    - Facet by space: `python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --facet space --ncols 5 --out scatter_by_space.png`
    - Facet by attribute: `python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --facet attribute --ncols 3 --out scatter_by_attr.png`
  - plot_summary.py
    - Grid: `MPLBACKEND=Agg python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" --out grid.png`

- Other plotting scripts will continue to use the original space names (Column A) until they add the same option.

