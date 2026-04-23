Author: Hoda Alem
Date: November 2025
Context: Part of PhD research for a PhD in Environmental Design and Planning
at Myers-Lawson School of Construction, Virginia Tech, Blacksburg, VA.

Recreate Results â€“ Updated Steps

Assumptions
- You have: Python 3.10+, the Python scripts in this repo, `BD GPT Prompts.xlsx`, and one folder per built space containing exactly five `.jpg`/`.jpeg` images (folder names exactly match the Spaces sheet).
- You have an OpenAI API key set in `.env` (OPENAI_API_KEY=sk-...) or exported in your shell.

Spaces display names (optional)
- If the Spaces sheet has a second column (Column B) with anonymized display names, you can pass `--names-xlsx "BD GPT Prompts.xlsx"` to supported plotting scripts to enforce the Spaces order and show display labels while still matching by Column A.

1) Environment setup
```
python -m venv .venv
# Windows
. .\.venv\Scripts\Activate.ps1
# macOS/Linux
source .venv/bin/activate
pip install -r requirements.txt
```

2) Generate per-space results (calls OpenAI API; costs money)
```
python biophilic_eval.py --xlsx "BD GPT Prompts.xlsx"
# Follow prompts: iteration (or All), runs, spaces (or All), attributes (or All)
```

3) Repair only connection errors (optional, safe)
```
# List (no changes)
python repair_missing_responses.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --dry-run --errors-only
# Repair in place
python repair_missing_responses.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --errors-only
```

4) LLM-based parsing of raw text (optional, fills/normalizes parsed numbers)
```
# In-place for all files (PowerShell example iterating results_*.xlsx)
Get-ChildItem results_*.xlsx | ForEach-Object { \
  python reparse_with_llm.py --file $_.Name --in-place \
    --model gpt-4o-mini \
    --timeout 60
}
```

5) Refresh in-place iteration summaries in each results workbook
```
python refresh_iteration_summaries.py --glob "results_*.xlsx"
```

6) Consolidate results for plotting/reporting
```
python results_analyzer.py --consolidated-xlsx combined_summary.xlsx --xlsx "BD GPT Prompts.xlsx" --glob "results_*.xlsx"
```

7) Optional QA: flag suspicious per-run points to CSV (no plots)
```
python plot_scatter_all.py --glob "results_*.xlsx" --suspicious-csv suspicious_points.csv --suspect-tol 0.05
```

8) Scatter plots of all per-run points
```
# Single figure (colors by space; uses Spaces display names/order if present)
python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --out scatter_all.png --jitter 0.1 --y-jitter 0.2 --alpha 0.5

# Faceted by space (colors by attribute)
python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --out scatter_by_space.png --facet space --ncols 5 --jitter 0.1 --y-jitter 0.2 --alpha 0.5

# Faceted by attribute (colors by space)
python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --out scatter_by_attr.png --facet attribute --ncols 3 --jitter 0.1 --y-jitter 0.2 --alpha 0.5
```

9) Summary plots (from consolidated workbook)
```
MPLBACKEND=Agg python plot_accuracy_progression.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out accuracy_progression.png
MPLBACKEND=Agg python plot_dispersion_progression.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out dispersion_progression.png

# Dispersion aggregate lines are also saved to dispersion_aggregate_lines.xlsx (override with --agg-out)
```

10) Accuracy matrices (per iteration)
```
python aggregate_accuracy.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out accuracy_by_iteration.xlsx
python aggregate_accuracy_10scale.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out accuracy_by_iteration_1to10.xlsx --tolerance 2.0
```

11) Final report table (last iteration)
```
python build_report_table.py --summary combined_summary.xlsx --expert "BD GPT Prompts.xlsx" --out final_report_table.xlsx
```

12) Attributes Ã— Spaces grid (pagination: 6 rows Ã— 5 columns per image)
```
# Attribute order (rows): Color, Egg, Oval, and Tubular Forms, Natural Light, Spirit of Place, Transitional Spaces, Prospect and Refuge

# Page 1 (spaces 1â€“5)
python plot_summary.py --xlsx combined_summary.xlsx --grid \
  --attr-list "Color" "Egg, Oval, and Tubular Forms" "Natural Light" "Spirit of Place" "Transitional Spaces" "Prospect and Refuge" \
  --space-list "Administrative Service Area" "Ambler Hall Common Area" "BioSciences Common Area" "Center for the Arts" "Davidson Hall Entrance" \
  --out grid_spaces_p1.png

# Page 2 (spaces 6â€“10)
python plot_summary.py --xlsx combined_summary.xlsx --grid \
  --attr-list "Color" "Egg, Oval, and Tubular Forms" "Natural Light" "Spirit of Place" "Transitional Spaces" "Prospect and Refuge" \
  --space-list "Football Locker Room" "Goodwin Hall Entrance" "Lavery Hall Dining Area" "Liberal Arts Entrance" "New Classroom Building" \
  --out grid_spaces_p2.png

# Page 3 (spaces 11â€“15)
python plot_summary.py --xlsx combined_summary.xlsx --grid \
  --attr-list "Color" "Egg, Oval, and Tubular Forms" "Natural Light" "Spirit of Place" "Transitional Spaces" "Prospect and Refuge" \
  --space-list "Recreational Food Court" "Residential Common Area" "Sandy Hall Common Area" "Theater Entrance" "VetMed Building" \
  --out grid_spaces_p3.png

# Current Spaces sheet order (uses display names when present via --names-xlsx)
python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" \
  --attr-list "Color" "Egg, Oval, and Tubular Forms" "Natural Light" "Spirit of Place" "Transitional Spaces" "Prospect and Refuge" \
  --space-list "Goodwin Hall Entrance" "Lavery Hall Dining Area" "Theater Entrance" "Residential Common Area" "Administrative Service Area" \
  --out grid_part1.png
python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" \
  --attr-list "Color" "Egg, Oval, and Tubular Forms" "Natural Light" "Spirit of Place" "Transitional Spaces" "Prospect and Refuge" \
  --space-list "Ambler Hall Common Area" "New Classroom Building" "Recreational Food Court" "Center for the Arts" "Football Locker Room" \
  --out grid_part2.png
python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" \
  --attr-list "Color" "Egg, Oval, and Tubular Forms" "Natural Light" "Spirit of Place" "Transitional Spaces" "Prospect and Refuge" \
  --space-list "VetMed Building" "Sandy Hall Common Area" "Davidson Hall Entrance" "Liberal Arts Entrance" "BioSciences Common Area" \
  --out grid_part3.png
```

Notes
- If Matplotlib warns about a non-writable config dir, use: `MPLCONFIGDIR=.cache-mpl` (e.g., `MPLBACKEND=Agg MPLCONFIGDIR=.cache-mpl ...`).
- All scripts auto-detect the number of attributes, spaces, and iterations.

13) Validation plots (Expert vs Model)
```
# Dual Expert vs Model (per-run scatter)
python plot_expert_vs_model_dual.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out expert_vs_model.png --color-by iteration

# Dual Expert vs Model (aggregated, Q1â€“Q3 for 0â€“2; Â±std for 1â€“10)
python plot_expert_vs_model_dual.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --out expert_vs_model_agg.png --color-by iteration --aggregate --rating-error q1q3 --iter-filter all

# Per-attribute panels (per-run scatter)
python plot_expert_vs_model_facet_attr.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --color-by iteration --out02 expert_vs_model_by_attr_0to2_scatter.png --out10 expert_vs_model_by_attr_1to10_scatter.png

# Per-attribute panels (aggregated central value + error bars)
python plot_expert_vs_model_facet_attr.py --glob "results_*.xlsx" --xlsx "BD GPT Prompts.xlsx" --aggregate --iter-filter all --color-by iteration
```

Using Spaces display names and enforced order (optional)

If your Spaces sheet has a second column with anonymized display names, add `--names-xlsx "BD GPT Prompts.xlsx"` to supported plotting scripts to show those labels and enforce the order from the sheet (matching still uses the original names in Column A):

```
# Scatter (single figure)
python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --out scatter_all.png --jitter 0.1 --y-jitter 0.2 --alpha 0.5

# Scatter (faceted by space)
python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --facet space --ncols 5 --jitter 0.1 --y-jitter 0.2 --alpha 0.5 --out scatter_by_space.png

# Scatter (faceted by attribute)
python plot_scatter_all.py --glob "results_*.xlsx" --names-xlsx "BD GPT Prompts.xlsx" --facet attribute --ncols 3 --jitter 0.1 --y-jitter 0.2 --alpha 0.5 --out scatter_by_attr.png

# Summary grid (from consolidated workbook)
MPLBACKEND=Agg python plot_summary.py --xlsx combined_summary.xlsx --grid --names-xlsx "BD GPT Prompts.xlsx" --out grid.png
```
