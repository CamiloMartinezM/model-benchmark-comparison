# Model benchmark performance and cost

Compare how published reasoning settings affect benchmark scores and evaluation costs. A standalone Python script reads editable Excel tables into pandas and produces one figure with six benchmark panels.

[![Checks](https://github.com/CamiloMartinezM/model-benchmark-comparison/actions/workflows/checks.yml/badge.svg)](https://github.com/CamiloMartinezM/model-benchmark-comparison/actions/workflows/checks.yml)

The September 6, 2026 snapshot contains 678 observations, 43 exact model labels, 27 benchmark names, and 89 source links. The default figure shows 11 models: GPT-6 Astra, GPT-5.6 Sol, Claude Fable 5.1, GLM-5.3, Claude Opus 5, Gemini 3.8 Flash, DeepSeek V4 Pro 0813, Kimi K3, GPT-5.6 Terra, GPT-5.6 Luna, and Grok 4.6.

![Model performance and evaluation cost across six benchmarks](model_benchmark_comparison/model_benchmark_comparison.png)

Higher and farther left indicates a higher score for less cost **within one panel**. Colors identify models. Marker shapes identify effort settings. Lines connect published settings in effort order, including cases where additional effort costs less or scores lower. Benchmark and harness differences remain explicit.

## View the results

The repository includes the following outputs:

| File | Purpose |
| --- | --- |
| [Interactive HTML](model_benchmark_comparison/model_benchmark_comparison.html) | Hover for exact values and sources, filter models, and zoom. Download the file and open it locally. |
| [Editable Excel dataset](model_benchmark_comparison/model_benchmarks.xlsx) | Source observations, API prices, model settings, data dictionary, and audit records. |
| [Analysis workbook](model_benchmark_comparison/model_benchmark_analysis.xlsx) | Derived efficiency measures, eligible Pareto configurations, and coverage. |
| [PDF](model_benchmark_comparison/model_benchmark_comparison.pdf), [SVG](model_benchmark_comparison/model_benchmark_comparison.svg), [PNG](model_benchmark_comparison/model_benchmark_comparison.png) | Export or share the consolidated figure. |
| [Analysis CSV](model_benchmark_comparison/analysis.csv), [coverage CSV](model_benchmark_comparison/coverage.csv) | Read the results in other analysis tools. |
| [Readable JSON snapshot](data/model_benchmark_snapshot_2026-09-06.json) | Inspect the same tables embedded in the Python script. |

The HTML embeds Plotly's JavaScript bundle and works offline. GitHub displays HTML source rather than running it. Use the file's download control or download a [release](https://github.com/CamiloMartinezM/model-benchmark-comparison/releases).

## Run the comparison

Use Python 3.12 or later. Clone the repository, and then choose an installation method:

```bash
git clone https://github.com/CamiloMartinezM/model-benchmark-comparison.git
cd model-benchmark-comparison
```

With `uv`, the script's dependency metadata handles installation:

```bash
uv run model_benchmark_comparison.py
```

With Python and pip:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python model_benchmark_comparison.py
```

On Windows PowerShell, activate the environment with `.venv\Scripts\Activate.ps1` instead of the `source` command.

The script embeds the complete snapshot. You can copy the Python file alone and run it with its dependencies installed. Plot generation makes no network requests and runs no paid model evaluations. It does not refresh benchmark results from the web.

## Edit or filter the comparison

Edit the **Results** sheet in `model_benchmark_comparison/model_benchmarks.xlsx`, and then rerun the script. It preserves the input workbook and refreshes the separate analysis files and figures.

To show only the four focus models:

```bash
python model_benchmark_comparison.py \
  --models "GPT-6 Astra" "GPT-5.6 Sol" "Claude Fable 5.1" "GLM-5.3" \
  --output-dir scratch/focus
```

To use an edited workbook in a separate output directory:

```bash
python model_benchmark_comparison.py \
  --workbook model_benchmark_comparison/model_benchmarks.xlsx \
  --output-dir scratch/edited
```

If `--workbook` is omitted, the script reads or creates a workbook in the output directory. A fresh directory starts from the embedded snapshot. An existing workbook takes precedence over that snapshot.

Additional options are available:

| Option | Behavior |
| --- | --- |
| `--all-models` | Include every model label in the workbook. The legend can become crowded. |
| `--validate-only` | Check the workbook without exporting figures. |
| `--no-html` | Export the static figures without the interactive HTML file. |
| `--dpi 300` | Set PNG resolution. Accepted values range from 50 to 600. |
| `--help` | Display all command-line options. |

## Understand the comparisons

Each panel keeps its benchmark-specific score and cost unit:

| Panel | Score | Cost axis |
| --- | --- | --- |
| DeepSWE v1.1 | Repository issues resolved, percent | USD per scored attempt |
| Terminal-Bench 4.0 | Terminal tasks resolved, percent | USD per trial |
| FrontierCode 1.1 Extended | Weighted code-quality rubric, percent | USD per rollout |
| Humanity's Last Exam | Artificial Analysis text-only accuracy, percent | USD per question |
| AA Intelligence Index v4.2 | Composite index points | USD per weighted index task |
| ARC-AGI-3 | Action efficiency, percent | USD for the full evaluation |

The workbook also includes benchmarks such as GPQA Diamond, SciCode, CritPt, CursorBench, and provider-reported evaluations. Rows without compatible score-cost pairs remain in the data and do not become plotted zeroes.

The results show why effort deserves separate attention. On FrontierCode Extended, Fable 5.1 medium scores 63.60 at $2.6825 per rollout. Max scores 62.03 at $10.7206. These point estimates have no published confidence intervals. [Cognition source data](https://cognition.com/data/frontiercode-leaderboard/data.json).

The following limitations affect interpretation:

- Effort labels do not imply equal compute across providers. Agent harnesses, graders, tool access, and benchmark versions can differ.
- DeepSWE uses provisional prices for Astra. Hollow points flag them, and efficiency calculations exclude them.
- DeepSWE has no owner-published Fable 5.1 score-cost pair. Terminal-Bench publishes only max effort for Sol and Fable 5.1 in this snapshot.
- The HLE panel uses 2,158 text-only questions with no tools. Provider HLE variants stay separate. Fable costs incorporate the source's fallback routing and rates.
- The AA Index panel uses version 4.2. Estimated indices without complete evaluation costs are omitted.
- ARC-AGI-3 costs describe the full evaluation. Its Standard and Provider Adapter harnesses remain separate. Sol retains its original evaluation prices.
- Small score differences do not establish statistically significant rankings. Published uncertainty intervals appear when available.

The [methodology](docs/methodology.md), [data dictionary](docs/data-dictionary.md), [source catalog](docs/sources.md), and [research notes](docs/research/README.md) explain these choices and link to the evidence.

## Verify the code

Install the tested environment, and then run the checks:

```bash
python -m pip install -r requirements-verified.txt
python -m ruff check .
python -m ruff format --check .
python -m mypy
python -m pytest -q
python model_benchmark_comparison.py --validate-only
```

Ruff checks Google-style docstrings as well as code style. Mypy runs in strict mode. The tests cover cost normalization, Pareto comparisons, missing data, workbook edits, validation, uncertainty rendering, and generation from a copy of the script alone. GitHub Actions runs these checks on Python 3.12 and 3.13 and generates all figure formats from a fresh directory.

See [CONTRIBUTING.md](CONTRIBUTING.md) for changes to code or benchmark data.

## Attribution and license

Maintained by [Camilo Martínez M.](https://github.com/CamiloMartinezM). This project consolidates published evaluations; it does not claim to have run the benchmarks itself.

Original analysis code and documentation use the [MIT license](LICENSE). Benchmark data retain their source attribution and applicable upstream terms. The embedded Plotly bundle retains its own license. See [third-party notices](THIRD_PARTY_NOTICES.md).
