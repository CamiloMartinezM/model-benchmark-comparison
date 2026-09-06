# Intelligence benchmark research — September 6, 2026

The data file contains 230 rows covering 46 model and effort combinations. Its five benchmarks are the Artificial Analysis Intelligence Index, Humanity's Last Exam, GPQA Diamond, SciCode, and CritPt. [Open the numeric rows](../../data/intelligence_rows_2026-09-06.json).

## Compare one benchmark version at a time

Use Intelligence Index v4.2 throughout its cost plot. The owner replaced constituents and weights after v4.1.1. Older launch-article index scores must stay separate. Fable 5.1 low, medium, high, and xhigh have estimated v4.2 indices and no complete-suite cost. Their individual HLE and GPQA results remain available. [Artificial Analysis methodology](https://artificialanalysis.ai/methodology/intelligence-benchmarking).

HLE uses 2,158 text-only questions from the May 2025 dataset, with no tools. The equality checker is GPT-5.6 Luna at medium effort. GPQA uses its 198-question Diamond subset with regex answer extraction. SciCode v1.0.1 reports subproblem pass@1 with scientist background supplied. CritPt covers 70 test challenges. [Artificial Analysis methodology](https://artificialanalysis.ai/methodology/intelligence-benchmarking).

## Interpret the cost axis

The index cost is a weighted average across heterogeneous benchmark tasks. It is not the price of one HLE question. Its total-suite cost is stored separately. The owner includes token-type prices and cache behavior when calculating those costs. [GPT-6 Astra model page](https://artificialanalysis.ai/models/gpt-6-astra).

HLE and GPQA costs use their own token counts and denominators. The owner publishes input, answer, and reasoning cost components per question. The numeric rows preserve those tokens, rates, and calculation fields. [HLE costs](https://artificialanalysis.ai/evaluations/humanitys-last-exam#cost), [GPQA costs](https://artificialanalysis.ai/evaluations/gpqa-diamond#cost).

Fable 5.1 uses default server fallback. Its costs blend primary and Opus 4.8 token prices using the owner's effort-specific and benchmark-specific token shares. Applying Fable prices to every token overstates the owner's costs. The shares are captured in the rows. [Fable evaluation context](https://artificialanalysis.ai/articles/claude-fable-5-1), [owner cost implementation](https://artificialanalysis.ai/_next/static/chunks/16578-995298aff8c9dcfc.js).

For each token class, the calculation is `canonical_tokens × [primary_price × (1 − fallback_share) + fallback_price × fallback_share] / 1,000,000 / task_count`. Add input, answer, and reasoning costs. HLE, GPQA, SciCode, and CritPt have no cacheable input in this snapshot. The task denominators are 2,158, 198, 288, and 70. [Owner task definitions](https://artificialanalysis.ai/_next/static/chunks/67304-b37621c20521f19c.js).

These are token-based evaluation cost estimates using rates attached to the owner snapshot. They are not invoice observations or universal request prices. Keep published costs from other benchmark owners under their stated pricing basis. An effort label represents that model's control setting, so equal labels do not imply equal computation.

## Inspect the requested models

All HLE scores in this table use the same text-only evaluation. Full-precision values and source URLs are in the JSON. [HLE leaderboard](https://artificialanalysis.ai/evaluations/humanitys-last-exam).

| Model | Effort | HLE score (%) | HLE USD/question | Index v4.2 | Index USD/weighted task |
| --- | --- | ---: | ---: | ---: | ---: |
| GPT-6 Astra | non-reasoning | 37.118 | 0.018966 | 47.85 | 1.4233 |
| GPT-6 Astra | low | 49.212 | 0.038865 | 49.32 | 0.6310 |
| GPT-6 Astra | medium | 52.734 | 0.100529 | 52.25 | 1.1612 |
| GPT-6 Astra | high | 53.058 | 0.161848 | 53.36 | 1.4110 |
| GPT-6 Astra | xhigh | 54.588 | 0.253223 | 54.31 | 1.8473 |
| GPT-6 Astra | max | 54.680 | 0.378077 | 54.66 | 2.5673 |
| GPT-5.6 Sol | non-reasoning | 16.682 | 0.005538 | 32.86 estimated | Unavailable |
| GPT-5.6 Sol | low | 39.388 | 0.022391 | 40.84 | 0.2251 |
| GPT-5.6 Sol | medium | 42.215 | 0.042960 | 45.97 | 0.3671 |
| GPT-5.6 Sol | high | 46.015 | 0.082836 | 48.30 | 0.6145 |
| GPT-5.6 Sol | xhigh | 47.312 | 0.142096 | 49.84 | 0.8891 |
| GPT-5.6 Sol | max | 49.490 | 0.268382 | 51.26 | 1.2495 |
| Claude Fable 5.1 | low | 48.888 | 0.123833 | 47.78 estimated | Unavailable |
| Claude Fable 5.1 | medium | 53.800 | 0.225417 | 49.87 estimated | Unavailable |
| Claude Fable 5.1 | high | 55.931 | 0.396984 | 51.67 estimated | Unavailable |
| Claude Fable 5.1 | xhigh | 58.712 | 1.025439 | 53.80 estimated | Unavailable |
| Claude Fable 5.1 | max | 59.129 | 1.585892 | 56.76 | 6.1169 |
| GLM-5.3 | max | 42.261 | 0.232843 | 48.58 | 1.2594 |

The comparison cohort also includes Claude Opus 5, Gemini 3.8 Flash, DeepSeek V4 Pro 0813, Kimi K3, Grok 4.6, GPT-5.6 Terra, GPT-5.6 Luna, and GLM-5.3-Flash. Only published effort variants are included. GLM-5.3 and DeepSeek V4 Pro 0813 have one published effort in this cohort. [Artificial Analysis model data](https://artificialanalysis.ai/models).

The Astra non-reasoning row is a research configuration. It remains in the audit data but is excluded from public API effort curves and efficiency comparisons.

## Verify the extraction

The extraction preserves the exact floating-point source values. It does not fill missing scores or complete-suite costs. At retrieval, 40 public HLE and GPQA chart costs matched the token calculation. The [published evidence](../../data/aa_cost_validation_evidence_2026-09-06.json) retains 35 chart comparisons after model selection. Ninety benchmark costs also match the index's weighted contributions after division by their 10% weights. Comparisons use a relative tolerance of `1e-10`.

The page's public data manifest supplies all selectable model records. Its browser loader supplies the decoding procedure. This is the same dataset that supports the owner's chart selectors. The selected records and calculation inputs are preserved in the `Source records` table of [the canonical snapshot](../../data/model_benchmark_snapshot_2026-09-06.json). [HLE page](https://artificialanalysis.ai/evaluations/humanitys-last-exam), [public page loader](https://artificialanalysis.ai/_next/static/chunks/71794-8af778c4808787fc.js).

No model-specific score uncertainty interval was supplied with these records. Small score differences and nonmonotonic effort results must not be treated as decisive improvements. Use separate benchmark panels and a Pareto frontier instead of averaging unrelated percentage scores.
