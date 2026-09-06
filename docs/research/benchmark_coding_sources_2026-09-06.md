# Coding benchmark source audit

Research date: 2026-09-06. This file covers the three requested coding benchmarks. Values come from benchmark owners.

## Use the saved data

[`coding_rows_2026-09-06.json`](../../data/coding_rows_2026-09-06.json) contains 276 normalized rows after the comparison-model selection. The owner inventory in the following table describes the original retrieval, before filtering. Scores use percentage points. Cost means one attempt or rollout, including failures when the owner includes them.

| Benchmark | Revision | Subset | Configurations | Score |
| --- | --- | --- | ---: | --- |
| DeepSWE | 1.1 | 113 tasks | 70 | Attempt pass rate |
| Terminal-Bench | 4.0.0 | 66 tasks | 18 | Resolution rate |
| FrontierCode | 1.1 | Main, 100 tasks | 95 | Weighted rubric score |
| FrontierCode | 1.1 | Extended, 150 tasks | 95 | Weighted rubric score |

Use separate panels within one figure. Connect efforts only within the same model, revision, subset, and harness. Do not average these percentages into an intelligence score. FrontierCode Main is nested within Extended, so treating them as independent benchmarks would double-count tasks.

## Identify the benchmarks

DeepSWE here means Datacurve's long-horizon engineering benchmark. It is distinct from Agentica and Together AI's 2025 DeepSWE-Preview model. The benchmark owner reports mini-swe-agent across its model evaluations. [Datacurve leaderboard](https://deepswe.datacurve.ai/), [Together AI model announcement](https://www.together.ai/blog/deepswe).

Terminal-Bench 4.0 is a maintained revision with changed environments, resources, and tasks. Its announcement sets an eight-hour agent timeout and says previous trials require rerunning. The announcement date is 2026-08-28. The GitHub tag predates it, on 2026-08-26. [Benchmark announcement](https://www.tbench.ai/news/terminal-bench-4-0), [release tag](https://github.com/harbor-framework/terminal-bench/releases/tag/v4.0.0).

FrontierCode 1.1 distinguishes permitted documentation lookup from access to solution-bearing sources. Flagged runs score zero. It retains Main and Extended while retiring Diamond. A weighted rubric score differs from its separate binary blocker pass rate. [Revision methodology](https://cognition.com/blog/frontier-code-1.1), [score definitions and five-run methodology](https://cognition.com/blog/frontier-code).

## Check coverage

| Requested model | DeepSWE 1.1 | Terminal-Bench 4.0 | FrontierCode 1.1 |
| --- | --- | --- | --- |
| GPT-6 Astra | low, medium, high, xhigh, max | low, medium, high, xhigh, max | low, medium, high, xhigh, max |
| GPT-5.6 Sol | low, medium, high, xhigh, max | max | low, medium, high, xhigh, max |
| Claude Fable 5.1 | Not published in retrieved data | max | low, medium, high, xhigh, max |
| GLM-5.3 | max | max | max |

Fable 5 and Fable 5.1 remain distinct. DeepSWE reports Fable 5, so it cannot fill the Fable 5.1 gap. GLM-5.3 and GLM-5.3 Flash are separate models. The three owner datasets also differ in their DeepSeek snapshot names. DeepSWE's unversioned V4 Pro must not be relabeled as FrontierCode's V4 Pro 0813. [DeepSWE data](https://deepswe.datacurve.ai/artifacts/v1.1/leaderboard-live.json), [Terminal-Bench leaderboard](https://www.tbench.ai/), [FrontierCode data](https://cognition.com/data/frontiercode-leaderboard/data.json).

## Read the requested model values

Each cell contains score percent and mean USD cost per rollout. The three columns use different task sets and metrics. Their scores do not share an absolute difficulty scale. The data sources are the [DeepSWE owner artifact](https://deepswe.datacurve.ai/artifacts/v1.1/leaderboard-live.json), [Terminal-Bench owner page](https://www.tbench.ai/), and [FrontierCode owner artifact](https://cognition.com/data/frontiercode-leaderboard/data.json).

| Model | Effort | DeepSWE: % / USD | Terminal-Bench: % / USD | FrontierCode Main: % / USD |
| --- | --- | ---: | ---: | ---: |
| GPT-6 Astra | low | 67.04 / $2.1888 | 50.61 / $4.7191 | 45.27 / $1.5970 |
| GPT-6 Astra | medium | 72.79 / $4.3795 | 54.24 / $5.8024 | 48.83 / $2.3360 |
| GPT-6 Astra | high | 73.23 / $5.7237 | 57.88 / $6.8770 | 50.94 / $2.9290 |
| GPT-6 Astra | xhigh | 74.12 / $6.5238 | 57.88 / $7.1228 | 50.62 / $3.1710 |
| GPT-6 Astra | max | 73.23 / $12.3690 | 58.18 / $9.9005 | 53.26 / $4.4850 |
| GPT-5.6 Sol | low | 45.35 / $0.8170 | Not reported | 35.44 / $2.0998 |
| GPT-5.6 Sol | medium | 61.06 / $1.4158 | Not reported | 39.93 / $3.1245 |
| GPT-5.6 Sol | high | 69.40 / $2.6618 | Not reported | 45.06 / $4.1319 |
| GPT-5.6 Sol | xhigh | 70.73 / $3.5986 | Not reported | 46.84 / $4.9582 |
| GPT-5.6 Sol | max | 72.67 / $6.4560 | 37.27 / $7.7021 | 47.49 / $6.2938 |
| Claude Fable 5.1 | low | Not reported | Not reported | 49.82 / $2.3849 |
| Claude Fable 5.1 | medium | Not reported | Not reported | 50.91 / $3.2845 |
| Claude Fable 5.1 | high | Not reported | Not reported | 50.34 / $5.2741 |
| Claude Fable 5.1 | xhigh | Not reported | Not reported | 48.73 / $9.2738 |
| Claude Fable 5.1 | max | Not reported | 57.88 / $18.9197 | 50.28 / $12.8257 |
| GLM-5.3 | max | 68.96 / $3.9934 | 41.82 / $8.2655 | 40.14 / $16.9089 |

The published point estimates show diminishing returns in some sweeps. On DeepSWE, Astra xhigh costs less than max and has a higher point estimate. On FrontierCode Main, Fable 5.1 medium costs less than max and has a higher point estimate. This observation does not establish a statistically significant ordering. DeepSWE publishes uncertainty, while FrontierCode's retrieved data do not.

## Compare other model configurations

The following table selects each model's highest published point estimate within each benchmark. The selected effort can differ across columns. It is a compact view, not an independent evaluation. The same three owner data sources support these numerical facts.

| Model | DeepSWE: effort, % / USD | Terminal-Bench: effort, % / USD | FrontierCode Main: effort, % / USD |
| --- | ---: | ---: | ---: |
| Claude Opus 5 | max, 73.65 / $11.8376 | max, 51.82 / $18.0882 | medium, 53.38 / $4.3127 |
| Gemini 3.8 Flash | high, 73.83 / $2.3623 | high, 19.09 / $5.5417 | medium, 41.19 / $2.5954 |
| GLM-5.3 Flash | max, 63.39 / $0.2410 | Not reported | max, 31.83 / $1.1466 |
| GPT-5.6 Terra | max, 69.62 / $3.9567 | max, 21.52 / $5.2531 | max, 41.31 / $1.8060 |
| GPT-5.6 Luna | max, 67.19 / $0.6056 | max, 17.27 / $1.0505 | max, 39.81 / $0.3607 |
| Kimi K3 | max, 68.51 / $4.6547 | Not reported | none, 44.17 / $3.8167 |
| DeepSeek V4 Pro | max, 62.83 / $1.6660 | Not reported | none, 17.64 / $1.5548 |
| DeepSeek V4 Pro 0813 | Not reported | Not reported | high, 28.55 / $1.8120 |
| Grok 4.6 | medium, 67.48 / $3.4490 | high, 20.30 / $10.8836 | high, 48.01 / $2.8758 |

The source effort label `none` is retained. It does not prove that a model used no internal reasoning.

## Preserve evaluation conditions

| Benchmark | Agents in the saved comparisons | Repetitions and uncertainty | Cost basis |
| --- | --- | --- | --- |
| DeepSWE 1.1 | mini-swe-agent for every model | Four repeated benchmark passes. The artifact reports a 95% interval based on run-to-run standard error. | Mean cost per scored attempt after official frontend price adjustments. |
| Terminal-Bench 4.0 | Codex, Claude Code, mini-SWE-agent, and Grok Build | 330 trials over 66 tasks. The leaderboard reports a 95% half-width but does not specify its calculation in the retrieved payload. | Reported total USD spend divided by 330 trials. |
| FrontierCode 1.1 | Codex, Claude Code, Chisel, mini-swe-agent, Grok Build, and Cursor CLI | Original methodology specifies five runs per effort. The JSON omits completed-trial counts and confidence intervals. | Owner's mean USD spend per rollout. |

For GLM-5.3, these agents are mini-swe-agent on DeepSWE, Claude Code on Terminal-Bench, and Chisel on FrontierCode. For Gemini 3.8 Flash, FrontierCode uses Chisel while the other two use mini-swe-agent. Agent differences can affect performance, token consumption, and cost. These results measure configurations, not isolated base models. [DeepSWE artifact](https://deepswe.datacurve.ai/artifacts/v1.1/leaderboard-live.json), [Terminal-Bench embedded owner data](https://www.tbench.ai/), [FrontierCode artifact](https://cognition.com/data/frontiercode-leaderboard/data.json).

DeepSWE excludes provider, verifier, and network errors from scored attempts. It counts context-window failures and agent timeouts as failures. Fable 5 has incomplete coverage: the owner reports 73 unfinished trials out of 2,260 planned across its sweep. [DeepSWE revision report](https://deepswe.datacurve.ai/blog/deepswe-v1-1).

## Reproduce the DeepSWE cost adjustments

The raw JSON and displayed DeepSWE leaderboard costs differ for seven models. The owner applies adjustments in its frontend. Reading only `mean_cost_usd` from the JSON would produce outdated displayed costs. [Owner adjustment source](https://deepswe.datacurve.ai/assets/index-C8-z8dCr.js), [price-change history](https://deepswe.datacurve.ai/changelog).

For each affected row, the frontend calculates a ratio:

```text
uncached_input = mean_input_tokens - mean_cache_tokens
old_spend = uncached_input * old_input_rate
          + mean_cache_tokens * old_cached_rate
          + mean_output_tokens * old_output_rate
new_spend = uncached_input * new_input_rate
          + mean_cache_tokens * new_cached_rate
          + mean_output_tokens * new_output_rate
adjusted_cost = raw_mean_cost_usd * new_spend / old_spend
```

The common million-token divisor cancels. The adjustment rescales the recorded cost, rather than replacing it with the estimated token-weighted spend. The normalized JSON retains the raw cost, adjustment factor, and old and replacement rate vectors.

| Model | Recorded rates: input, cache, output | Replacement rates: input, cache, output |
| --- | --- | --- |
| GLM-5.3 Flash | $0.15, $0.03, $0.50 | $0.075, $0.015, $0.25 |
| GPT-5.6 Luna | $1.00, $0.10, $6.00 | $0.20, $0.02, $1.20 |
| GPT-5.6 Terra | $2.50, $0.25, $15.00 | $2.00, $0.20, $12.00 |
| GPT-5.6 Sol | $5.00, $0.50, $30.00 | $4.00, $0.40, $20.00 |
| DeepSeek V4 Pro, v1.1 costs | $0.435, $0.003625, $0.87 | $1.32, $0.044, $3.96 |
| DeepSeek V4 Flash | $0.14, $0.0028, $0.28 | $0.44, $0.014, $1.32 |
| Gemini 3.6 Flash | $1.50, $0.15, $7.50 | $0.75, $0.075, $3.75 |

Rates in this table are USD per million tokens and reproduce the owner frontend. They are not a separately verified retail price catalog. The DeepSeek replacement costs use peak rates. The owner's history describes lower off-peak rates and temporary promotions for some other models. [Owner repricing logic](https://deepswe.datacurve.ai/assets/index-C8-z8dCr.js), [owner change history](https://deepswe.datacurve.ai/changelog).

Astra is a separate caveat. Its DeepSWE cost uses expected launch prices: $12 per million uncached input tokens, $15 for cache writes, $1.20 for cache reads, $50 for output, and $2 per million compute units. Its cost includes more than output reasoning tokens. Treat this as an expected-rate estimate for measured usage. [Astra cost basis in owner data](https://deepswe.datacurve.ai/artifacts/v1.1/leaderboard-live.json).

The Terminal-Bench and FrontierCode payloads do not specify complete per-row rate cards or repricing dates. Their costs remain the owners' published values. Do not apply DeepSWE adjustments to them without separate evidence.

## Audit the extraction

The data were retrieved through HTTPS on September 6, 2026. The [canonical snapshot](../../data/model_benchmark_snapshot_2026-09-06.json) preserves original values and pricing calculations in its `Source records` table. The [source catalog](../sources.md) lists the owner URLs and retrieval dates.

Terminal-Bench embeds its leaderboard in a serialized JSON hydration payload. Extraction decoded this payload without executing downloaded JavaScript. FrontierCode and DeepSWE expose dedicated JSON endpoints.

At retrieval, all 70 adjusted DeepSWE means matched the corresponding rendered leaderboard costs at relative tolerance `1e-12`. The published comparison retains 68 DeepSWE rows, 18 Terminal-Bench rows, and 190 FrontierCode rows after model selection. Each FrontierCode row preserves `new_score` as the plotted metric and `correct` as `blocker_pass_rate`.
