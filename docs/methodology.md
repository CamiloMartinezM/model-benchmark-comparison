# Methodology and interpretation

Source snapshot: September 6, 2026. The program analyzes published results and does not rerun the underlying model evaluations.


Higher scores and lower costs place a point toward the upper left within a panel. Colors identify models, and marker shapes identify effort settings. Lines connect published settings in effort order, including cases where more effort costs less or scores lower.

The panels preserve different cost units and score meanings. Coding results use attempt or rollout costs. HLE uses cost per question. The AA Index uses a weighted task across its component benchmarks. ARC-AGI-3 uses total evaluation spend. Compare ratios within a panel and compatible setup, rather than averaging scores across panels.

Several results illustrate why the effort setting matters:

- On FrontierCode 1.1 Extended, Fable 5.1 medium scores 63.60 at $2.6825 per rollout. Max scores 62.03 at $10.7206. These are point estimates without published confidence intervals. [Cognition data](https://cognition.com/data/frontiercode-leaderboard/data.json).
- On the same FrontierCode split, Astra high scores 63.07 at $2.5140. Max scores 64.48 at $3.7880. The agent setups differ across providers. [Cognition data](https://cognition.com/data/frontiercode-leaderboard/data.json).
- On AA's text-only HLE, Astra max scores 54.680% at $0.37808 per question. Fable 5.1 max scores 59.129% at $1.58589. Fable's reported configuration includes fallback behavior. [Artificial Analysis HLE](https://artificialanalysis.ai/evaluations/humanitys-last-exam).
- ARC Prize's Astra results show that higher effort can reduce total cost by reducing the number of actions. The Standard and Provider Adapter harnesses produce substantially different results, so the figure keeps both configurations explicit. [ARC Prize evaluation](https://arcprize.org/blog/astra).

## ARC-AGI-3 harnesses

Astra appears in two curves because ARC Prize evaluated two harnesses. Standard carries forward model-written visible notes. Provider Adapter also preserves opaque reasoning state between requests and uses context compaction. Both curves use the same model, with different state management. Direct labels identify each curve in the static and interactive figures. [ARC Prize evaluation](https://arcprize.org/blog/astra).

At the same **max** effort setting:

| Harness | Action efficiency score | Full evaluation cost |
| --- | --- | --- |
| Standard (solid) | 62.7% | $26,098 |
| Provider Adapter (dashed) | 98.6% | $17,332 |

These are separate configurations, not duplicate observations. Their efficiency calculations remain in separate comparison groups. Costs can decrease as effort increases because the model can finish with fewer actions. [ARC Prize evaluation](https://arcprize.org/blog/astra).

## Preserve the limitations

The main figure uses benchmark-owner data. Provider launch tables remain separate observations. For example, Terminal-Bench's owner results differ from launch-table scores because evaluations differ. The workbook does not substitute one source's effort setting or score for another's.

GLM-5.3 is the verified flagship used here. Published owner data usually include only its max setting. Terminal-Bench has only max for Sol and Fable 5.1. DeepSWE's owner table omits Fable 5.1, while Anthropic reports a 67.4% result without an exact corresponding cost. These gaps remain explicit. [GLM model card](https://huggingface.co/zai-org/GLM-5.3), [Terminal-Bench](https://www.tbench.ai/), [DeepSWE](https://deepswe.datacurve.ai/).

DeepSWE prices Astra with a provisional rate card that differs from the published direct API rates. Hollow points identify these estimates. Other DeepSWE costs include the owner's price adjustments. Terminal-Bench and FrontierCode do not publish complete per-row billing rate cards. [DeepSWE source data](https://deepswe.datacurve.ai/artifacts/v1.1/leaderboard-live.json), [OpenAI API prices](https://developers.openai.com/api/docs/pricing).

The AA Index panel uses version 4.2 throughout. It omits estimated indices that lack complete evaluation costs. HLE uses 2,158 text-only questions with no tools. Full-set provider HLE scores, tool-enabled HLE scores, and their different graders remain separate. Costs derived from owner token counts are estimates at the source's rates, not invoice observations. [Artificial Analysis methodology](https://artificialanalysis.ai/methodology/intelligence-benchmarking).

Small score differences do not establish a statistically significant ranking. DeepSWE reports intervals over repeated full-benchmark runs. Terminal-Bench reports its own intervals. The script does not invent intervals for the remaining data.

## Calculate efficiency

The script calculates `score_points_per_usd = score / cost_usd` as a descriptive ratio. It does not normalize different benchmarks into a universal intelligence-per-dollar score. Index points and weighted rubric scores are not success probabilities.

For binary success metrics with cost per task, `usd_per_success = cost_usd / (score / 100)`. Costs include failed attempts when the source includes them. Zero success rates and missing costs remain missing in this derived field. The calculation does not assume repeated attempts are independent.

An eligible Pareto configuration has no cheaper result with an equal or higher score, and no result at the same cost with a higher score, within the same comparison group. Equivalent ties remain efficient. Groups keep benchmark, version, score unit, cost unit, and harness fixed. Provisional costs, known outdated prices, incomplete estimated costs, and unsupported API research configurations do not displace eligible observations.

Efficiency flags use the full eligible workbook group. Filtering the figure changes the displayed models; it does not recalculate the exported Pareto flags for a smaller cohort.

## Validate and reproduce the snapshot

The script checks required columns, unique observation IDs, source references, numeric ranges, interval bounds, boolean values, and comparison-group consistency before plotting. It rejects invalid data rather than replacing missing values with zero.

The [canonical JSON snapshot](../data/model_benchmark_snapshot_2026-09-06.json) contains the same tables embedded in the standalone script. The Excel workbook takes precedence on subsequent runs, so edits are preserved. Use a fresh output directory to regenerate the embedded release data.

The [research notes](research/README.md) record extraction and pricing validation. They distinguish the source owner's published costs from costs calculated from source token counts and rates. Validation of a cost formula does not establish that the underlying benchmark measurements are statistically reliable or transferable to another workload.
