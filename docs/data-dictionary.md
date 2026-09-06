# Data dictionary

The Excel workbook contains one observation per row in **Results**. A row identifies its exact model, source effort, benchmark version, harness, cost unit, and provenance. Blank numeric cells represent missing values and become `NaN` in pandas.

| Field | Meaning |
| --- | --- |
| `row_id` | Unique observation identifier. Joins Results to Source records. |
| `model` | Exact model release label. Different checkpoints remain separate. |
| `effort` | Published effort label. An omitted or unspecified setting is not inferred. |
| `benchmark` | Benchmark name. Compare only matching versions and evaluation conditions. |
| `version` | Version and split or task subset. |
| `score` | Native score on a 0–100 scale. Blank means unavailable, never zero. |
| `score_unit` | percent or index. A percentage can still be partial credit, not binary success. |
| `cost_usd` | Cost amount for the cost_basis unit. Blank means unavailable. |
| `cost_basis` | usd_per_task, usd_per_weighted_index_task, usd_per_full_evaluation, or unverified_cost_unit. |
| `cost_status` | Owner-reported, repriced, token-derived, provisional, estimated, or unavailable cost provenance. |
| `harness` | The agent, tools, and evaluation controller. Its behavior affects both scores and costs. |
| `score_lower` | Lower published score interval. Blank if not published. |
| `score_upper` | Upper published score interval. Blank if not published. |
| `uncertainty` | Published interval method. No confidence bounds are invented. |
| `n_tasks` | Number of tasks or questions, when documented. |
| `n_trials` | Total repeated trials, when documented. |
| `input_tokens` | Mean input tokens per source task, when available. Check cache accounting in Source records. |
| `output_tokens` | Mean output tokens, including reasoning if the source includes it. |
| `reasoning_tokens` | Mean reasoning tokens when separately published. Do not add to output_tokens again. |
| `cached_tokens` | Mean cached input tokens when separately published. |
| `cost_explanation` | Source cost formula, rate-card assumptions, and exclusions. |
| `source_id` | Source identifier that joins to Sources. |
| `source_url` | Primary page or document supporting this observation. |
| `source_date` | Published artifact date when available. Blank means not established. |
| `retrieved_date` | Date this source was retrieved. |
| `provider` | Model provider as identified by the source. |
| `source_model` | Original identifier or label before punctuation normalization. |
| `source_effort` | Unmodified source effort description. |
| `comparison_group` | Matching benchmark, split, harness, and cost unit used for efficiency calculations. |
| `success_metric` | True only when the score represents binary task success or accuracy. |
| `efficiency_eligible` | Whether this observation can enter the descriptive within-group cost frontier. |
| `panel` | Main panel identifier. Blank keeps a result in Excel without plotting it by default. |
| `raw_cost_usd` | Original source cost before normalization, or a cost with an unverified unit. Consult Source records. |
| `notes` | Material evaluation limitations, fallback behavior, dates, and source locations. |
| `score_points_per_usd` | Derived score/cost ratio. Compare only within an evaluation group and its native units. |
| `usd_per_success` | Derived mean cost divided by success fraction. Blank for nonbinary scores or unavailable costs. |
| `pareto_efficient` | True when no eligible comparable configuration is no costlier and no worse, with one strict improvement. |
| `source_record_json` | Original extracted fields, including tokens, pricing adjustments, fallback fractions, and exact score precision. |

## Workbook tables

| Sheet | Contents |
| --- | --- |
| Read me | Interpretation and update instructions. |
| Results | Validated benchmark observations. |
| Prices | API rates, cache pricing, promotions, context tiers, and source IDs. |
| Models | Verified API model identifiers and effort controls, plus preserved comparison labels. |
| Sources | Primary-source URLs and source and retrieval dates. |
| Dictionary | Field definitions used to produce this document. |
| Source records | Original numerical records serialized as JSON, linked by `row_id`. |

The derived analysis workbook contains **Metrics**, **Efficient configurations**, and **Coverage**. These sheets are regenerated from the editable source workbook.
