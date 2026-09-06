# Snapshot data

`model_benchmark_snapshot_2026-09-06.json` is the canonical release snapshot. It contains the seven workbook tables and matches the compressed JSON embedded in `model_benchmark_comparison.py`.

The supporting files preserve evidence used to assemble that snapshot:

| File | Contents |
| --- | --- |
| `coding_rows_2026-09-06.json` | 276 normalized owner observations for DeepSWE, Terminal-Bench, and FrontierCode. |
| `intelligence_rows_2026-09-06.json` | 230 Artificial Analysis observations across 46 model-effort combinations. |
| `benchmark_model_sources_2026-09-06.json` | Provider model definitions, prices, benchmark results, and separate historical results. |
| `arc_sources_2026-09-06.json` | ARC Prize evaluations, preserving distinct harnesses and original cost fields. |
| `aa_cost_validation_evidence_2026-09-06.json` | Owner chart cost components used to validate the token-based calculations at retrieval. |
| `aa_fable_fallback_rates_2026-09-06.json` | Effort-specific fallback pricing inputs for Fable 5.1. |

Supporting extraction records preserve source vocabulary and units. Use the canonical snapshot's `panel`, `cost_basis`, `comparison_group`, and `efficiency_eligible` fields for the consolidated comparison. For example, unsupported API research configurations and ARC costs with an unverified per-task unit remain audit evidence without becoming eligible plotted comparisons.

No file is a live feed. See the [methodology](../docs/methodology.md) and [source catalog](../docs/sources.md) for provenance and interpretation.
