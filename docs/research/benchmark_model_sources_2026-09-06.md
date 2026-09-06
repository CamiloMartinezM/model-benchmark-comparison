# Model benchmark sources for September 6, 2026

This note records provider evidence retrieved on September 6, 2026. The companion [JSON dataset](../../data/benchmark_model_sources_2026-09-06.json) contains 13 model definitions, 18 price tiers, 78 benchmark records, and 24 separate historical records. Each record identifies its source. Missing task costs are `null`.

## Read prices and effort settings

Prices are USD per million tokens. Input means uncached input. Cached input means a cache read. Token prices do not determine task cost without token usage, caching, tool charges, and retries.

| Model | Input | Cached input | Output | Native effort levels | Evidence |
| --- | ---: | ---: | ---: | --- | --- |
| Claude Fable 5.1 | 10.00 | 0.25 | 50.00 | low, medium, high, xhigh, max | [API overview](https://platform.claude.com/docs/en/models/fable-5-1/overview), [effort](https://platform.claude.com/docs/en/build-with-claude/effort) |
| Claude Fable 5 | 10.00 | 1.00 | 50.00 | low, medium, high, xhigh, max | [Pricing](https://platform.claude.com/docs/en/about-claude/pricing), [effort](https://platform.claude.com/docs/en/build-with-claude/effort) |
| Claude Opus 5 | 5.00 | 0.50 | 25.00 | low, medium, high, xhigh, max | [Pricing](https://platform.claude.com/docs/en/about-claude/pricing), [effort](https://platform.claude.com/docs/en/build-with-claude/effort) |
| Claude Opus 4.8 | 5.00 | 0.50 | 25.00 | low, medium, high, xhigh, max | [Pricing](https://platform.claude.com/docs/en/about-claude/pricing), [effort](https://platform.claude.com/docs/en/build-with-claude/effort) |
| Claude Sonnet 5 | 2.00 | 0.20 | 10.00 | low, medium, high, xhigh, max | [Pricing](https://platform.claude.com/docs/en/about-claude/pricing), [effort](https://platform.claude.com/docs/en/build-with-claude/effort) |
| GLM-5.3 | 1.40 | 0.26 | 4.40 | low, high, max | [Pricing](https://docs.z.ai/guides/overview/pricing), [model card](https://huggingface.co/zai-org/GLM-5.3) |
| GLM-5.3-Flash, promotion | 0.075 | 0.015 | 0.25 | low, high, max | [Pricing](https://docs.z.ai/guides/overview/pricing), [model card](https://huggingface.co/zai-org/GLM-5.3-Flash) |
| GLM-5.3-Flash, list | 0.15 | 0.03 | 0.50 | low, high, max | [Pricing](https://docs.z.ai/guides/overview/pricing) |
| Gemini 3.8 Flash and 3.7 Flash, introductory | 0.75 | 0.075 | 3.75 | low, medium, high | [Pricing](https://ai.google.dev/gemini-api/docs/pricing?hl=en), [API guide](https://ai.google.dev/gemini-api/docs/latest-model?hl=en) |
| DeepSeek-V4-Pro-0813, off-peak | 0.66 | 0.022 | 1.98 | none, low, high, max | [Pricing](https://api-docs.deepseek.com/quick_start/pricing/), [thinking mode](https://api-docs.deepseek.com/guides/thinking_mode/) |
| DeepSeek-V4-Pro-0813, peak | 1.32 | 0.044 | 3.96 | none, low, high, max | [Pricing](https://api-docs.deepseek.com/quick_start/pricing/) |
| DeepSeek-V4-Flash-0731, off-peak | 0.22 | 0.007 | 0.66 | none, low, high, max | [Pricing](https://api-docs.deepseek.com/quick_start/pricing/), [thinking mode](https://api-docs.deepseek.com/guides/thinking_mode/) |
| DeepSeek-V4-Flash-0731, peak | 0.44 | 0.014 | 1.32 | none, low, high, max | [Pricing](https://api-docs.deepseek.com/quick_start/pricing/) |
| Kimi K3 | 3.00 | 0.30 | 15.00 | low, high, max | [Provider announcement](https://forum.moonshot.ai/t/kimi-k3-is-here-our-most-capable-model/480), [API guide](https://platform.kimi.ai/docs/guide/kimi-k3-quickstart) |
| Qwen3.8-Max, Beijing, implicit cache | 1.65 | 0.206 | 4.951 | none, low, medium, xhigh | [Model pricing](https://www.alibabacloud.com/help/en/model-studio/qwen3-8-max), [API reference](https://www.alibabacloud.com/help/en/model-studio/qwen-api-via-dashscope) |

Anthropic defaults to high. Fable reasoning remains enabled. Its cache writes cost 1.25 times input for five minutes or twice input for one hour. GLM-5.3, GLM-5.3-Flash, and Kimi K3 default to max and require reasoning. Gemini defaults to medium. DeepSeek defaults to high and permits disabling reasoning. Qwen defaults to xhigh. These are provider-specific controls, so identical effort names do not imply equal compute. [Anthropic effort](https://platform.claude.com/docs/en/build-with-claude/effort), [GLM card](https://huggingface.co/zai-org/GLM-5.3), [Kimi guide](https://platform.kimi.ai/docs/guide/kimi-k3-quickstart), [Gemini guide](https://ai.google.dev/gemini-api/docs/latest-model?hl=en), [DeepSeek guide](https://api-docs.deepseek.com/guides/thinking_mode/), [Qwen reference](https://www.alibabacloud.com/help/en/model-studio/qwen-api-via-dashscope).

GLM-5.3 is the verified flagship. GLM-5.3-Flash is a separate multimodal sibling. Its promotion ends at 24:00 Singapore time on September 9, 2026. Gemini's introductory prices expire after December 31, 2026, then double. DeepSeek peak hours are Monday through Friday, 01:00–04:00 and 06:00–10:00 UTC. Qwen's explicit cache read costs 0.137, with a separate 2.063 creation charge. [GLM documentation](https://docs.z.ai/guides/llm/glm-5.3), [Z.ai pricing](https://docs.z.ai/guides/overview/pricing), [Gemini pricing](https://ai.google.dev/gemini-api/docs/pricing?hl=en), [DeepSeek pricing](https://api-docs.deepseek.com/quick_start/pricing/), [Qwen pricing](https://www.alibabacloud.com/help/en/model-studio/qwen3-8-max).

## Compare exact HLE effort scores

The following percentages are printed point labels in figures 8.12.1.A and 8.12.1.B, pages 177–178 of the [Fable 5.1 system card](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf).

| Model | Tool configuration | Low | Medium | High | Xhigh | Max |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| Claude Fable 5.1 | With tools | 59.9 | 62.8 | 64.6 | 64.9 | 65.0 |
| Claude Fable 5.1 | No tools | 53.2 | 55.9 | 58.0 | 60.4 | 60.9 |
| Claude Fable 5 | With tools | 59.5 | 61.3 | 63.0 | 63.2 | 63.8 |
| Claude Fable 5 | No tools | 50.6 | 55.9 | 56.9 | 57.4 | 57.8 |
| Claude Opus 5 | With tools | 54.6 | 60.6 | 62.6 | 63.6 | 63.6 |
| Claude Opus 5 | No tools | 47.4 | 53.9 | 55.5 | 56.5 | 56.6 |

Anthropic caps total usage at 1M tokens and uses Opus 4.6 to grade answers. Tool runs include restricted web fetching and contamination checks. The plot's cost axis assumes perfect caching and excludes web-search charges. Costs have no exact point labels, so they remain missing. Fable evaluations include production safeguards and fallback routing. [System card, sections 8.1 and 8.12](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf), [fallback explanation](https://www.anthropic.com/claude/fable).

## Preserve benchmark versions

Selected results illustrate why versions and harnesses need separate records.

| Model | Benchmark | Effort | Score | Source details |
| --- | --- | --- | ---: | --- |
| Fable 5.1 | DeepSWE v1.1 | max | 67.4% | Five trials, [system card §8.3](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf) |
| Fable 5.1 | Terminal-Bench 4.0 | max | 55.8% | Claude Code bare, 15 trials per task, [system card §8.6](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf) |
| Fable 5.1 | Terminal-Bench-Science 0.1 | max | 52.6% | Claude Code bare, 10 trials per task, [system card §8.7](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf) |
| Fable 5.1 | FrontierCode v1.1 Extended | medium | 63.6% | Distinct from Main, [system card §8.4](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf) |
| Fable 5.1 | FrontierCode v1.1 Main | medium | 50.9% | Scope penalties reduce higher-effort scores, [system card §8.4](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf) |
| Fable 5.1 | CursorBench 3.2.0 | medium | 68.0% | Published cost 3.53 USD per task, [system card §8.8](https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20%26%20Claude%20Mythos%205.1%20System%20Card.pdf) |
| GLM-5.3 | Terminal-Bench 3.0 | max | 28.3% | Claude Code 2.1.207, three runs, 10-hour timeout, [model card](https://huggingface.co/zai-org/GLM-5.3) |
| GLM-5.3 | DeepSWE v1.1 | max | 66.9% | mini-swe-agent, 400K context, six-hour timeout, [model card](https://huggingface.co/zai-org/GLM-5.3) |
| GLM-5.3 | HLE with tools | max | 62.5% | 300K context, GPT-5.6-luna medium judge, [model card](https://huggingface.co/zai-org/GLM-5.3) |
| GLM-5.3-Flash | DeepSWE v1.1 | max | 63.4% | [API documentation](https://docs.z.ai/guides/vlm/glm-5.3-flash), [reproduction settings](https://huggingface.co/zai-org/GLM-5.3-Flash) |
| Gemini 3.8 Flash | DeepSWE v1.1 | high | 73.7% | mini-swe-agent, [model card](https://deepmind.google/models/model-cards/gemini-3-8-flash/), [methodology](https://deepmind.com/models/evals-methodology/gemini-3-8-flash) |
| Gemini 3.8 Flash | HLE-Verified | Not disclosed | 54.9% | 1,811 verified/revised items, [methodology](https://deepmind.com/models/evals-methodology/gemini-3-8-flash) |
| DeepSeek-V4-Pro-0813 | DeepSWE | max | 62.7% | DeepSeek Harness minimal, [model card](https://huggingface.co/deepseek-ai/DeepSeek-V4-Pro-0813) |
| DeepSeek-V4-Pro-0813 | HLE without tools / with tools | Not disclosed | 42.7% / 60.0% | Release effort footnote applies only to code-agent tasks, [model card](https://huggingface.co/deepseek-ai/DeepSeek-V4-Pro-0813) |

The [Gemini methodology](https://deepmind.com/models/evals-methodology/gemini-3-8-flash) retracts the image's rounded Opus 5 DeepSWE score of 74.0%. That competitor score is excluded. Its HLE-Verified set differs from original HLE.

The [DeepSeek Preview card](https://huggingface.co/deepseek-ai/DeepSeek-V4-Pro) contains effort sweeps for GPQA Diamond, HLE, LiveCodeBench, and SWE-bench. Those April weights differ from the [August release](https://huggingface.co/deepseek-ai/DeepSeek-V4-Pro-0813), so the JSON stores them under `historical_benchmarks`. No reviewed provider source supplied GPQA Diamond scores for Fable 5.1 or GLM-5.3.

The [Qwen model page](https://www.alibabacloud.com/help/en/model-studio/qwen3-8-max) lists the September 2 snapshot `qwen3.8-max-0902`. Do not attach August Qwen3.8-Max benchmark scores to that snapshot without evidence.
