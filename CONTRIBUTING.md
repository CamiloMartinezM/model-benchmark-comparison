# Contribute changes

Report a source correction or proposed change in an issue. Include the model identifier, effort setting, benchmark version, harness, source URL, retrieval date, and the values that need attention.

## Set up the environment

Use Python 3.12 or 3.13 and install the verified dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-verified.txt
```

For Windows PowerShell, use `.venv\Scripts\Activate.ps1` to activate the environment.

Run the project checks before submitting a pull request:

```bash
python -m ruff check .
python -m ruff format --check .
python -m mypy
python -m pytest -q
python model_benchmark_comparison.py --validate-only
```

Document functions with Google-style docstrings. A `Returns:` section includes the explicit return type, matching the annotation. Add tests for behavior changes that could affect comparisons, calculations, workbook preservation, or figure interpretation.

## Change benchmark data

Use benchmark-owner data or provider documentation and record its provenance. Preserve unknown values as blanks. Do not infer effort settings, costs, uncertainty intervals, or model snapshot identities from nearby records.

Keep benchmark versions, subsets, harnesses, tool settings, and cost units separate. A price per million tokens is not a task cost without usage information. Preserve original costs and calculation inputs whenever you reprice a result.

The editable workbook is the input for an ordinary rerun. A snapshot release also keeps the embedded data in the standalone script and `data/model_benchmark_snapshot_2026-09-06.json` consistent. Update all three representations when changing a published snapshot. Preserve earlier releases rather than silently replacing their dates or methods.

Generate outputs in a fresh directory to check the embedded snapshot:

```bash
python model_benchmark_comparison.py --output-dir scratch/reproduced
```

Compare the workbook values and CSV data, then inspect the PNG and HTML figures. Spreadsheet archives and figure metadata can differ between runs, so byte-for-byte output equality is not required. Update the README inventory, coverage, source catalog, and research notes when their facts change.

Keep credentials, local environments, caches, and unrelated files out of commits. Preserve third-party source attribution and license notices.
