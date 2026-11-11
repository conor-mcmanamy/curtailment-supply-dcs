# curtailment-supply-dcs

## Preparing model inputs

Run the standalone converter from the repository root:

```sh
uv run prepare_inputs.py
```

`uv` resolves the script's inline `openpyxl` dependency automatically. Each run
creates `inputs/<YYYYMMDD_HHMMSS>/` with `model_inputs.json`,
`model_inputs.jl`, a manifest, and a copy of the assumptions workbook. In
`model.ipynb`, load the generated Julia `data` tuple with the include path
printed by the script, replacing `<run_id>` with the generated folder name:

```julia
include("inputs/<run_id>/model_inputs.jl")
```

The script also exposes `load_run_inputs(run_id, case=None)` for reading the
JSON data in Python; pass a sensitivity case name from the run's manifest to
load that case. To prepare sensitivity files, set the relevant entries in
`SENSITIVITY_SWITCHES` to `True`; `SENSITIVITY_CURVES` selects a `linear` or
`normal` sweep separately for each parameter (all default to `linear`). Solar
availability holds its 8 p.m. value for all later hours. The default 20-year
asset life is explicitly editable as `ASSET_LIFETIME_YEARS` because the
workbook does not specify a lifetime.