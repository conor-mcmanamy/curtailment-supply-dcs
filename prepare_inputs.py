# /// script
# requires-python = ">=3.9"
# dependencies = ["openpyxl>=3.1,<4"]
# ///
"""Convert input_assumptions.xlsx into model-ready, timestamped run inputs.

Run with `uv run prepare_inputs.py`. Generated Julia files define `data` and
can be loaded from model.ipynb with `include("inputs/<run>/model_inputs.jl")`.
"""

from __future__ import annotations

import copy
import json
import math
import re
import shutil
from datetime import datetime
from pathlib import Path
from statistics import NormalDist
from typing import Any, Dict, List, Optional, Sequence, Tuple

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
WORKBOOK_PATH = ROOT / "input_assumptions.xlsx"
INPUTS_DIR = ROOT / "inputs"

# Workbook assumptions do not specify an asset life. This explicit, editable
# default is used with the workbook's cost of capital to annualize capital cost.
ASSET_LIFETIME_YEARS = 20

# Override these booleans to create one-at-a-time sensitivity cases. All
# sensitivities are off by default.
SENSITIVITY_SWITCHES = {
    "events_per_year": False,
    "sigma_max": False,
    "power_capex_diesel": False,
    "power_capex_solar": False,
    "energy_capex_diesel": False,
    "energy_capex_iron_air": False,
    "fixed_om_solar": False,
    "vpp_yearly_fee": False,
    "vpp_call_price": False,
    "carbon_price": False,
    "unserved_penalty": False,
    "vpp_availability": False,
    "p_max_solar": False,
    "p_max_vpp": False,
}
SENSITIVITY_POINTS = 5

# When rows share a category, symbol, and resource, choose the row by its Input
# text. The Brattle system-wide VOLL is the default rather than the much higher
# FTI high-case value; change this mapping to select another workbook row.
DUPLICATE_ROW_CHOICES = {
    ("Reliability", "pi_UE", "All"): (
        "Penalty for unserved load (VOLL) - Brattle (system)"
    ),
}

RESOURCE_NAMES = (
    "Diesel (Tier 4)",
    "Li-ion battery",
    "Iron-air (100 h)",
    "Solar PV",
    "VPP contract",
)
RESOURCE_TYPES = ("Diesel", "BESS", "LDES", "PV", "VPP")
ENERGY_INDICES = (1, 2, 3)  # Julia's 1-based indices

# Sensitivity definitions reference raw workbook assumptions, so derived
# annualized costs are recalculated consistently for every case.
SENSITIVITY_ROWS = {
    "events_per_year": ("Curtailment", "N", "All", "Curtailment events per year"),
    "sigma_max": ("Curtailment", "sigma_max", "All", "Max share of load that can be shifted"),
    "power_capex_diesel": ("Build cost", "C^P", RESOURCE_NAMES[0], "Power build cost"),
    "power_capex_solar": ("Build cost", "C^P", RESOURCE_NAMES[3], "Power build cost"),
    "energy_capex_diesel": ("Build cost", "C^E", RESOURCE_NAMES[0], "Energy build cost"),
    "energy_capex_iron_air": ("Build cost", "C^E", RESOURCE_NAMES[2], "Energy build cost"),
    "fixed_om_solar": ("Fixed cost", None, RESOURCE_NAMES[3], "Fixed O&M"),
    "vpp_yearly_fee": ("Fixed cost", None, RESOURCE_NAMES[4], "VPP yearly fee"),
    "vpp_call_price": ("Running cost", "pi_MC", RESOURCE_NAMES[4], "VPP call price"),
    "carbon_price": ("Carbon", "pi_c", "All", "Carbon price"),
    "unserved_penalty": (
        "Reliability",
        "pi_UE",
        "All",
        DUPLICATE_ROW_CHOICES[("Reliability", "pi_UE", "All")],
    ),
    "vpp_availability": (
        "Technical",
        "alpha",
        RESOURCE_NAMES[4],
        "Availability at start of event",
    ),
    "p_max_solar": ("Site limits", "P^max", RESOURCE_NAMES[3], "Max power that can be built"),
    "p_max_vpp": ("Site limits", "P^max", RESOURCE_NAMES[4], "Max power that can be contracted"),
}
SENSITIVITY_CURVES = {key: "linear" for key in SENSITIVITY_ROWS}


def _records(workbook_path: Path) -> List[Dict[str, Any]]:
    if not workbook_path.is_file():
        raise FileNotFoundError("Assumptions workbook not found: {}".format(workbook_path))
    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    if "Inputs" not in workbook.sheetnames:
        raise ValueError("Workbook must contain an 'Inputs' sheet")

    worksheet = workbook["Inputs"]
    rows = worksheet.iter_rows(values_only=True)
    headers = next(rows, None)
    if not headers:
        raise ValueError("The 'Inputs' sheet is empty")
    header_index = {str(value).strip(): i for i, value in enumerate(headers) if value}
    required_headers = ("Category", "Symbol", "Input", "Resource", "Value (base)", "Range")
    missing = [header for header in required_headers if header not in header_index]
    if missing:
        raise ValueError("Missing workbook columns: {}".format(", ".join(missing)))

    records = []
    for row_number, row in enumerate(rows, start=2):
        if not row or not any(value is not None for value in row):
            continue
        record = {header: row[index] for header, index in header_index.items()}
        record["_row_number"] = row_number
        records.append(record)
    return records


def _find_record(
    records: Sequence[Dict[str, Any]],
    category: str,
    symbol: Optional[str],
    resource: str,
    input_name: Optional[str] = None,
) -> Dict[str, Any]:
    matches = [
        row
        for row in records
        if row.get("Category") == category
        and row.get("Symbol") == symbol
        and row.get("Resource") == resource
    ]
    if input_name is not None:
        matches = [row for row in matches if row.get("Input") == input_name]

    if len(matches) > 1:
        choice = DUPLICATE_ROW_CHOICES.get((category, symbol, resource))
        if choice is not None:
            matches = [row for row in matches if row.get("Input") == choice]
    if len(matches) != 1:
        descriptions = [
            "row {}: {}".format(row["_row_number"], row.get("Input")) for row in matches
        ]
        raise ValueError(
            "Expected one workbook row for ({}, {}, {}), found {}{}".format(
                category,
                symbol,
                resource,
                len(matches),
                ": " + "; ".join(descriptions) if descriptions else "",
            )
        )
    return matches[0]


def _number(value: Any, label: str) -> float:
    if isinstance(value, bool) or value is None:
        raise ValueError("{} must be a numeric workbook value, got {!r}".format(label, value))
    if isinstance(value, (int, float)):
        result = float(value)
    else:
        text = str(value).strip().replace(",", "")
        try:
            result = float(text)
        except ValueError as exc:
            raise ValueError("{} must be numeric, got {!r}".format(label, value)) from exc
    if not math.isfinite(result):
        raise ValueError("{} must be finite".format(label))
    return result


def _base_value(
    records: Sequence[Dict[str, Any]],
    category: str,
    symbol: Optional[str],
    resource: str,
    input_name: Optional[str] = None,
) -> float:
    row = _find_record(records, category, symbol, resource, input_name)
    return _number(row.get("Value (base)"), row.get("Input", category))


def _range_bounds(value: Any, label: str) -> Optional[Tuple[float, float]]:
    if value is None or str(value).strip().lower() in ("", "equal"):
        return None
    text = str(value).strip()
    if re.search(r"\bto\b", text, flags=re.IGNORECASE):
        parts = re.split(r"\bto\b", text, maxsplit=1, flags=re.IGNORECASE)
        low, high = (_number(part.strip(), label) for part in parts)
        return min(low, high), max(low, high)

    # Comma-separated range entries such as "0, 20, 50" are choices. Values
    # such as "1,360" are handled as a single number when they contain no
    # spaces after the comma.
    parts = [part.strip() for part in text.split(",")]
    if len(parts) == 1:
        number = _number(text, label)
        return number, number
    values = [_number(part, label) for part in parts]
    return min(values), max(values)


def _range_for_key(records: Sequence[Dict[str, Any]], key: str) -> Tuple[float, float]:
    category, symbol, resource, input_name = SENSITIVITY_ROWS[key]
    record = _find_record(records, category, symbol, resource, input_name)
    bounds = _range_bounds(record.get("Range"), key)
    if bounds is None:
        raise ValueError("Sensitivity '{}' is enabled but its workbook Range is empty".format(key))
    if bounds[0] == bounds[1]:
        raise ValueError("Sensitivity '{}' requires a nonzero range".format(key))
    return bounds


def _sweep(low: float, high: float, curve: str) -> List[float]:
    if SENSITIVITY_POINTS < 2:
        raise ValueError("SENSITIVITY_POINTS must be at least 2")
    if curve == "linear":
        return [
            low + (high - low) * index / (SENSITIVITY_POINTS - 1)
            for index in range(SENSITIVITY_POINTS)
        ]
    if curve == "normal":
        # Equally spaced quantiles of a normal curve truncated to [low, high].
        # This gives a monotone sweep with more points near the midpoint.
        normal = NormalDist()
        lower_cdf = normal.cdf(-2.0)
        upper_cdf = normal.cdf(2.0)
        values = []
        for index in range(SENSITIVITY_POINTS):
            if index == 0:
                values.append(low)
            elif index == SENSITIVITY_POINTS - 1:
                values.append(high)
            else:
                quantile = lower_cdf + (upper_cdf - lower_cdf) * index / (
                    SENSITIVITY_POINTS - 1
                )
                values.append(
                    low + (high - low) * (normal.inv_cdf(quantile) + 2.0) / 4.0
                )
        return values
    raise ValueError("Sensitivity curve must be 'linear' or 'normal', got {!r}".format(curve))


def _capital_recovery_factor(cost_of_capital: float) -> float:
    rate = cost_of_capital / 100.0 if cost_of_capital > 1 else cost_of_capital
    if not 0 <= rate:
        raise ValueError("Cost of capital must be nonnegative")
    if ASSET_LIFETIME_YEARS <= 0:
        raise ValueError("ASSET_LIFETIME_YEARS must be positive")
    if rate == 0:
        return 1.0 / ASSET_LIFETIME_YEARS
    factor = (1.0 + rate) ** ASSET_LIFETIME_YEARS
    return rate * factor / (factor - 1.0)


def _assumptions(records: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    start_text = str(
        _find_record(records, "Curtailment", "h0", "All").get("Value (base)")
    )
    starts = [int(value) for value in re.findall(r"\d+", start_text)]
    if "p.m." in start_text.lower():
        starts = [hour + 12 if 1 <= hour <= 11 else hour for hour in starts]
    duration_text = str(_find_record(records, "Curtailment", "L", "All").get("Value (base)"))
    duration_match = re.search(r"(\d+)\s+to\s+(\d+)", duration_text, flags=re.IGNORECASE)
    if duration_match:
        durations = list(
            range(int(duration_match.group(1)), int(duration_match.group(2)) + 1)
        )
    else:
        durations = [int(value) for value in re.findall(r"\d+", duration_text)]
    depth_text = str(_find_record(records, "Curtailment", "kappa", "All").get("Value (base)"))
    depths = [float(value) for value in re.findall(r"\d+(?:\.\d+)?", depth_text)]

    if not starts or not durations or not depths:
        raise ValueError("Could not parse scenario start hours, durations, or depths")
    if any(not 0 <= value <= 1 for value in depths):
        raise ValueError("Curtailment depths must be in [0, 1]")
    scenario_count = len(starts) * len(durations) * len(depths)
    expected_count = int(_base_value(records, "Curtailment", "Omega", "All"))
    if scenario_count != expected_count:
        raise ValueError(
            "Workbook scenario dimensions produce {}, but Omega is {}".format(
                scenario_count, expected_count
            )
        )

    power_capex = {
        RESOURCE_NAMES[0]: _base_value(records, "Build cost", "C^P", RESOURCE_NAMES[0]),
        RESOURCE_NAMES[1]: _base_value(records, "Build cost", "C^P", RESOURCE_NAMES[1]),
        RESOURCE_NAMES[2]: _base_value(records, "Build cost", "C^P", RESOURCE_NAMES[2]),
        RESOURCE_NAMES[3]: _base_value(records, "Build cost", "C^P", RESOURCE_NAMES[3]),
        RESOURCE_NAMES[4]: _base_value(records, "Build cost", "C^P", RESOURCE_NAMES[4]),
    }
    energy_capex = {
        name: _base_value(records, "Build cost", "C^E", name)
        for name in RESOURCE_NAMES[:3]
    }
    fixed_om = {
        name: _base_value(records, "Fixed cost", None, name, "Fixed O&M")
        for name in RESOURCE_NAMES[:4]
    }
    vpp_fee = _base_value(records, "Fixed cost", None, RESOURCE_NAMES[4], "VPP yearly fee")

    energy_efficiency = [
        _base_value(records, "Technical", "eta", name) for name in RESOURCE_NAMES[:3]
    ]
    availability_row = _find_record(
        records, "Technical", "alpha", "Diesel, Li-ion, iron-air"
    )
    solar_profile = {}
    for hour, input_name in (
        (16, "Solar output, 4 p.m."),
        (17, "Solar output, 5 p.m."),
        (18, "Solar output, 6 p.m."),
        (19, "Solar output, 7 p.m."),
        (20, "Solar output, 8 p.m. onward"),
    ):
        solar_profile[hour] = _base_value(
            records, "Technical", "alpha", RESOURCE_NAMES[3], input_name
        )

    return {
        "start_hours": starts,
        "durations": durations,
        "depths": depths,
        "expected_scenarios": expected_count,
        "events_per_year": _base_value(records, "Curtailment", "N", "All"),
        "demand_mw": _base_value(records, "Curtailment", "D", "All"),
        "sigma_max": _base_value(
            records, "Curtailment", "sigma_max", "All"
        ),
        "cost_of_capital": _base_value(records, "Build cost", None, "All", "cost of capital"),
        "power_capex_usd_per_kw": power_capex,
        "energy_capex_usd_per_kwh": energy_capex,
        "fixed_om_usd_per_kw_year": fixed_om,
        "vpp_yearly_fee_usd_per_kw_year": vpp_fee,
        "marginal_cost_usd_per_mwh": [
            _base_value(records, "Running cost", "pi_MC", name)
            for name in RESOURCE_NAMES
        ],
        "carbon_price_usd_per_tonne": _base_value(records, "Carbon", "pi_c", "All"),
        "emissions_tonnes_per_mwh": [
            _base_value(records, "Carbon", "e", RESOURCE_NAMES[0]),
            *[
                _base_value(records, "Carbon", "e", "Batteries, solar, VPP")
                for _ in RESOURCE_NAMES[1:]
            ],
        ],
        "unserved_penalty_usd_per_mwh": _base_value(
            records,
            "Reliability",
            "pi_UE",
            "All",
            DUPLICATE_ROW_CHOICES[("Reliability", "pi_UE", "All")],
        ),
        "energy_efficiency": energy_efficiency,
        "availability_during_event": _number(
            availability_row.get("Value (base)"), availability_row.get("Input")
        ),
        "vpp_availability_at_event_start": _base_value(
            records, "Technical", "alpha", RESOURCE_NAMES[4], "Availability at start of event"
        ),
        "vpp_hourly_decline": _base_value(
            records,
            "Technical",
            "delta",
            RESOURCE_NAMES[4],
            "VPP hourly decline (linear)",
        ),
        "solar_profile": solar_profile,
        "p_max_mw": [
            _base_value(records, "Site limits", "P^max", "Diesel, Li-ion, iron-air"),
            _base_value(records, "Site limits", "P^max", "Diesel, Li-ion, iron-air"),
            _base_value(records, "Site limits", "P^max", "Diesel, Li-ion, iron-air"),
            _base_value(records, "Site limits", "P^max", RESOURCE_NAMES[3]),
            _base_value(records, "Site limits", "P^max", RESOURCE_NAMES[4]),
        ],
        "e_max_mwh": [
            _base_value(records, "Site limits", "E^max", "Diesel, Li-ion, iron-air")
            for _ in RESOURCE_NAMES[:3]
        ],
    }


def _make_model_data(settings: Dict[str, Any]) -> Dict[str, Any]:
    starts = settings["start_hours"]
    durations = settings["durations"]
    depths = settings["depths"]
    scenarios = [
        (start, duration, depth)
        for start in starts
        for duration in durations
        for depth in depths
    ]
    padded_horizon = max(durations)
    scenario_count = len(scenarios)
    generator_count = len(RESOURCE_NAMES)
    capital_recovery = _capital_recovery_factor(settings["cost_of_capital"])

    # The formulation charges annualized power investment plus power-based
    # fixed O&M against P, and annualized energy investment against E:
    #   C^P = 1,000 * (power_capex[$/kW] * CRF + fixed_O&M[$/kW-year])
    #   C^E = 1,000 * energy_capex[$/kWh] * CRF
    # Fixed O&M is not added to C^E because the workbook reports it per kW-year,
    # while E is measured in kWh; adding it there would mismatch units or count
    # power-related O&M twice. VPP has no capex, so its annual fee is C^P.
    power_cap_cost = []
    for resource in RESOURCE_NAMES[:4]:
        power_cap_cost.append(
            1000.0
            * (
                settings["power_capex_usd_per_kw"][resource] * capital_recovery
                + settings["fixed_om_usd_per_kw_year"][resource]
            )
        )
    power_cap_cost.append(1000.0 * settings["vpp_yearly_fee_usd_per_kw_year"])

    energy_cap_cost = [
        1000.0 * settings["energy_capex_usd_per_kwh"][resource] * capital_recovery
        for resource in RESOURCE_NAMES[:3]
    ] + [0.0, 0.0]
    energy_max = settings["e_max_mwh"] + [0.0, 0.0]
    eta = settings["energy_efficiency"] + [1.0, 1.0]

    demand = []
    availability = []
    for start_hour, _duration, _depth in scenarios:
        demand.append([settings["demand_mw"]] * padded_horizon)
        generator_profiles = []
        generator_profiles.extend(
            [[settings["availability_during_event"]] * padded_horizon for _ in range(3)]
        )
        solar = [
            settings["solar_profile"].get(
                min(start_hour + t, max(settings["solar_profile"]))
            )
            for t in range(padded_horizon)
        ]
        generator_profiles.append(solar)
        vpp_start = settings["vpp_availability_at_event_start"]
        decline = settings["vpp_hourly_decline"]
        vpp = [
            max(0.0, min(1.0, vpp_start * (1.0 - decline * t)))
            for t in range(padded_horizon)
        ]
        generator_profiles.append(vpp)
        availability.append(generator_profiles)

    data = {
        "p_max": settings["p_max_mw"],
        "p_cap_cost": power_cap_cost,
        "energy_indices": list(ENERGY_INDICES),
        "e_max": energy_max,
        "e_cap_cost": energy_cap_cost,
        "eta": eta,
        "marginal_cost": settings["marginal_cost_usd_per_mwh"],
        "emissions_rate": settings["emissions_tonnes_per_mwh"],
        "resource_type": list(RESOURCE_TYPES),
        "availability": availability,
        "scenario_weight": [1.0 / scenario_count] * scenario_count,
        "duration": [scenario[1] for scenario in scenarios],
        "kappa": [scenario[2] for scenario in scenarios],
        "start_hour": [scenario[0] for scenario in scenarios],
        "demand": demand,
        "sigma_max": settings["sigma_max"],
        "events_per_year": settings["events_per_year"],
        "carbon_price": settings["carbon_price_usd_per_tonne"],
        "unserved_penalty": settings["unserved_penalty_usd_per_mwh"],
    }
    _validate_model_data(data, scenario_count, generator_count, padded_horizon)
    return data


def _validate_model_data(
    data: Dict[str, Any], scenario_count: int, generator_count: int, horizon: int
) -> None:
    required = (
        "p_max",
        "p_cap_cost",
        "energy_indices",
        "e_max",
        "e_cap_cost",
        "eta",
        "marginal_cost",
        "emissions_rate",
        "resource_type",
        "availability",
        "scenario_weight",
        "duration",
        "kappa",
        "start_hour",
        "demand",
        "sigma_max",
        "events_per_year",
        "carbon_price",
        "unserved_penalty",
    )
    missing = [field for field in required if field not in data]
    if missing:
        raise ValueError("Generated data is missing model fields: {}".format(", ".join(missing)))

    for field in (
        "p_max",
        "p_cap_cost",
        "e_max",
        "e_cap_cost",
        "eta",
        "marginal_cost",
        "emissions_rate",
        "resource_type",
    ):
        if len(data[field]) != generator_count:
            raise ValueError("{} must have one item per generator".format(field))
    for field in ("duration", "kappa", "start_hour", "scenario_weight"):
        if len(data[field]) != scenario_count:
            raise ValueError("{} must have one item per scenario".format(field))
    if len(data["demand"]) != scenario_count or any(
        len(row) != horizon for row in data["demand"]
    ):
        raise ValueError("demand must have shape (scenarios, padded periods)")
    if len(data["availability"]) != scenario_count or any(
        len(scenario) != generator_count
        or any(len(generator_profile) != horizon for generator_profile in scenario)
        for scenario in data["availability"]
    ):
        raise ValueError(
            "availability must have shape (scenarios, generators, padded periods)"
        )
    if not math.isclose(sum(data["scenario_weight"]), 1.0, abs_tol=1e-12):
        raise ValueError("Scenario weights must sum to 1")
    if any(
        not math.isfinite(value) or not 0 <= value <= 1
        for scenario in data["availability"]
        for generator_profile in scenario
        for value in generator_profile
    ):
        raise ValueError("Availability values must be finite and in [0, 1]")
    if any(not 0 < data["eta"][index - 1] <= 1 for index in data["energy_indices"]):
        raise ValueError("Energy-limited generator efficiencies must be in (0, 1]")


def _julia_vector(values: Sequence[Any], kind: str = "Float64") -> str:
    if kind == "String":
        rendered = [json.dumps(str(value), ensure_ascii=True) for value in values]
    elif kind == "Int":
        rendered = [str(int(value)) for value in values]
    else:
        rendered = [repr(float(value)) for value in values]
    return "{}[{}]".format(kind, ", ".join(rendered))


def _julia_data(data: Dict[str, Any]) -> str:
    scenarios = len(data["duration"])
    generators = len(data["p_max"])
    horizon = len(data["demand"][0])
    # Julia reshape is column-major: the first index (scenario) must vary
    # fastest when flattening the scenario x generator x time array.
    availability_flat = [
        data["availability"][scenario][generator][time]
        for time in range(horizon)
        for generator in range(generators)
        for scenario in range(scenarios)
    ]
    demand_flat = [
        data["demand"][scenario][time]
        for time in range(horizon)
        for scenario in range(scenarios)
    ]
    fields = [
        "data = (;",
        "    p_max = {},".format(_julia_vector(data["p_max"])),
        "    p_cap_cost = {},".format(_julia_vector(data["p_cap_cost"])),
        "    energy_indices = {},".format(_julia_vector(data["energy_indices"], "Int")),
        "    e_max = {},".format(_julia_vector(data["e_max"])),
        "    e_cap_cost = {},".format(_julia_vector(data["e_cap_cost"])),
        "    eta = {},".format(_julia_vector(data["eta"])),
        "    marginal_cost = {},".format(_julia_vector(data["marginal_cost"])),
        "    emissions_rate = {},".format(_julia_vector(data["emissions_rate"])),
        "    resource_type = {},".format(_julia_vector(data["resource_type"], "String")),
        "    availability = reshape({}, {}, {}, {}),".format(
            _julia_vector(availability_flat), scenarios, generators, horizon
        ),
        "    scenario_weight = {},".format(_julia_vector(data["scenario_weight"])),
        "    duration = {},".format(_julia_vector(data["duration"], "Int")),
        "    kappa = {},".format(_julia_vector(data["kappa"])),
        "    start_hour = {},".format(_julia_vector(data["start_hour"], "Int")),
        "    demand = reshape({}, {}, {}),".format(
            _julia_vector(demand_flat), scenarios, horizon
        ),
        "    sigma_max = {},".format(repr(float(data["sigma_max"]))),
        "    events_per_year = {},".format(repr(float(data["events_per_year"]))),
        "    carbon_price = {},".format(repr(float(data["carbon_price"]))),
        "    unserved_penalty = {}".format(repr(float(data["unserved_penalty"]))),
        ")",
    ]
    return "\n".join(fields) + "\n"


def _case_slug(key: str, index: int, value: float) -> str:
    numeric = "{:.6g}".format(value).replace("-", "m").replace(".", "p")
    return "{}_{:02d}_{}".format(key, index + 1, numeric)


def _build_sensitivity_cases(
    records: Sequence[Dict[str, Any]], base_settings: Dict[str, Any]
) -> List[Dict[str, Any]]:
    cases = []
    for key, enabled in SENSITIVITY_SWITCHES.items():
        if not enabled:
            continue
        if key not in SENSITIVITY_ROWS:
            raise ValueError("Unknown sensitivity switch: {}".format(key))
        low, high = _range_for_key(records, key)
        curve = SENSITIVITY_CURVES.get(key, "linear")
        for index, value in enumerate(_sweep(low, high, curve)):
            settings = copy.deepcopy(base_settings)
            if key == "events_per_year":
                settings["events_per_year"] = value
            elif key == "sigma_max":
                settings["sigma_max"] = value
            elif key == "power_capex_diesel":
                settings["power_capex_usd_per_kw"][RESOURCE_NAMES[0]] = value
            elif key == "power_capex_solar":
                settings["power_capex_usd_per_kw"][RESOURCE_NAMES[3]] = value
            elif key == "energy_capex_diesel":
                settings["energy_capex_usd_per_kwh"][RESOURCE_NAMES[0]] = value
            elif key == "energy_capex_iron_air":
                settings["energy_capex_usd_per_kwh"][RESOURCE_NAMES[2]] = value
            elif key == "fixed_om_solar":
                settings["fixed_om_usd_per_kw_year"][RESOURCE_NAMES[3]] = value
            elif key == "vpp_yearly_fee":
                settings["vpp_yearly_fee_usd_per_kw_year"] = value
            elif key == "vpp_call_price":
                settings["marginal_cost_usd_per_mwh"][4] = value
            elif key == "carbon_price":
                settings["carbon_price_usd_per_tonne"] = value
            elif key == "unserved_penalty":
                settings["unserved_penalty_usd_per_mwh"] = value
            elif key == "vpp_availability":
                settings["vpp_availability_at_event_start"] = value
            elif key == "p_max_solar":
                settings["p_max_mw"][3] = value
            elif key == "p_max_vpp":
                settings["p_max_mw"][4] = value
            filename = _case_slug(key, index, value)
            cases.append(
                {
                    "name": filename,
                    "parameter": key,
                    "value": value,
                    "unit": _find_record(
                        records, *SENSITIVITY_ROWS[key][:3], SENSITIVITY_ROWS[key][3]
                    ).get("Unit"),
                    "curve": curve,
                    "data": _make_model_data(settings),
                }
            )
    return cases


def load_run_inputs(run_id: str, case: Optional[str] = None) -> Dict[str, Any]:
    """Load base or named sensitivity data from ``inputs/<run_id>``."""
    run_folder = (INPUTS_DIR / run_id).resolve()
    if run_folder.parent != INPUTS_DIR.resolve() or not run_folder.is_dir():
        raise FileNotFoundError("Run folder not found under inputs/: {}".format(run_id))
    if case is None or case == "base":
        input_path = run_folder / "model_inputs.json"
    else:
        if Path(case).name != case:
            raise ValueError("case must be 'base' or a generated sensitivity case name")
        case_name = case[:-5] if case.endswith(".json") else case
        input_path = run_folder / "sensitivity_cases" / (case_name + ".json")
    if not input_path.is_file():
        raise FileNotFoundError("Input file not found for run: {}".format(input_path))
    with input_path.open("r", encoding="utf-8") as source:
        return json.load(source)


def prepare_inputs() -> Path:
    records = _records(WORKBOOK_PATH)
    settings = _assumptions(records)
    base_data = _make_model_data(settings)
    sensitivity_cases = _build_sensitivity_cases(records, settings)

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_folder = INPUTS_DIR / run_id
    run_folder.mkdir(parents=True, exist_ok=False)
    (run_folder / "model_inputs.json").write_text(
        json.dumps(base_data, indent=2, allow_nan=False) + "\n", encoding="utf-8"
    )
    (run_folder / "model_inputs.jl").write_text(
        _julia_data(base_data), encoding="utf-8"
    )

    sensitivity_dir = run_folder / "sensitivity_cases"
    if sensitivity_cases:
        sensitivity_dir.mkdir()
    manifest_cases = []
    for case in sensitivity_cases:
        json_name = case["name"] + ".json"
        julia_name = case["name"] + ".jl"
        (sensitivity_dir / json_name).write_text(
            json.dumps(case["data"], indent=2, allow_nan=False) + "\n",
            encoding="utf-8",
        )
        (sensitivity_dir / julia_name).write_text(
            _julia_data(case["data"]), encoding="utf-8"
        )
        manifest_cases.append(
            {
                "name": case["name"],
                "parameter": case["parameter"],
                "value": case["value"],
                "unit": case["unit"],
                "curve": case["curve"],
                "json_file": "sensitivity_cases/" + json_name,
                "julia_file": "sensitivity_cases/" + julia_name,
            }
        )

    manifest = {
        "run_id": run_id,
        "generated_at": datetime.now().astimezone().isoformat(),
        "source_workbook": WORKBOOK_PATH.name,
        "source_workbook_copy": "input_assumptions.xlsx",
        "asset_lifetime_years": ASSET_LIFETIME_YEARS,
        "capital_recovery_factor": _capital_recovery_factor(settings["cost_of_capital"]),
        "cost_of_capital_as_read": settings["cost_of_capital"],
        "capital_cost_method": (
            "CRF = r(1+r)^n / ((1+r)^n - 1); r is the workbook cost of capital "
            "(7.7 is interpreted as 7.7%), and n is ASSET_LIFETIME_YEARS."
        ),
        "sensitivity_curves": SENSITIVITY_CURVES,
        "sensitivity_points": SENSITIVITY_POINTS,
        "duplicate_row_choices": {
            "{}|{}|{}".format(*key): choice for key, choice in DUPLICATE_ROW_CHOICES.items()
        },
        "sensitivity_cases": manifest_cases,
        "usage": (
            "In model.ipynb run include(\"inputs/{}/model_inputs.jl\"); "
            "the generated file defines data.".format(run_id)
        ),
    }
    (run_folder / "manifest.json").write_text(
        json.dumps(manifest, indent=2, allow_nan=False) + "\n", encoding="utf-8"
    )
    shutil.copy2(WORKBOOK_PATH, run_folder / "input_assumptions.xlsx")
    return run_folder


if __name__ == "__main__":
    output_folder = prepare_inputs()
    print("Prepared model inputs: {}".format(output_folder.relative_to(ROOT)))
    print("Julia include: include({})".format(json.dumps(
        str(output_folder.relative_to(ROOT) / "model_inputs.jl")
    )))
