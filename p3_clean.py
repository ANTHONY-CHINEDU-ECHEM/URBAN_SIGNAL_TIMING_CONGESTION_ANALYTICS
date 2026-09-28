"""Project 3: Urban Signal Timing and Congestion Cost Simulator. Cleaning pipeline."""
import glob
import numpy as np
import pandas as pd
from lib.cleaning import Cleaner

SRC = sorted(glob.glob("data/raw/*.xlsx"))[0]
BOOL = ["coordination_flag", "adaptive_control_flag", "peak_hour_flag", "school_zone_flag", "construction_zone_flag",
        "event_day_flag", "signal_malfunction_flag"]
N = lambda **k: dict(**k)
PEAK = {"07:00-09:00": "AM peak", "15:00-18:00": "PM peak"}
LOS_BREAKS = [(10, "A"), (20, "B"), (35, "C"), (55, "D"), (80, "E")]

CFG = dict(
    key="record_id", anchor="sensor_camera_id", key_prefix="TRF-", key_width=7, flag_identity=[],
    categorical=["corridor_name", "time_of_day", "day_of_week", "district_zone", "jurisdiction", "road_type",
                 "weather_condition", "level_of_service", "baseline_or_optimized"],
    canonical_override={"level_of_service": {x.lower(): x for x in "ABCDEF"}},
    id_text=["record_id", "intersection_id", "sensor_camera_id", "signal_id", "retiming_scenario_id", "simulation_run_id"],
    bool=BOOL, dates=["observation_date"],
    numeric={
        "speed_limit_mph": N(lo=15, hi=75, fence=False, unit="mph", integer=True, rationale="Posted urban and suburban limits"),
        "vehicle_count": N(lo=0, unit="vehicles", rationale="Counts cannot be negative; extreme counts above the fence removed"),
        "avg_speed_mph": N(lo=0, hi=85, unit="mph", rationale="Speeds cannot be negative; above 85 mph is a detector error"),
        "cycle_length_sec": N(lo=30, hi=240, fence=False, unit="seconds", rationale="Practical signal cycle range"),
        "green_time_main_sec": N(lo=1, hi=240, fence=False, unit="seconds", rationale="Must be positive and within a cycle"),
        "green_time_side_sec": N(lo=1, hi=240, fence=False, unit="seconds", rationale="Must be positive and within a cycle"),
        "yellow_time_sec": N(lo=3, hi=6, fence=False, unit="seconds", rationale="Standard yellow change interval range"),
        "all_red_time_sec": N(lo=0, hi=6, fence=False, unit="seconds", rationale="Red clearance range"),
        "num_phases": N(lo=2, hi=8, fence=False, unit="phases", integer=True, rationale="Dual ring controller maximum"),
        "avg_delay_sec_per_vehicle": N(lo=0, unit="s/veh", rationale="Delay cannot be negative; sentinel 128 caught by fence"),
        "queue_length_vehicles": N(lo=0, unit="vehicles", rationale="Cannot be negative"),
        "v_c_ratio": N(lo=0, hi=2, fence=False, unit="ratio", rationale="Degree of saturation, oversaturation allowed to 2.0"),
        "fuel_wasted_gallons": N(lo=0, unit="gallons", rationale="Cannot be negative"),
        "co2_emissions_kg": N(lo=0, unit="kg", rationale="Cannot be negative"),
        "congestion_cost_usd": N(lo=0, unit="USD", rationale="Magnitude errors above the fence removed"),
        "pedestrian_volume": N(lo=0, fence=False, unit="persons", rationale="Count"),
        "bike_volume": N(lo=0, fence=False, unit="bikes", rationale="Count"),
        "travel_time_index": N(lo=0.5, hi=5, fence=False, unit="index", rationale="Ratio of peak to free flow travel time"),
        "throughput_vph": N(lo=0, fence=False, unit="veh/h", rationale="Cannot be negative"),
        "left_turn_volume": N(lo=0, fence=False, unit="veh/h", rationale="Cannot be negative"),
        "right_turn_volume": N(lo=0, fence=False, unit="veh/h", rationale="Cannot be negative"),
        "through_volume": N(lo=0, fence=False, unit="veh/h", rationale="Cannot be negative"),
        "pedestrian_wait_time_sec": N(lo=0, unit="seconds", rationale="Cannot be negative"),
    },
    impute={
        "speed_limit_mph": ("median", "road_type"), "vehicle_count": ("median", "corridor_name"),
        "avg_speed_mph": ("median", "corridor_name"), "avg_delay_sec_per_vehicle": ("median", "corridor_name"),
        "queue_length_vehicles": ("median", "corridor_name"), "congestion_cost_usd": ("median", "corridor_name"),
        "pedestrian_wait_time_sec": ("median", "corridor_name"),
        "observation_date": ("leave", "observation dates are never imputed; record excluded from dated views only"),
    },
)

FEATURES = [
    ("node_id", "Analysis node = corridor and operating jurisdiction (40 nodes; see C8 rule 6)", "=corridor_name&\" | \"&jurisdiction"),
    ("period_band", "AM peak (07:00 to 09:00), PM peak (15:00 to 18:00) or Off peak, from the clock band", "=IF(time_of_day=\"07:00-09:00\",\"AM peak\",IF(time_of_day=\"15:00-18:00\",\"PM peak\",\"Off peak\"))"),
    ("is_peak", "TRUE for AM and PM peak bands (replaces the inconsistent peak_hour_flag)", "=period_band<>\"Off peak\""),
    ("approach_volume_vph", "Through plus left plus right turning volume", "=through_volume+left_turn_volume+right_turn_volume"),
    ("side_volume_vph", "Turning and minor movements served by the side phase = left plus right", "=left_turn_volume+right_turn_volume"),
    ("clearance_sec", "Change plus clearance interval per phase = yellow + all red", "=yellow_time_sec+all_red_time_sec"),
    ("timing_feasible_flag", "Main green + side green + two clearance intervals fits within the cycle", "=green_main+green_side+2*clearance<=cycle"),
    ("los_hcm", "Level of service recomputed from delay with HCM signalised thresholds (10, 20, 35, 55, 80 s)", "=LOOKUP(delay,{0,10.0001,20.0001,35.0001,55.0001,80.0001},{\"A\",\"B\",\"C\",\"D\",\"E\",\"F\"})"),
    ("obs_weekday", "Weekday derived from the observation date", "=TEXT(observation_date,\"dddd\")"),
    ("obs_month", "First day of the observation month", "=DATE(YEAR(d),MONTH(d),1)"),
]


def los(d):
    if pd.isna(d):
        return None
    for b, l in LOS_BREAKS:
        if d <= b:
            return l
    return "F"


def run():
    raw = pd.read_excel(SRC, sheet_name="Dataset", dtype=object)
    dd = pd.read_excel(SRC, sheet_name="Data Dictionary", header=None, dtype=object)
    cl = Cleaner(raw, CFG)
    cl.profile(); cl.structural(); cl.key_collisions(); cl.text(); cl.booleans(); cl.dates(); cl.numerics()
    checks = [
        dict(Check="Timing plan does not fit its cycle", Description="green main + green side + 2 x (yellow + all red) greater than cycle length",
             mask=lambda d: d.green_time_main_sec + d.green_time_side_sec + 2 * (d.yellow_time_sec + d.all_red_time_sec) > d.cycle_length_sec, flag="timing_infeasible_flag",
             Decision="Cycle length is authoritative (it is the controller setting). For modelling, effective greens are rescaled to fill the cycle less lost time while keeping the recorded main to side split."),
        dict(Check="Recorded LOS contradicts recorded delay", Description="level_of_service differs from the HCM letter implied by avg_delay_sec_per_vehicle",
             mask=lambda d: d.level_of_service != d.avg_delay_sec_per_vehicle.map(los), flag="los_conflict_flag",
             Decision="Delay is the measured quantity and LOS is a label derived from it. los_hcm is recomputed from delay using HCM thresholds and replaces level_of_service."),
        dict(Check="Peak flag contradicts time band", Description="peak_hour_flag TRUE outside the 07:00 to 09:00 and 15:00 to 18:00 bands, or FALSE inside them",
             mask=lambda d: d.peak_hour_flag != d.time_of_day.isin(list(PEAK)),
             Decision="The clock band is objective. period_band and is_peak are derived from time_of_day and replace peak_hour_flag."),
        dict(Check="Day of week contradicts date", Description="day_of_week differs from the weekday of observation_date",
             mask=lambda d: d.observation_date.notna() & (d.day_of_week != d.observation_date.dt.day_name()),
             Decision="The date is authoritative; obs_weekday is derived from it. day_of_week is retained for audit only."),
        dict(Check="Movement volumes disagree with throughput", Description="|through + left + right minus throughput_vph| greater than 10 percent of throughput",
             mask=lambda d: ((d.through_volume + d.left_turn_volume + d.right_turn_volume) - d.throughput_vph).abs() > 0.1 * d.throughput_vph,
             Decision="Movement volumes are needed phase by phase and are authoritative for delay modelling; throughput_vph is used descriptively only."),
        dict(Check="Intersection identifier never repeats", Description="intersection_id appears on only one record",
             mask=lambda d: ~d.intersection_id.duplicated(keep=False),
             Decision="A physical intersection cannot be reconstructed from the ID. The 40 analysis nodes required by the brief are defined as corridor x operating jurisdiction (10 x 4), the unit at which retiming is commissioned."),
    ]
    cl.cross_field(checks)
    cl.missing()
    df = cl.df
    df["node_id"] = df.corridor_name + " | " + df.jurisdiction
    df["period_band"] = df.time_of_day.map(PEAK).fillna("Off peak")
    df["is_peak"] = df.period_band != "Off peak"
    df["approach_volume_vph"] = df.through_volume + df.left_turn_volume + df.right_turn_volume
    df["side_volume_vph"] = df.left_turn_volume + df.right_turn_volume
    df["clearance_sec"] = df.yellow_time_sec + df.all_red_time_sec
    df["timing_feasible_flag"] = ~df.timing_infeasible_flag
    df["los_hcm"] = df.avg_delay_sec_per_vehicle.map(los)
    df["obs_weekday"] = df.observation_date.dt.day_name()
    df["obs_month"] = df.observation_date.dt.to_period("M").dt.to_timestamp()
    cl._log("S10", "C10 Feature Engineering", "Derived analysis nodes, peak bands, volumes and HCM level of service", len(df), len(df), len(FEATURES) * len(df), f"{len(FEATURES)} features added")
    cl._log("S11", "Clean Data", "Published analysis ready table", len(df), len(df), 0, "Clean Data sheet, with named ranges for model formulas")
    cl.df = df
    return raw, dd, cl


if __name__ == "__main__":
    raw, dd, cl = run()
    d = cl.df
    print(d.shape)
    print(cl.cross[["Check", "RecordsFailing", "FailRate"]])
    print(cl.numeric_rules[["Column", "AppliedMax", "BelowMin", "AboveMax"]].to_string())
    for c in ["corridor_name", "jurisdiction", "time_of_day", "road_type", "baseline_or_optimized", "period_band"]:
        print(c, d[c].value_counts().to_dict())
    pk = d[d.is_peak & (d.baseline_or_optimized == "Baseline")]
    print("peak baseline per node", pk.groupby("node_id").size().describe())
    g = pk.groupby("node_id").agg(X=("v_c_ratio", "mean"), C=("cycle_length_sec", "mean"), gm=("green_time_main_sec", "mean"),
                                  gs=("green_time_side_sec", "mean"), cl=("clearance_sec", "mean"), vt=("through_volume", "mean"),
                                  vs=("side_volume_vph", "mean"), dly=("avg_delay_sec_per_vehicle", "mean"))
    print(g.describe().T)
    print(d.groupby("baseline_or_optimized").avg_delay_sec_per_vehicle.mean())
