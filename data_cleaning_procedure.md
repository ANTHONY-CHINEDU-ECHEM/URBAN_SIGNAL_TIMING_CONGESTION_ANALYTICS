# Data Cleaning Procedure, Project 3

This document mirrors the eleven cleaning sheets (C1 to C11) in the workbook. Every figure below was produced by the same Python engine that wrote those sheets, and each sheet repeats the key counts as live Excel formulas so the reconciliation can be checked inside the workbook.

## Guiding principles

1. No row is deleted unless it is completely empty or an exact copy of an earlier row.
2. No value is changed silently: every changed cell is traceable through a flag or audit column (key_repaired_flag, outlier_fields, imputed_fields, and the cross field flags).
3. Where two fields contradict each other, one is declared authoritative, the decision is written down, and the other is kept for audit.
4. Dates, identifiers, survey answers and legacy model outputs are never imputed.

## Step summary

<table>
  <tr><th>Step</th><th>Sheet</th><th>Action</th><th>Rows in</th><th>Rows out</th><th>Cells changed</th></tr>
  <tr><td>S1</td><td>C2 Structural Integrity</td><td>Removed fully blank rows</td><td>15,394</td><td>15,372</td><td>0</td></tr>
  <tr><td>S2</td><td>C2 Structural Integrity</td><td>Removed exact duplicate records</td><td>15,372</td><td>15,250</td><td>0</td></tr>
  <tr><td>S3</td><td>C3 Key Collision Repair</td><td>Repaired duplicate record_id values</td><td>15,250</td><td>15,250</td><td>90</td></tr>
  <tr><td>S4</td><td>C4 Text Standardization</td><td>Collapsed categorical variants to one standard label</td><td>15,250</td><td>15,250</td><td>3,001</td></tr>
  <tr><td>S5</td><td>C5 Boolean Normalization</td><td>Converted 12 token encodings to TRUE/FALSE</td><td>15,250</td><td>15,250</td><td>106,750</td></tr>
  <tr><td>S6</td><td>C6 Date Standardization</td><td>Parsed four mixed date formats to true Excel dates</td><td>15,250</td><td>15,250</td><td>15,067</td></tr>
  <tr><td>S7</td><td>C7 Numeric Validation</td><td>Nullified impossible and extreme values</td><td>15,250</td><td>15,250</td><td>546</td></tr>
  <tr><td>S8</td><td>C8 Cross Field Consistency</td><td>Tested business rules between related fields</td><td>15,250</td><td>15,250</td><td>0</td></tr>
  <tr><td>S9</td><td>C9 Missing Value Treatment</td><td>Filled or explicitly retained every missing value</td><td>15,250</td><td>15,250</td><td>2,071</td></tr>
  <tr><td>S10</td><td>C10 Feature Engineering</td><td>Derived analysis nodes, peak bands, volumes and HCM level of service</td><td>15,250</td><td>15,250</td><td>152,500</td></tr>
  <tr><td>S11</td><td>Clean Data</td><td>Published analysis ready table</td><td>15,250</td><td>15,250</td><td>0</td></tr>
</table>

## C1 Data profile

The raw export holds 15,394 rows and 46 columns. Each column was measured for blanks, distinct values and distinct values after normalising case, whitespace, trailing periods and underscores. A gap between those two distinct counts proves the column carries formatting variants.

## C2 Structural integrity

22 rows were completely blank and were removed. 122 rows were exact copies of an earlier row and were removed, keeping the first occurrence. The original Excel row number of every removed row is listed on the sheet.

## C3 Key collision repair

After deduplication, 90 values of record_id were still shared by different records. The collision free anchor sensor_camera_id carries the true record sequence, so 90 records were re keyed from it instead of being deleted. The live check on the sheet confirms that no key in Clean Data is duplicated.

## C4 Text standardisation

9 categorical columns were standardised. 202 non standard spellings covering 3,001 cells were mapped to one governed label. Each raw value was normalised (trim, drop trailing period, underscores to spaces, collapse spaces, ignore case) and mapped to the most frequent clean spelling. Identifier columns were trimmed and upper cased.

<table>
  <tr><th>Issue type</th><th>Variants</th><th>Cells</th></tr>
  <tr><td>Inconsistent letter case</td><td>80</td><td>1,536</td></tr>
  <tr><td>Leading or trailing whitespace</td><td>51</td><td>687</td></tr>
  <tr><td>Trailing period</td><td>51</td><td>551</td></tr>
  <tr><td>Underscore used as separator</td><td>20</td><td>227</td></tr>
</table>

## C5 Boolean normalisation

7 flag columns used twelve encodings of yes and no (true, TRUE, yes, Yes, Y, 1 and their negatives). They were converted to native TRUE and FALSE. No unrecognised tokens were found.

## C6 Date standardisation

Dates arrived as text in four formats. Unambiguous patterns were parsed directly; slash dates where both parts are 12 or below were resolved as month first (the convention shown in the data dictionary) and flagged in an _ambiguous_flag column.

<table>
  <tr><th>Column</th><th>Detected pattern</th><th>Records</th><th>Parsed</th></tr>
  <tr><td>observation_date</td><td>ISO 8601 (YYYY MM DD)</td><td>3,062</td><td>3,062</td></tr>
  <tr><td>observation_date</td><td>Slash ISO (YYYY/MM/DD)</td><td>3,041</td><td>3,041</td></tr>
  <tr><td>observation_date</td><td>Day Month abbrev (DD Mon YYYY)</td><td>3,036</td><td>3,036</td></tr>
  <tr><td>observation_date</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,337</td><td>2,337</td></tr>
  <tr><td>observation_date</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,799</td><td>1,799</td></tr>
  <tr><td>observation_date</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,792</td><td>1,792</td></tr>
  <tr><td>observation_date</td><td>Missing</td><td>183</td><td>0</td></tr>
</table>

## C7 Numeric validation

Each numeric column was tested against physical or definitional limits and, where no hard maximum exists, against the Tukey outer fence (Q3 plus 3 x IQR). Failing cells were blanked (never the whole row) and then treated in C9.

<table>
  <tr><th>Column</th><th>Unit</th><th>Minimum</th><th>Applied maximum</th><th>Below minimum</th><th>Above maximum</th><th>Rationale</th></tr>
  <tr><td>speed_limit_mph</td><td>mph</td><td>15.00</td><td>75.00</td><td>0</td><td>0</td><td>Posted urban and suburban limits</td></tr>
  <tr><td>vehicle_count</td><td>vehicles</td><td>0.00</td><td>1,288.25</td><td>69</td><td>79</td><td>Counts cannot be negative; extreme counts above the fence removed</td></tr>
  <tr><td>avg_speed_mph</td><td>mph</td><td>0.00</td><td>73.76</td><td>85</td><td>65</td><td>Speeds cannot be negative; above 85 mph is a detector error</td></tr>
  <tr><td>cycle_length_sec</td><td>seconds</td><td>30.00</td><td>240.00</td><td>0</td><td>0</td><td>Practical signal cycle range</td></tr>
  <tr><td>green_time_main_sec</td><td>seconds</td><td>1.00</td><td>240.00</td><td>0</td><td>0</td><td>Must be positive and within a cycle</td></tr>
  <tr><td>green_time_side_sec</td><td>seconds</td><td>1.00</td><td>240.00</td><td>0</td><td>0</td><td>Must be positive and within a cycle</td></tr>
  <tr><td>yellow_time_sec</td><td>seconds</td><td>3.00</td><td>6.00</td><td>0</td><td>0</td><td>Standard yellow change interval range</td></tr>
  <tr><td>all_red_time_sec</td><td>seconds</td><td>0.00</td><td>6.00</td><td>0</td><td>0</td><td>Red clearance range</td></tr>
  <tr><td>num_phases</td><td>phases</td><td>2.00</td><td>8.00</td><td>0</td><td>0</td><td>Dual ring controller maximum</td></tr>
  <tr><td>avg_delay_sec_per_vehicle</td><td>s/veh</td><td>0.00</td><td>102.89</td><td>80</td><td>68</td><td>Delay cannot be negative; sentinel 128 caught by fence</td></tr>
  <tr><td>queue_length_vehicles</td><td>vehicles</td><td>0.00</td><td>49.93</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>v_c_ratio</td><td>ratio</td><td>0.00</td><td>2.00</td><td>0</td><td>0</td><td>Degree of saturation, oversaturation allowed to 2.0</td></tr>
  <tr><td>fuel_wasted_gallons</td><td>gallons</td><td>0.00</td><td>112.27</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>co2_emissions_kg</td><td>kg</td><td>0.00</td><td>335.77</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>congestion_cost_usd</td><td>USD</td><td>0.00</td><td>1,131.08</td><td>0</td><td>100</td><td>Magnitude errors above the fence removed</td></tr>
  <tr><td>pedestrian_volume</td><td>persons</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Count</td></tr>
  <tr><td>bike_volume</td><td>bikes</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Count</td></tr>
  <tr><td>travel_time_index</td><td>index</td><td>0.50</td><td>5.00</td><td>0</td><td>0</td><td>Ratio of peak to free flow travel time</td></tr>
  <tr><td>throughput_vph</td><td>veh/h</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>left_turn_volume</td><td>veh/h</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>right_turn_volume</td><td>veh/h</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>through_volume</td><td>veh/h</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>pedestrian_wait_time_sec</td><td>seconds</td><td>0.00</td><td>124.54</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
</table>

## C8 Cross field consistency

<table>
  <tr><th>Rule</th><th>Records failing</th><th>Fail rate</th><th>Decision</th></tr>
  <tr><td>Timing plan does not fit its cycle</td><td>6,951</td><td>45.6 percent</td><td>Cycle length is authoritative (it is the controller setting). For modelling, effective greens are rescaled to fill the cycle less lost time while keeping the recorded main to side split.</td></tr>
  <tr><td>Recorded LOS contradicts recorded delay</td><td>12,706</td><td>83.3 percent</td><td>Delay is the measured quantity and LOS is a label derived from it. los_hcm is recomputed from delay using HCM thresholds and replaces level_of_service.</td></tr>
  <tr><td>Peak flag contradicts time band</td><td>5,726</td><td>37.5 percent</td><td>The clock band is objective. period_band and is_peak are derived from time_of_day and replace peak_hour_flag.</td></tr>
  <tr><td>Day of week contradicts date</td><td>12,971</td><td>85.1 percent</td><td>The date is authoritative; obs_weekday is derived from it. day_of_week is retained for audit only.</td></tr>
  <tr><td>Movement volumes disagree with throughput</td><td>13,795</td><td>90.5 percent</td><td>Movement volumes are needed phase by phase and are authoritative for delay modelling; throughput_vph is used descriptively only.</td></tr>
  <tr><td>Intersection identifier never repeats</td><td>15,250</td><td>100.0 percent</td><td>A physical intersection cannot be reconstructed from the ID. The 40 analysis nodes required by the brief are defined as corridor x operating jurisdiction (10 x 4), the unit at which retiming is commissioned.</td></tr>
</table>

## C9 Missing value treatment

<table>
  <tr><th>Column</th><th>Blanks before</th><th>Strategy</th><th>Filled</th><th>Blanks after</th></tr>
  <tr><td>speed_limit_mph</td><td>152</td><td>Median within road_type (overall median if group empty)</td><td>152</td><td>0</td></tr>
  <tr><td>vehicle_count</td><td>377</td><td>Median within corridor_name (overall median if group empty)</td><td>377</td><td>0</td></tr>
  <tr><td>avg_speed_mph</td><td>379</td><td>Median within corridor_name (overall median if group empty)</td><td>379</td><td>0</td></tr>
  <tr><td>avg_delay_sec_per_vehicle</td><td>453</td><td>Median within corridor_name (overall median if group empty)</td><td>453</td><td>0</td></tr>
  <tr><td>queue_length_vehicles</td><td>305</td><td>Median within corridor_name (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>congestion_cost_usd</td><td>100</td><td>Median within corridor_name (overall median if group empty)</td><td>100</td><td>0</td></tr>
  <tr><td>pedestrian_wait_time_sec</td><td>305</td><td>Median within corridor_name (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>observation_date</td><td>183</td><td>Left blank: observation dates are never imputed; record excluded from dated views only</td><td>0</td><td>183</td></tr>
</table>

## C10 Feature engineering

Derived fields are documented on the C10 sheet with their definition, worksheet equivalent and a live populated count.

## C11 Audit log and reconciliation

The final sheet lists every step with rows in, rows out and cells changed, and reconciles live: rows in Raw Data, less blank rows, less exact duplicates, must equal rows in Clean Data. The check returns PASS in the delivered workbook.
