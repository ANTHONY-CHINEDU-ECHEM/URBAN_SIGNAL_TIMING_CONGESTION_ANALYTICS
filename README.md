# Urban Signal Timing and Congestion Cost Simulator

An Excel traffic engineering model that optimises signal timing across a 40 intersection network with Webster's method, computes control delay with the Highway Capacity Manual equations, values the result in dollars, and tests robustness to traffic growth. Built from a supplied 15,394 row observation extract through a fully documented, eleven step cleaning procedure.

Excel Project 3 of 10 in a data analyst portfolio. Prepared by Anthony Chinedu Echem.

## Business problem

A 40 intersection corridor suffers heavy peak delay costing an estimated 14 million dollars a year, and timing plans have not been updated since 2016.

## Results at a glance

<table>
  <tr><th>Measure</th><th>Current volumes</th><th>Growth plus 20 percent</th></tr>
  <tr><td>Delay per vehicle, existing plans</td><td>25.9 s</td><td>46.0 s</td></tr>
  <tr><td>Delay per vehicle, optimised plans</td><td>22.2 s</td><td>36.3 s</td></tr>
  <tr><td>Delay reduction</td><td>14.5 percent</td><td>21.0 percent</td></tr>
  <tr><td>Annual saving</td><td>2.0 million dollars</td><td>6.3 million dollars</td></tr>
  <tr><td>Nodes where retiming gives a clear gain</td><td>38 of 40</td><td></td></tr>
</table>

## Success metrics

<table>
  <tr><th>Catalogue metric</th><th>Result</th><th>Status</th></tr>
  <tr><td>Delay formulas within 5 percent of HCM reference values</td><td>46 of 46 checks match an independent Python implementation of the HCM equations</td><td>Met</td></tr>
  <tr><td>Corridor delay reduction of 15 to 22 percent</td><td>14.5 percent at current volumes; 16.3 and 21.0 percent at 10 and 20 percent growth</td><td>Just below target at current volumes</td></tr>
  <tr><td>Full 40 intersection coverage</td><td>40 of 40 nodes calibrated</td><td>Met</td></tr>
  <tr><td>Recalculation under 5 seconds</td><td>Under 2 seconds per selector change, measured</td><td>Met</td></tr>
</table>

## Workbook structure

<table>
  <tr><th>Sheet</th><th>Purpose</th></tr>
  <tr><td>01 Executive Brief</td><td>Cost savings summary, success metrics, recommendation</td></tr>
  <tr><td>02 Corridor Dashboard</td><td>Before and after comparison with corridor, growth and peak period selectors</td></tr>
  <tr><td>03 Signal Timing Optimizer</td><td>Existing versus Webster timing for all 40 nodes, with engineer overrides</td></tr>
  <tr><td>04 Delay and Cost Calculator</td><td>Step by step HCM and Webster delay through to annual cost for one intersection</td></tr>
  <tr><td>05 Growth Sensitivity</td><td>Current, plus 5, plus 10, plus 20 percent and custom growth, with node headroom</td></tr>
  <tr><td>06 Methodology</td><td>Equations, data decisions and limitations</td></tr>
  <tr><td>07 HCM Validation</td><td>Excel versus independent reference, success metrics, assumptions</td></tr>
  <tr><td>M1 and M2</td><td>Node calibration with all inputs, and the 200 row Webster and HCM engine</td></tr>
  <tr><td>C1 to C11</td><td>Data cleaning procedure, one step per sheet</td></tr>
</table>

## Method

1. Clean the extract (see docs/data_cleaning_procedure.md). intersection_id never repeats, so the 40 intersections are represented by 40 analysis nodes: 10 corridors x 4 operating jurisdictions.
2. Calibrate each node from peak period baseline observations. Recorded greens that overrun the cycle are rescaled to fit it.
3. Derive critical flow ratios from the definition of degree of saturation (y = X g / C) for the main phase and scale the side phase from observed volumes.
4. Optimise the cycle with Webster's formula C0 = (1.5 L + 5) / (1 minus Y), bounded and rounded, and split greens in proportion to flow ratios.
5. Compute control delay with HCM uniform and incremental delay, project it onto observed delay with the pivot point method, and value vehicle hours with occupancy and value of time.

## Key findings

* Retiming cuts peak delay by 14.5 percent now, and the benefit grows with demand, reaching 21.0 percent at 20 percent growth.
* Matching the brief's 14 million dollars requires a value of time of about 32.5 dollars per person hour, high for personal travel alone, so the brief figure probably also includes fuel, freight or off peak delay. The cost basis can be switched to an agency value of time.
* Earlier Optimized records in the extract show no delay improvement over Baseline, and recorded level of service letters contradict recorded delay on 83 percent of records.

## How to use

Open the workbook in Microsoft Excel 2010 or later. Use the teal bordered selectors on 02 and 04, enter engineer overrides on 03 and engineering or economic inputs on M1. The selectors are in cell drop down lists (data validation), which behave like combo boxes in every Excel version; a Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

```
pip install pandas numpy openpyxl scipy
python3 src/p3_build.py
```

Run from the repository root. src/p3_ref.py is the independent reference implementation used for validation.

## Repository structure

```
README.md
workbook/      finished Excel model
data/raw/      supplied dataset, unchanged
src/           cleaning engine, reference model and build script
docs/          data cleaning procedure and methodology
```

## Limitations

Two critical phase representation without coordination or offsets, planning lane counts, and averaged peak volumes. The brief's 18 minute figure is an end to end trip delay that the extract does not record. Results support prioritisation, not controller programming, until turning movement counts are collected.

## Author

Anthony Chinedu Echem
