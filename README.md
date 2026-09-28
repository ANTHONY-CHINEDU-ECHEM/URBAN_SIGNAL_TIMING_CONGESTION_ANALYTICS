# Urban Signal Timing and Congestion Cost Simulator

An Excel traffic engineering model that optimises signal timing across a 40 intersection arterial network using Webster's method, computes control delay with the Highway Capacity Manual (HCM) equations, values the result in dollars, and tests how robust the benefits are to future traffic growth. The model is built from a supplied 15,394 row observation extract through a fully documented, eleven step cleaning procedure, and it runs entirely in native Excel with no add ins, macros or external services.

## Project brief

Traffic signals are among the most influential and least visible pieces of public infrastructure a city owns. Every second of green time given to one approach is a second taken from another, and the timing plans programmed into signal controllers quietly decide how long thousands of drivers, bus passengers and freight operators wait every day. When those plans are well matched to demand, a corridor moves smoothly; when they are not, queues build, vehicles stop repeatedly, journeys lengthen, fuel is wasted and emissions rise. Because the cost of poor timing is spread thinly across many travellers, it rarely appears as a single line in any budget, yet in aggregate it can be very large.

The city in this study operates a 40 intersection arterial corridor where commuters experience an average of 18 minutes of peak hour delay, a burden the city estimates at 14 million dollars a year. The timing plans on the corridor have not been updated since 2016. In the years since, traffic volumes, land use, travel patterns and the balance between main road and side street demand have all shifted, but the signals continue to run plans designed for conditions that no longer exist. Without a structured way to calculate optimal timings and quantify their value, the agency has had no clear basis for deciding which intersections to retime first, how much benefit to expect, or whether retiming would still pay off as traffic continues to grow.

This project was commissioned to answer those questions with engineering rigour and in financial terms. The brief called for a model that calibrates every intersection from observed data, computes optimal cycle lengths and green splits with Webster's method, calculates delay per vehicle with the recognised HCM uniform and incremental delay equations, and translates vehicle hours of delay into annual economic cost using a value of time methodology. It also asked for a before and after corridor dashboard, growth sensitivity scenarios at 5, 10 and 20 percent additional volume, and validation of the delay calculations against an independent reference. Most traffic analyses in a portfolio setting stop at a dashboard of observed conditions; this one computes the optimal timing itself and shows what it is worth.

## Objectives

* Calibrate all 40 intersections from peak period baseline observations.
* Compute optimal cycle lengths and green splits for every node using Webster's method, with the option for engineers to override any timing.
* Calculate control delay per vehicle with the HCM uniform and incremental delay equations for both existing and optimised plans.
* Convert delay savings into annual economic value using vehicle occupancy and value of time.
* Test the benefit under current volumes and under 5, 10 and 20 percent traffic growth, and show the remaining capacity headroom at each node.
* Validate the Excel delay calculations against an independent Python implementation of the HCM equations.

## What makes this model different

This workbook does not simply report the delay that was observed. It rebuilds each intersection as a queuing system, derives critical flow ratios from the definition of degree of saturation, calculates the optimal cycle with Webster's formula, and then recomputes delay for both the existing and the optimised plans with the same HCM equations, so the comparison is like for like. A pivot point method anchors the modelled change to observed delay, and the economic layer converts every second saved into dollars. The full engine of 200 rows recalculates in under two seconds, so engineers and managers can test overrides and growth scenarios interactively.

## Data

The supplied extract contains 15,394 observation rows covering intersection volumes, cycle lengths, green times, lane counts, observed delay, level of service letters and a baseline or optimised label by corridor, jurisdiction and time period. Cleaning followed eleven documented steps, one per worksheet (C1 to C11).

Several data issues shaped the analysis and are worth stating plainly:

* intersection_id never repeats, so no single intersection can be tracked across records. The 40 intersections are therefore represented by 40 analysis nodes: 10 corridors by 4 operating jurisdictions.
* Some recorded green times add up to more than the cycle length, which is physically impossible. These were rescaled proportionally to fit the cycle.
* Recorded level of service letters contradict recorded delay on 83 percent of records, even though level of service is defined by delay thresholds. The letters were treated as unreliable and level of service was recalculated from delay.
* Records labelled as previously optimised show no delay improvement over baseline records, so they could not be used as evidence of what retiming achieves.

## Results at a glance

<table>
  <tr><th>Measure</th><th>Current volumes</th><th>Growth of 20 percent</th></tr>
  <tr><td>Delay per vehicle, existing plans</td><td>25.9 s</td><td>46.0 s</td></tr>
  <tr><td>Delay per vehicle, optimised plans</td><td>22.2 s</td><td>36.3 s</td></tr>
  <tr><td>Delay reduction</td><td>14.5 percent</td><td>21.0 percent</td></tr>
  <tr><td>Annual saving</td><td>2.0 million dollars</td><td>6.3 million dollars</td></tr>
  <tr><td>Nodes where retiming gives a clear gain</td><td>38 of 40</td><td></td></tr>
</table>

<img width="1092" height="345" alt="Screenshot 2026-09-28 at 23 27 35" src="https://github.com/user-attachments/assets/0419f44f-ecf9-4fa3-a3d2-e1a22bf6aca7" />


## Success metrics

<table>
  <tr><th>Target</th><th>Result</th><th>Status</th></tr>
  <tr><td>Delay formulas within 5 percent of HCM reference values</td><td>46 of 46 checks match an independent Python implementation of the HCM equations</td><td>Met</td></tr>
  <tr><td>Corridor delay reduction of 15 to 22 percent</td><td>14.5 percent at current volumes; 16.3 and 21.0 percent at 10 and 20 percent growth</td><td>Just below target at current volumes, within range under growth</td></tr>
  <tr><td>Full 40 intersection coverage</td><td>40 of 40 nodes calibrated</td><td>Met</td></tr>
  <tr><td>Recalculation under 5 seconds</td><td>Under 2 seconds per selector change, measured</td><td>Met</td></tr>
</table>

## Findings in detail

### 1. Retiming delivers an immediate, broad based benefit

Replacing the 2016 plans with Webster optimised timings reduces average control delay from 25.9 to 22.2 seconds per vehicle, a cut of 14.5 percent at current volumes, worth about 2.0 million dollars a year. The benefit is not concentrated in a few problem junctions: 38 of the 40 nodes show a clear gain. For a transport agency, signal retiming is typically one of the lowest cost interventions available, because it changes controller settings rather than physical infrastructure. A saving of this scale from a software change is a strong case for acting on the whole corridor rather than piloting a handful of sites.

<img width="1092" height="345" alt="Screenshot 2026-09-28 at 23 27 35" src="https://github.com/user-attachments/assets/6867210b-6465-4f79-92c6-cb5ac8e8d492" />

### 2. The value of retiming grows as traffic grows

Delay at signals rises sharply, not steadily, as intersections approach capacity. Under the existing plans, a 20 percent rise in traffic pushes delay from 25.9 to 46.0 seconds per vehicle, an increase of almost 80 percent. Optimised plans absorb the same growth far better, holding delay to 36.3 seconds. The delay reduction therefore widens from 14.5 percent today to 16.3 percent at 10 percent growth and 21.0 percent at 20 percent growth, and the annual saving rises from 2.0 million to 6.3 million dollars. Retiming is best understood not only as a fix for today's congestion but as insurance against tomorrow's, and every year the old plans remain in place the cost of inaction increases.

<img width="754" height="352" alt="Screenshot 2026-09-28 at 23 30 32" src="https://github.com/user-attachments/assets/7981b005-2d2d-43a0-9021-1e4ada6f82bf" />

### 3. The corridor target is essentially met, and exceeded under realistic growth

The brief targeted a delay reduction of 15 to 22 percent. At current volumes the model delivers 14.5 percent, half a point short of the lower bound, and under 10 and 20 percent growth it delivers 16.3 and 21.0 percent, comfortably within the range. Given that the model represents each intersection conservatively, without coordination between signals, the current volume result is likely a floor rather than a ceiling. Coordinated timing with optimised offsets along the corridor would be expected to add further benefit on top of the isolated intersection gains shown here.

<img width="705" height="381" alt="Screenshot 2026-09-28 at 23 31 31" src="https://github.com/user-attachments/assets/b9823db0-973f-4bf3-a8b7-2e7746892dce" />

### 4. The 14 million dollar estimate needs a clearer cost basis

To reproduce the brief's figure of 14 million dollars a year from the delay in the extract, the model would require a value of time of about 32.5 dollars per person hour. That is high for personal travel alone, which suggests the city's estimate also includes other costs such as fuel, freight time, emissions or off peak delay. This matters for decision making: savings should be quoted on a basis the finance team and funding bodies will accept. The workbook allows the cost basis to be switched to an agency value of time, so the savings can be restated on whichever basis the city adopts.

<img width="727" height="335" alt="Screenshot 2026-09-28 at 23 33 31" src="https://github.com/user-attachments/assets/92af13b7-49f2-41e5-bc10-3c61f214adbc" />

### 5. The brief's 18 minute figure measures something different

The 18 minutes of peak hour delay quoted in the brief is an end to end trip delay per commuter, accumulated across the whole journey. The extract records control delay at individual intersections, measured in seconds per vehicle. The two figures are not directly comparable, and this model deliberately does not try to reconcile them. Quantifying the full journey experience would require travel time survey or probe vehicle data.

<img width="1050" height="454" alt="Screenshot 2026-09-28 at 23 35 11" src="https://github.com/user-attachments/assets/68f35b80-8596-438d-8bd6-0547a1810ae6" />

### 6. Existing records overstate the quality of the data

Two data findings should concern the agency as much as the delay results. Level of service letters contradict recorded delay on 83 percent of records, and records labelled as previously optimised show no improvement at all over the baseline. Either past optimisation work was never implemented as intended, or the labels and grades in the data are unreliable. In both cases, performance reports built on these fields would have painted a misleading picture of corridor health, and this should be investigated before the fields are used again.

<img width="1219" height="340" alt="Screenshot 2026-09-28 at 23 36 30" src="https://github.com/user-attachments/assets/436e5a96-3baf-47c2-8797-f785939410ca" />

### 7. The calculations are verified and fast enough for daily use

All 46 validation checks of the Excel delay calculations match an independent Python implementation of the HCM equations, well within the 5 percent tolerance set in the brief. Every one of the 40 nodes is calibrated, and each selector change recalculates in under two seconds. The model is therefore reliable enough to support prioritisation and funding decisions, and responsive enough to be used live in engineering and management meetings.

## Recommendations

1. Approve a corridor wide retiming programme covering the 38 nodes where the model shows a clear gain, rather than a limited pilot.
2. Prioritise nodes with the least capacity headroom under growth, since they will deteriorate fastest if left on the 2016 plans.
3. Commission turning movement counts at priority intersections so that timings can move from planning estimates to controller ready plans.
4. Follow isolated retiming with a coordination study to optimise offsets along the corridor for additional benefit.
5. Agree an official value of time and cost basis with finance before publishing savings figures.
6. Investigate why previously optimised records show no benefit and why level of service grades contradict delay, and correct the data feed.
7. Put retiming on a regular review cycle so plans are refreshed as demand changes, instead of waiting another decade.

## Method

1. **Clean the extract.** Eleven documented steps (see docs/data_cleaning_procedure.md). Because intersection_id never repeats, the 40 intersections are represented by 40 analysis nodes: 10 corridors by 4 operating jurisdictions.
2. **Calibrate each node.** Nodes are calibrated from peak period baseline observations. Recorded greens that overrun the cycle are rescaled to fit it.
3. **Derive flow ratios.** Critical flow ratios for the main phase come from the definition of degree of saturation (y = X g / C), and the side phase is scaled from observed volumes.
4. **Optimise timing.** The cycle is optimised with Webster's formula, C0 = (1.5 L + 5) / (1 minus Y), bounded to practical limits and rounded, and greens are split in proportion to flow ratios.
5. **Compute delay and cost.** Control delay is computed with the HCM uniform and incremental delay terms, projected onto observed delay with the pivot point method, and vehicle hours are valued using occupancy and value of time.
6. **Test and validate.** Results are recalculated under 5, 10 and 20 percent growth and a custom growth rate, and the delay engine is checked against an independent Python reference implementation.

## Workbook structure

<table>
  <tr><th>Sheet</th><th>Purpose</th></tr>
  <tr><td>01 Executive Brief</td><td>Cost savings summary, success metrics and recommendation</td></tr>
  <tr><td>02 Corridor Dashboard</td><td>Before and after comparison with corridor, growth and peak period selectors</td></tr>
  <tr><td>03 Signal Timing Optimizer</td><td>Existing versus Webster timing for all 40 nodes, with engineer overrides</td></tr>
  <tr><td>04 Delay and Cost Calculator</td><td>Step by step HCM and Webster delay through to annual cost for one intersection</td></tr>
  <tr><td>05 Growth Sensitivity</td><td>Current volumes, 5, 10 and 20 percent growth and a custom rate, with node headroom</td></tr>
  <tr><td>06 Methodology</td><td>Equations, data decisions and limitations</td></tr>
  <tr><td>07 HCM Validation</td><td>Excel versus independent reference, success metrics and assumptions</td></tr>
  <tr><td>M1 and M2</td><td>Node calibration with all inputs, and the 200 row Webster and HCM engine</td></tr>
  <tr><td>C1 to C11</td><td>Data cleaning procedure, one step per sheet</td></tr>
</table>

## How to use

Open the workbook in Microsoft Excel 2010 or later; calculation is automatic.

* Use the teal bordered selectors on sheet 02 to choose corridor, growth scenario and peak period, and on sheet 04 to step through the delay and cost calculation for any intersection.
* Enter engineer overrides for cycle length or green splits on sheet 03.
* Adjust engineering and economic inputs, such as saturation flow, occupancy and value of time, on sheet M1.

The selectors are in cell drop down lists (data validation), which behave like combo boxes in every Excel version. A Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

The workbook is generated by Python, so every sheet can be rebuilt from the raw data.

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

## Assumptions and limitations

* **Two phase representation.** Each node is modelled with two critical phases and without coordination or offsets between signals, so corridor progression benefits are not captured.
* **Planning inputs.** Lane counts are planning values and peak volumes are averaged, so timings are suitable for prioritisation rather than direct controller programming until turning movement counts are collected.
* **Node representation.** Because intersection_id never repeats, nodes represent corridor and jurisdiction combinations rather than individually surveyed junctions.
* **Trip delay.** The brief's 18 minute figure is an end to end trip delay that the extract does not record, so it is not reproduced by the model.
* **Cost basis.** Matching the 14 million dollar estimate implies a value of time of about 32.5 dollars per person hour; the savings shown depend on the value of time selected on sheet M1.
