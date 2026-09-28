# Methodology, Project 3

## Network definition
40 analysis nodes = 10 corridors x 4 operating jurisdictions, because intersection_id never repeats. Each node is calibrated from peak band (07:00 to 09:00 and 15:00 to 18:00) Baseline observations.

## Flow ratios and effective green
Effective greens fill the cycle less lost time (L = 2 x (yellow + all red)) in the recorded main to side ratio. The main phase flow ratio is y = X g / C from the observed degree of saturation; the side phase is scaled by observed volumes and planning lane counts. Implied saturation flows s = v / y close the system.

## Webster optimisation
C0 = (1.5 L + 5) / (1 minus Y), bounded to 60 to 150 seconds, rounded up to 5 seconds, set to the maximum when Y reaches 0.95. Greens are split in proportion to y.

## Control delay
HCM uniform delay d1 = 0.5 C (1 minus g/C) squared / (1 minus min(1, X) g/C) and incremental delay d2 = 900 T [(X minus 1) + sqrt((X minus 1) squared + 8 k I X / (c T))], T = 0.25, k = 0.5, I = 1. Webster's three term delay is shown as a cross check for X below 1.

## Pivot projection and economics
Projected delay = observed delay x model delay (scenario) / model delay (existing, current volumes). Annual cost = vehicle hours per peak hour x peak hours per year x occupancy x value of time; value of time is calibrated to the 14 million dollar brief or entered directly.

## Validation
Six benchmark cases and all 40 nodes are compared with an independent Python implementation (src/p3_ref.py).
