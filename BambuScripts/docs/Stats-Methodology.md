# How the Efficiency Stats Work - Methodology & Reliability

*A plain-English overview of the Stats tab in the Card Queue Editor: what it measures, where the numbers come from, and how much to trust them.*

---

## What we measure, and why

The headline number is **throughput - sellable objects produced per print-day (obj/day).** It's the bottom-line business metric: how much sellable product a design yields per day of printer time. Everything else (filament use, print time, color changes) matters only insofar as it moves throughput or cost.

## Where the data comes from

We analyzed **298 real production designs** (currently all X1C-Standard). Every number is pulled from each design's own `_Data.tsv` file that the pipeline already produces - print time, object count, filament grams, color changes. **No estimates or hand-entry** - it's the same data the printers and slicer report.

> **Scope note:** these baselines describe **X1C-Standard** designs. Keychains, Big Wigs, or P2S prints will get their own baselines once we've collected enough of each (we need ~25+ per group to trust an average).

## The baselines (the "average" line)

For each variable we calculate the **mean (simple average)** across all 298 designs. We deliberately use the average rather than the median so that genuinely hard designs (our LEGENDARY/EPIC tiers) stay in the picture and pull the bar where it really is.

## How "standard deviation" (SD) is calculated, and what it means

Standard deviation is a standard statistical measure of **spread** - the typical distance a design sits from the average. We compute it the textbook way: take each design's distance from the average, square it, average those, and take the square root.

In plain terms, on the bell-curve view:

- **~2 out of 3 designs** fall within **1 SD** of the average.
- **~19 out of 20** fall within **2 SD**.
- So a design sitting "+2 SD" on color changes is in the worst ~2.5% - a real outlier worth a look.

## The bell curve - what it is and its limits

The bell curve shows **where a design sits relative to all the others.** It's an idealized "normal distribution" drawn from the average and the SD. It's a quick read of *"is this design typical, or unusual?"*

**The honest caveat:** the bell curve assumes the data is roughly symmetric around the average. Most of our metrics are close enough that it's a fair gauge. A few (like color changes) are mildly lopsided, so near the extreme tails the bell is an **approximation, not an exact probability.** We use it as a "how unusual is this" indicator, not a precise statistical claim.

## The part that's exact - the two real levers

This is the most reliable finding and worth emphasizing: throughput is **mathematically determined** by exactly two factors:

> **throughput = 1440 / (filament-per-unit x time-per-gram)**

(The `1440` is just the number of minutes in a day, 24 x 60, which is why throughput comes out *per day*.)

This isn't a model or a correlation - it's an identity, verified exact across all 298 designs. **Filament-per-unit** (how much plastic each sellable object needs) and **time-per-gram** (how fast we print that plastic) are the *only* two things that move throughput, and they're independent of each other. Everything we do to improve throughput acts through one of these two.

## The line graphs and "fit strength" (R-squared)

The line graphs show how each variable *relates* to throughput, with a **fit-strength score (R-squared)** - the share of throughput variation that variable explains:

| Tier | Variables | R-squared | How to use it |
|---|---|---|---|
| **Strong** | Filament / unit, Time / gram | 0.46-0.49 | The true levers. Trustworthy. |
| **Moderate** | Print time, Color changes | 0.38-0.41 | Real trends, but they act *through* time-per-gram. Symptoms, not separate knobs. |
| **Weak** | Objects / plate, Model filament | 0.03-0.14 | Little/no reliable relationship - the tool refuses to draw a confident trend line. |

R-squared runs from 0 (no relationship) to 1 (perfectly explained). An R-squared of 0.49 means that variable alone accounts for ~49% of the differences in throughput between designs.

## Bottom line on reliability

- **Trust completely:** throughput itself, and the two-factor decomposition (it's exact math).
- **Trust as direction:** the correlations / line graphs, read as "where does this design sit and is that good or bad," not as precise predictions.
- **Don't over-read:** the extreme tails of any single bell curve, and the weak-fit variables.
- **Remember the scope:** conclusions apply to the X1C-Standard population we measured; other product/printer types get their own baselines as data comes in.

---

## Quick reference - corpus baselines (n=298, X1C-Standard)

| Variable | Average | SD | Throughput fit (R-squared) |
|---|---|---|---|
| Throughput | 75.1 obj/day | 17.6 | (the metric itself) |
| Filament / unit | 3.52 g | 0.60 | 0.49 (strong) |
| Time / gram | 5.75 min/g | 0.99 | 0.46 (strong) |
| Print time | 30.0 h | 7.5 | 0.41 (moderate) |
| Color changes | 145.3 | 49.0 | 0.38 (moderate) |
| Objects / plate | 90.5 | 18.1 | 0.14 (weak) |
| Model filament | 315.0 g | 67.1 | 0.03 (weak) |

*Generated from `data/production_metrics.csv`. The Efficiency "score" shown per design is the throughput index: design throughput / 75.1 x 100, so 100 = corpus-average throughput.*
