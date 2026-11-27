# Methods and implementation notes

## Problem and visualization

The app minimizes Sphere, Rastrigin, or Rosenbrock over 2–100 dimensions. The landscape background is a two-dimensional slice: x₁ and x₂ vary, and all later coordinates are fixed to zero. Candidate markers show the first two coordinates, but each score is calculated across the full vector. The visible surface and a high-dimensional candidate's objective score therefore need not correspond directly.

## Method selector

The lab offers a configurable baseline and 13 named browser profiles. Each generation uses one-to-one greedy selection. Values outside the objective bounds are reflected into the feasible interval. Peer vectors are sampled without replacement and exclude the target where the method permits.

| Method | Main mechanism represented in this app | Profile details |
|---|---|---|
| Configurable DE | User-selected mutation, crossover, F, and CR | Supports rand/1, best/1, current-to-best/1; binomial, exponential, blend, and contiguous-segment crossover. Optional jitter is explicitly experimental. |
| DE/rand/1/bin | Random base plus one scaled difference | Fixed F and binomial crossover. |
| DE/best/1/bin | Best-so-far base plus one scaled difference | Fixed F and binomial crossover. |
| DE/current-to-best/1/bin | Target moves toward current best plus one scaled difference | Fixed F and binomial crossover. |
| DE/rand/2/bin | Random base plus two scaled differences | Fixed F and binomial crossover; needs at least six population members. |
| DE/current-to-rand/1 | Randomized movement toward a peer plus one scaled difference | Uses a random K per trial and no crossover. |
| jDE | Per-individual F and CR adaptation | Successful parameter values persist with their candidate; compact τ-based update profile. |
| JADE | Current-to-pbest/1 with an external archive | Adaptive Cauchy F and normal CR sampling; simplified running means. |
| SHADE | Current-to-pbest/1 with archive and success-history memories | Random memory slot, weighted Lehmer mean for F, weighted arithmetic mean for CR. |
| L-SHADE | SHADE-style profile plus linear population-size reduction | Population is reduced toward four as the evaluation budget is consumed. |
| jSO | L-SHADE-family p-best search with staged weighting | Compact educational profile with simplified stage thresholds and parameter weighting; not the full CEC jSO specification. |
| SaDE | Learns selection probabilities for four mutation strategies | Recent success/failure counts update probabilities; this compact profile is not a bit-exact original SaDE port. |
| CoDE | Composite strategy and parameter pool | Tries up to three configured trial vectors per target and lets the best compete; each trial consumes one evaluation. |
| EPSDE | Per-candidate strategy and parameter choices | Failed choices are reinitialized from small built-in pools; simplified educational profile. |

The adaptive profiles own their F/CR controls; fixed controls are disabled when they do not apply. F and CR are bounded by 0 and 1 in the UI and the method samplers. For some classic variants, the labels indicate the crossover method as part of the named algorithm; those variants always use binomial crossover regardless of the configurable baseline's last selection.

## Budget and population behavior

The evaluation count includes initial population scoring and every objective function call. CoDE may evaluate fewer than three trials for the final target if the remaining budget is nearly exhausted, and may stop in the middle of a generation. If the requested budget is below the requested initial population, initialization is shortened to the budget. L-SHADE and jSO reduce the population as progress advances; the minimum profile population is four. Use populations at least six for DE/rand/2/bin and CoDE.

The visualizer accepts dimensions up to 100, population sizes through 250, and budgets through 10,000. These are interface limits, not recommended experimental settings or validated performance limits. The Data Studio processes up to 25 MB, 50,000 rows, 1,000 fields, and 1,000,000 cells per imported view; oversized input is capped and reported.

## Reproducibility and scientific scope

All browser random choices use the displayed seed. Changing a setting requires a fresh run. The app has no CEC test suite, official boundary protocol, or validated timing/accuracy comparison. Its figure suggestions are structural heuristics, not statistical inference or domain interpretation. Its profiles help explain core ideas; they should not be cited as exact replications or used to claim comparative algorithm performance.

## Primary references

- JADE: [Adaptive Differential Evolution with Optional External Archive](https://ieeexplore.ieee.org/document/5208221/).
- SHADE: [Success-History Based Adaptive Differential Evolution](https://www.metahack.org/CEC2013-SHADE.pdf).
- L-SHADE: [Ryoji Tanabe's publications page and implementation notes](https://ryojitanabe.github.io/publication).
- jSO: [Success-History based Parameter Adaptation for Differential Evolution](https://labraj.feri.um.si/wp-content/uploads/janez/CEC2017-jSO.pdf).
- SaDE: [Differential Evolution with Adaptive Strategy and Parameter Selection](https://tiezhongyu2005.github.io/resources/popularization/SaDE_2009.pdf).
- CoDE and EPSDE: [Ensemble strategies and parameter choices for Differential Evolution](https://www.sciencedirect.com/science/article/pii/S1568494610001043).
