"""
src/optimizer.py
================
Function 2 — Lagrange + Monte Carlo demurrage optimizer.

Objective   : Minimize E[demurrage cost USD] per site.
Method      : L-BFGS-B with Lagrange ordering-constraint penalty.
              Objective = E[cost(w)] + λ · Σ max(0, w[i] − w[i+1])²
              E[cost] is estimated via Monte Carlo simulation.
Constraints : Ships must stay in chronological order (w[i] < w[i+1]).
              Enforced as a quadratic penalty — no hard bounds needed.
Freedom     : No shift limit. No minimum gap. Optimizer decides freely.
Granularity : Optimised independently per site (Gh-1 / SPP3).
"""

import random
import warnings
import numpy as np
from datetime import datetime, timedelta
from scipy.optimize import minimize

warnings.filterwarnings("ignore")

# Window width is fixed by contract design (= num_window in Assumption sheet)
_NW = 10


def _mc_cost(
    win_starts: np.ndarray,
    rate:       float,
    min_dis:    float,
    max_dis:    float,
    allowance:  float,
    n_sim:      int,
    seed:       int,
) -> float:
    """
    Monte Carlo estimate of expected demurrage cost for one site.

    Parameters
    ----------
    win_starts : array of window start offsets in days from BASE_DATE
    rate       : demurrage rate USD/day for this site
    n_sim      : number of Monte Carlo draws
    seed       : RNG seed for reproducibility

    Returns
    -------
    float — E[total demurrage cost USD] over all ships on this site
    """
    rng   = random.Random(seed)
    total = 0.0
    for _ in range(n_sim):
        prev_finish = None
        for ws_day in win_starts:
            # Arrival = uniform draw within [win_start, win_start + window_width]
            arrival   = ws_day + rng.random() * _NW
            nor12     = arrival + 0.5                     # NOR + 12 hours
            # Commence = arrival + small random fraction (up to 12 h)
            # but not before berth is free (prev ship's finish)
            commence  = arrival + rng.random() * 0.5
            if prev_finish is not None and prev_finish > commence:
                commence = prev_finish
            # Actual discharging time drawn uniformly in [min_dis, max_dis]
            dis_days  = rng.random() * (max_dis - min_dis) + min_dis
            finish    = commence + dis_days
            # Allowance clock starts at NOR+12, never from commence
            allowance_end = nor12 + allowance
            dem       = max(0.0, finish - allowance_end)
            total    += dem * rate
            prev_finish = finish
    return total / n_sim


def _lagrangian(
    w:        np.ndarray,
    rate:     float,
    min_dis:  float,
    max_dis:  float,
    allowance:float,
    n_sim:    int,
    seed:     int,
    lambda_:  float,
) -> float:
    """
    Lagrangian = E[cost(w)] + λ · Σ max(0, w[i] − w[i+1])²

    The penalty term fires when w[i] ≥ w[i+1], i.e. when ships are
    out of chronological order. λ = 1e6 makes ordering violations
    extremely costly, effectively enforcing the constraint.
    """
    cost = _mc_cost(w, rate, min_dis, max_dis, allowance, n_sim, seed)
    penalty = sum(
        lambda_ * max(0.0, w[i] - w[i + 1]) ** 2
        for i in range(len(w) - 1)
    )
    return cost + penalty


def _run_site(
    base_starts: np.ndarray,
    rate:        float,
    site_name:   str,
    min_dis:     float,
    max_dis:     float,
    allowance:   float,
    n_sim:       int,
    seed:        int,
    lambda_:     float,
) -> tuple[np.ndarray, str]:
    """Optimise window starts for one site. Returns (opt_starts, status_msg)."""
    baseline_cost = _mc_cost(
        base_starts, rate, min_dis, max_dis, allowance, n_sim, seed + 1
    )
    print(f"\n[optimizer] {site_name}  ships={len(base_starts)}  "
          f"baseline E[cost]=USD {baseline_cost:,.0f}")

    result = minimize(
        fun     = _lagrangian,
        x0      = base_starts.copy(),
        args    = (rate, min_dis, max_dis, allowance, n_sim, seed, lambda_),
        method  = "L-BFGS-B",
        options = {
            "maxiter": 200,
            "ftol":    1e-4,
            "gtol":    1e-4,
            "eps":     0.5,       # finite-difference step = 0.5 day
            "disp":    False,
        },
    )

    opt_starts = result.x
    gaps       = np.diff(opt_starts)
    print(f"  Status     : {result.message}")
    print(f"  Iterations : {result.nit}")
    print(f"  Gap (opt)  : min={gaps.min():.1f}d  mean={gaps.mean():.1f}d")
    return opt_starts, result.message


def optimize(
    shipments: list[dict],
    params:    dict,
    n_sim:     int   = 500,
    seed:      int   = 42,
    lambda_:   float = 1e6,
) -> dict:
    """
    Optimise window start dates for Gh-1 and SPP3 independently.

    Parameters
    ----------
    shipments : list[dict]   — output of loader.load_assumption()
    params    : dict         — output of loader.load_assumption()
    n_sim     : int          — Monte Carlo draws per function evaluation
    seed      : int          — base RNG seed
    lambda_   : float        — Lagrange penalty weight for ordering violations

    Returns
    -------
    dict with keys:
        gh1_ships, spp3_ships         list[dict]   — per-site ship lists
        gh1_base, gh1_opt             np.ndarray   — window starts (days from BASE_DATE)
        spp3_base, spp3_opt           np.ndarray
        gh1_base_cost, gh1_opt_cost   float        — E[total cost USD]
        spp3_base_cost, spp3_opt_cost float
        base_date                     datetime
        params                        dict
        optimizer_status_gh1          str
        optimizer_status_spp3         str
        n_sim                         int
    """
    min_dis   = params["min_dis"]
    max_dis   = params["max_dis"]
    allowance = params["allowance"]
    gh1_rate  = params["gh1_rate"]
    spp3_rate = params["spp3_rate"]

    BASE_DATE  = datetime(shipments[0]["win_start"].year, 1, 1)
    gh1_ships  = [s for s in shipments if s["site"] == "Gh-1"]
    spp3_ships = [s for s in shipments if s["site"] == "SPP3"]

    # Convert datetime window starts → float days from BASE_DATE
    gh1_base  = np.array(
        [(s["win_start"] - BASE_DATE).days for s in gh1_ships], dtype=float
    )
    spp3_base = np.array(
        [(s["win_start"] - BASE_DATE).days for s in spp3_ships], dtype=float
    )

    # ── Optimise per site ─────────────────────────────────────────────────────
    gh1_opt,  gh1_status  = _run_site(
        gh1_base, gh1_rate, "Gh-1",
        min_dis, max_dis, allowance, n_sim, seed, lambda_,
    )
    spp3_opt, spp3_status = _run_site(
        spp3_base, spp3_rate, "SPP3",
        min_dis, max_dis, allowance, n_sim, seed, lambda_,
    )

    # ── Final cost evaluation ─────────────────────────────────────────────────
    # Same seed for base and opt within each site — only window positions differ,
    # not the random draws. Eliminates sampling noise from the comparison.
    gh1_base_cost  = _mc_cost(gh1_base,  gh1_rate,  min_dis, max_dis, allowance, n_sim, seed + 10)
    gh1_opt_cost   = _mc_cost(gh1_opt,   gh1_rate,  min_dis, max_dis, allowance, n_sim, seed + 10)
    spp3_base_cost = _mc_cost(spp3_base, spp3_rate, min_dis, max_dis, allowance, n_sim, seed + 12)
    spp3_opt_cost  = _mc_cost(spp3_opt,  spp3_rate, min_dis, max_dis, allowance, n_sim, seed + 12)

    total_base = gh1_base_cost + spp3_base_cost
    total_opt  = gh1_opt_cost  + spp3_opt_cost

    print(f"\n[optimizer] RESULTS")
    print(f"  {'Site':<6} {'Baseline':>12}   {'Optimised':>12}   {'Saving':>12}   {'%':>6}")
    print(f"  {'-'*56}")
    for site, base, opt in [
        ("Gh-1",  gh1_base_cost,  gh1_opt_cost),
        ("SPP3",  spp3_base_cost, spp3_opt_cost),
        ("TOTAL", total_base,     total_opt),
    ]:
        saving = base - opt
        pct    = saving / base * 100 if base else 0.0
        print(f"  {site:<6} USD {base:>10,.0f}   USD {opt:>10,.0f}   "
              f"USD {saving:>10,.0f}   {pct:>5.1f}%")

    return {
        "gh1_ships":           gh1_ships,
        "spp3_ships":          spp3_ships,
        "gh1_base":            gh1_base,
        "gh1_opt":             gh1_opt,
        "spp3_base":           spp3_base,
        "spp3_opt":            spp3_opt,
        "gh1_base_cost":       gh1_base_cost,
        "gh1_opt_cost":        gh1_opt_cost,
        "spp3_base_cost":      spp3_base_cost,
        "spp3_opt_cost":       spp3_opt_cost,
        "base_date":           BASE_DATE,
        "params":              params,
        "optimizer_status_gh1":  gh1_status,
        "optimizer_status_spp3": spp3_status,
        "n_sim":               n_sim,
    }


def sim_per_ship(
    win_starts_days: np.ndarray,
    rate:            float,
    min_dis:         float,
    max_dis:         float,
    allowance:       float,
    n_eval:          int,
    seed:            int,
) -> tuple[list[float], list[float]]:
    """
    Evaluate expected demurrage days and cost per individual ship.

    Returns
    -------
    dem_per_ship  : list[float]   E[demurrage days] for each ship
    cost_per_ship : list[float]   E[demurrage cost USD] for each ship
    """
    k        = len(win_starts_days)
    dem_acc  = [0.0] * k
    cost_acc = [0.0] * k
    rng      = random.Random(seed)

    for _ in range(n_eval):
        pf = None
        for i, ws_day in enumerate(win_starts_days):
            arrival   = ws_day + rng.random() * _NW
            nor12     = arrival + 0.5
            commence  = arrival + rng.random() * 0.5
            if pf is not None and pf > commence:
                commence = pf
            dis_days  = rng.random() * (max_dis - min_dis) + min_dis
            finish    = commence + dis_days
            allow_end = nor12 + allowance
            dem       = max(0.0, finish - allow_end)
            dem_acc[i]  += dem
            cost_acc[i] += dem * rate
            pf = finish

    return [d / n_eval for d in dem_acc], [c / n_eval for c in cost_acc]


def sim_per_ship_full(
    win_starts_days: np.ndarray,
    rate:            float,
    min_dis:         float,
    max_dis:         float,
    allowance:       float,
    n_eval:          int,
    seed:            int,
) -> list[list[float]]:
    """
    Return per-ship demurrage cost for ALL n_eval runs.
    Result: runs[ship_idx] = list of n_eval cost values (USD)
    Used for computing per-ship SD / percentiles.
    """
    k   = len(win_starts_days)
    rng = random.Random(seed)
    runs = [[] for _ in range(k)]
    for _ in range(n_eval):
        pf = None
        for i, ws_day in enumerate(win_starts_days):
            arrival  = ws_day + rng.random() * _NW
            nor12    = arrival + 0.5
            commence = arrival + rng.random() * 0.5
            if pf is not None and pf > commence: commence = pf
            dis_days = rng.random() * (max_dis - min_dis) + min_dis
            finish   = commence + dis_days
            dem      = max(0.0, finish - (nor12 + allowance))
            runs[i].append(dem * rate)
            pf = finish
    return runs


def sim_trial_totals(
    win_starts_gh1:  np.ndarray,
    win_starts_spp3: np.ndarray,
    min_dis:  float,
    max_dis:  float,
    allowance:float,
    gh1_rate: float,
    spp3_rate:float,
    n_eval:   int,
    seed:     int,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    Run n_eval full-fleet simulations.
    Returns (gh1_totals, spp3_totals, grand_totals) arrays of length n_eval.
    Each element = total demurrage cost USD for one simulation run.
    """
    rng = random.Random(seed)
    gh1_totals = []; spp3_totals = []

    for _ in range(n_eval):
        gh1_sum = 0.0
        pf = None
        for ws_day in win_starts_gh1:
            arrival  = ws_day + rng.random() * _NW
            nor12    = arrival + 0.5
            commence = arrival + rng.random() * 0.5
            if pf is not None and pf > commence: commence = pf
            dis_days = rng.random() * (max_dis - min_dis) + min_dis
            finish   = commence + dis_days
            gh1_sum += max(0.0, finish - (nor12 + allowance)) * gh1_rate
            pf = finish

        spp3_sum = 0.0
        pf = None
        for ws_day in win_starts_spp3:
            arrival  = ws_day + rng.random() * _NW
            nor12    = arrival + 0.5
            commence = arrival + rng.random() * 0.5
            if pf is not None and pf > commence: commence = pf
            dis_days = rng.random() * (max_dis - min_dis) + min_dis
            finish   = commence + dis_days
            spp3_sum += max(0.0, finish - (nor12 + allowance)) * spp3_rate
            pf = finish

        gh1_totals.append(gh1_sum)
        spp3_totals.append(spp3_sum)

    gh1  = np.array(gh1_totals)
    spp3 = np.array(spp3_totals)
    return gh1, spp3, gh1 + spp3
