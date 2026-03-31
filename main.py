"""
main.py
=======
Orchestration entry point for the OPTIFUEL demurrage optimizer.

Pipeline (per input file)
-------------------------
1. load_assumption()  — read shipment data + parameters from Assumption sheet
2. Monte Carlo trials — baseline simulation (for Summary sheet trial table)
3. optimize()         — Lagrange + Monte Carlo window-start optimizer (per site)
4. sim_per_ship()     — evaluate per-ship E[cost] for baseline and optimised
5. write_excel()      — write 3-sheet result workbook to data/processed/

Output naming
-------------
Result filename is derived automatically from the input filename:
  Estimate_Demurrage_01.xlsm  →  data/processed/Result_Estimate_Demurrage_01.xlsx
  Estimate_Demurrage_02.xlsx  →  data/processed/Result_Estimate_Demurrage_02.xlsx

Where baseline cost comes from
-------------------------------
Baseline cost is calculated inside optimize() in src/optimizer.py.
It runs _mc_cost() on the ORIGINAL window starts read from the Assumption sheet
(gh1_base / spp3_base arrays) — nothing from the Simulation sheet is used.
Formula: baseline_cost = _mc_cost(original_window_starts, rate, params, n_sim, seed)

Usage
-----
  python3 main.py

To add a new input, append one line to INPUTS below — no other changes needed.
"""

import random
import numpy as np
from pathlib import Path
from datetime import datetime

import sys

sys.path.insert(0, str(Path(__file__).parent))

from src.loader import load_assumption, load_simulation
from src.optimizer import optimize, sim_per_ship, sim_trial_totals
from src.writer import write_excel


# ══════════════════════════════════════════════════════════════════════════════
# INPUTS  — list of files in data/raw/ to process
# Output → data/processed/Result_<stem>.xlsx  (stem = filename without extension)
# ══════════════════════════════════════════════════════════════════════════════

INPUTS = [
    "data/raw/Estimate Demurrage.xlsm",
]


# ══════════════════════════════════════════════════════════════════════════════
# SHARED CONFIG  — applied to every input
# ══════════════════════════════════════════════════════════════════════════════

CONFIG = {
    "n_trials": 300,  # MC trials for Summary sheet trial table
    "n_sim": 500,  # MC draws per optimizer function evaluation
    "n_eval": 2000,  # MC draws for final per-ship cost evaluation
    "lambda_": 1e6,  # Lagrange penalty weight for ordering violations
    "seed": 42,  # base RNG seed
}


# ══════════════════════════════════════════════════════════════════════════════
# run_one — full pipeline for a single input file
# ══════════════════════════════════════════════════════════════════════════════


def run_one(assumption_path: str, output_path: str) -> dict:
    """
    Run the full 5-step pipeline for one input file.

    Baseline cost origin
    --------------------
    Calculated in src/optimizer.py → optimize() → _mc_cost(gh1_base, ...) and
    _mc_cost(spp3_base, ...) where gh1_base / spp3_base are the original window
    starts read from the Assumption sheet. No data from the Simulation sheet
    is used — the model generates synthetic arrivals from the parameter ranges.

    Parameters
    ----------
    assumption_path : absolute path to source .xlsm / .xlsx
    output_path     : absolute path for result .xlsx
    """
    stem = Path(assumption_path).stem  # e.g. "Estimate_Demurrage_01"

    print()
    print("─" * 65)
    print(f"  FILE   : {stem}")
    print(f"  SOURCE : {assumption_path}")
    print(f"  OUTPUT : {output_path}")
    print("─" * 65)

    # ── Step 1: Load ──────────────────────────────────────────────────────────
    shipments, params = load_assumption(assumption_path)
    actual_sim = load_simulation(
        assumption_path
    )  # per-ship actual dem from Simulation sheet
    MIN_DIS = params["min_dis"]
    MAX_DIS = params["max_dis"]
    ALLOW = params["allowance"]
    GH1_RATE = params["gh1_rate"]
    SPP3_RATE = params["spp3_rate"]

    # ── Step 2: Monte Carlo baseline trials ───────────────────────────────────
    # Simulates n_trials full years from the Assumption window starts.
    # Used only for the Summary sheet trial table — not for optimizer input.
    print(f"\n  [Step 2] Running {CONFIG['n_trials']} baseline MC trials...")

    def one_trial(trial_seed: int) -> tuple:
        rng = random.Random(trial_seed)
        prev_finish = {"Gh-1": None, "SPP3": None}
        dis_sum = total_dem = gh1_dem = spp3_dem = 0.0
        for sm in shipments:
            site = sm["site"]
            ws_day = (sm["win_start"] - datetime(sm["win_start"].year, 1, 1)).days
            arrival = ws_day + rng.random() * sm["num_window"]
            nor12 = arrival + 0.5
            pf = prev_finish[site]
            commence = arrival + rng.random() * 0.5
            if pf and pf > commence:
                commence = pf
            dis_days = rng.random() * (MAX_DIS - MIN_DIS) + MIN_DIS
            finish = commence + dis_days
            dem = max(0.0, finish - (nor12 + ALLOW))
            dis_sum += dis_days
            total_dem += dem
            if site == "Gh-1":
                gh1_dem += dem
            else:
                spp3_dem += dem
            prev_finish[site] = finish
        return dis_sum / len(shipments), total_dem, gh1_dem, spp3_dem

    trials = [one_trial(i) for i in range(CONFIG["n_trials"])]
    total_dem_vals = [t[1] for t in trials]
    grand_avg = np.mean(total_dem_vals)
    grand_sd = np.std(total_dem_vals, ddof=1)
    conf_95 = 1.96 * grand_sd / np.sqrt(CONFIG["n_trials"])
    print(
        f"  AVG dem : {grand_avg:.4f}d  |  SD: {grand_sd:.4f}  |  "
        f"95% CI: [{grand_avg-conf_95:.4f}, {grand_avg+conf_95:.4f}]"
    )

    # ── Step 3: Optimize ──────────────────────────────────────────────────────
    # Baseline cost computed here inside optimize() via _mc_cost(original_window_starts)
    # See src/optimizer.py lines 201-206 for the exact calculation.
    print(f"\n  [Step 3] Lagrange + Monte Carlo optimizer...")
    opt_result = optimize(
        shipments=shipments,
        params=params,
        n_sim=CONFIG["n_sim"],
        seed=CONFIG["seed"],
        lambda_=CONFIG["lambda_"],
    )
    BASE_DATE = opt_result["base_date"]

    # ── Step 4: Per-ship evaluation ───────────────────────────────────────────
    print(f"\n  [Step 4] Per-ship E[cost] ({CONFIG['n_eval']} sims)...")
    gh1_bd, gh1_bc = sim_per_ship(
        opt_result["gh1_base"],
        GH1_RATE,
        MIN_DIS,
        MAX_DIS,
        ALLOW,
        CONFIG["n_eval"],
        seed=99,
    )
    gh1_od, gh1_oc = sim_per_ship(
        opt_result["gh1_opt"],
        GH1_RATE,
        MIN_DIS,
        MAX_DIS,
        ALLOW,
        CONFIG["n_eval"],
        seed=100,
    )
    spp3_bd, spp3_bc = sim_per_ship(
        opt_result["spp3_base"],
        SPP3_RATE,
        MIN_DIS,
        MAX_DIS,
        ALLOW,
        CONFIG["n_eval"],
        seed=101,
    )
    spp3_od, spp3_oc = sim_per_ship(
        opt_result["spp3_opt"],
        SPP3_RATE,
        MIN_DIS,
        MAX_DIS,
        ALLOW,
        CONFIG["n_eval"],
        seed=102,
    )

    # ── Step 4b: Trial totals for MC Simulation sheet ───────────────────────
    print(
        f"\n  [Step 4b] Trial totals for MC Simulation sheet ({CONFIG['n_eval']:,} sims)..."
    )
    gh1_base_tot, spp3_base_tot, total_base_tot = sim_trial_totals(
        opt_result["gh1_base"],
        opt_result["spp3_base"],
        MIN_DIS,
        MAX_DIS,
        ALLOW,
        GH1_RATE,
        SPP3_RATE,
        CONFIG["n_eval"],
        seed=202,
    )
    gh1_opt_tot, spp3_opt_tot, total_opt_tot = sim_trial_totals(
        opt_result["gh1_opt"],
        opt_result["spp3_opt"],
        MIN_DIS,
        MAX_DIS,
        ALLOW,
        GH1_RATE,
        SPP3_RATE,
        CONFIG["n_eval"],
        seed=202,
    )

    # ── Step 5: Write Excel ───────────────────────────────────────────────────
    print(f"\n  [Step 5] Writing → {output_path}")
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    write_excel(
        output_path=output_path,
        shipments=shipments,
        params=params,
        opt_result=opt_result,
        gh1_bd=gh1_bd,
        gh1_bc=gh1_bc,
        gh1_od=gh1_od,
        gh1_oc=gh1_oc,
        spp3_bd=spp3_bd,
        spp3_bc=spp3_bc,
        spp3_od=spp3_od,
        spp3_oc=spp3_oc,
        BASE_DATE=BASE_DATE,
        n_eval=CONFIG["n_eval"],
        actual_sim=actual_sim,
        gh1_base_tot=gh1_base_tot,
        spp3_base_tot=spp3_base_tot,
        total_base_tot=total_base_tot,
        gh1_opt_tot=gh1_opt_tot,
        spp3_opt_tot=spp3_opt_tot,
        total_opt_tot=total_opt_tot,
    )

    total_base = opt_result["gh1_base_cost"] + opt_result["spp3_base_cost"]
    total_opt = opt_result["gh1_opt_cost"] + opt_result["spp3_opt_cost"]
    saving = total_base - total_opt

    print(
        f"\n  Baseline : USD {total_base:>10,.0f}  ← _mc_cost(original window starts)"
    )
    print(f"  Optimised: USD {total_opt:>10,.0f}  ← _mc_cost(optimised window starts)")
    print(f"  Saving   : USD {saving:>10,.0f}  ({saving/total_base*100:.1f}%)")

    return {
        "stem": stem,
        "baseline_cost": total_base,
        "opt_cost": total_opt,
        "saving": saving,
        "output_path": output_path,
    }


# ══════════════════════════════════════════════════════════════════════════════
# MAIN — loop over all inputs, output named Result_<stem>.xlsx
# ══════════════════════════════════════════════════════════════════════════════


def main():
    base_dir = Path(__file__).parent

    print("=" * 65)
    print("OPTIFUEL — Demurrage Optimizer")
    print("=" * 65)
    print(f"Inputs : {len(INPUTS)}")
    print(
        f"Config : n_trials={CONFIG['n_trials']}  "
        f"n_sim={CONFIG['n_sim']}  n_eval={CONFIG['n_eval']}"
    )
    print()
    print("Output naming:  Result_<input_stem>.xlsx")
    print("Baseline cost:  _mc_cost(original_window_starts)  [src/optimizer.py]")

    results = []
    for file_path in INPUTS:
        stem = Path(file_path).stem
        assumption_path = str(base_dir / file_path)
        output_path = str(base_dir / "data" / "processed" / f"Result_{stem}.xlsx")
        result = run_one(assumption_path, output_path)
        results.append(result)

    # ── Final comparison table ─────────────────────────────────────────────────
    print()
    print("=" * 65)
    print("ALL DONE — COMPARISON")
    print("=" * 65)
    print(
        f"  {'Input file':<35} {'Baseline':>12}   {'Optimised':>12}   {'Saving':>12}   {'%':>6}"
    )
    print(f"  {'─'*80}")
    for r in results:
        pct = r["saving"] / r["baseline_cost"] * 100 if r["baseline_cost"] else 0
        print(
            f"  {r['stem']:<35} "
            f"USD {r['baseline_cost']:>10,.0f}   "
            f"USD {r['opt_cost']:>10,.0f}   "
            f"USD {r['saving']:>10,.0f}   "
            f"{pct:>5.1f}%"
        )
    print()
    for r in results:
        print(f"  → {r['output_path']}")


if __name__ == "__main__":
    main()
