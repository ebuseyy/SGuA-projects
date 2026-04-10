from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime
from typing import Callable, Dict, List, Tuple, Type

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# =========================================================
# SAFE OUTPUT HELPERS
# =========================================================

def get_output_dir(dirname: str) -> Path:
    outdir = Path.cwd() / dirname
    outdir.mkdir(parents=True, exist_ok=True)
    return outdir


def safe_output_path(filename: str, outdir: Path) -> Path:
    path = outdir / filename
    if path.exists():
        try:
            with open(path, "a", encoding="utf-8"):
                pass
        except PermissionError:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = outdir / f"{path.stem}_{ts}{path.suffix}"
    return path


# =========================================================
# ORC PROBLEM
# =========================================================

class ORCProblem:
    def __init__(self, penalty_weight: float = 1e6):
        self.n_relays = 6
        self.penalty_weight = penalty_weight
        self.tol = 1e-12

        self.Ikd = np.array([1263.0, 5639.0, 5639.0, 5639.0, 5639.0, 5639.0], dtype=float)

        self.td_lb = np.array([0.2, 0.2, 0.2, 0.2, 0.2, 0.2], dtype=float)
        self.td_ub = np.array([1.0, 1.0, 1.0, 1.0, 1.0, 1.0], dtype=float)

        self.ip_lb = np.array([117.152, 523.0, 400.0, 500.0, 600.0, 400.0], dtype=float)
        self.ip_ub = np.array([128.650, 575.0, 420.0, 510.0, 610.0, 420.0], dtype=float)

        self.ti_min = 1.0
        self.ti_max = 2.2

        self.lb = np.concatenate([self.td_lb, self.ip_lb])
        self.ub = np.concatenate([self.td_ub, self.ip_ub])
        self.dim = len(self.lb)

    def split_variables(self, X: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
        X = np.atleast_2d(X)
        td = X[:, :6]
        ip = X[:, 6:]
        return td, ip

    def compute_ti(self, X: np.ndarray) -> np.ndarray:
        X = np.atleast_2d(X)
        td, ip = self.split_variables(X)
        ratio = self.Ikd / ip
        denom = np.power(ratio, 0.02) - 1.0
        denom = np.where(np.abs(denom) < 1e-12, 1e-12, denom)
        ti = 0.14 * td / denom
        return ti

    def evaluate(self, X: np.ndarray):
        X = np.atleast_2d(X)
        td, ip = self.split_variables(X)
        ti = self.compute_ti(X)
        objective = np.sum(ti, axis=1)

        td_low_v = np.maximum(0.0, self.td_lb - td)
        td_high_v = np.maximum(0.0, td - self.td_ub)
        ip_low_v = np.maximum(0.0, self.ip_lb - ip)
        ip_high_v = np.maximum(0.0, ip - self.ip_ub)
        ti_low_v = np.maximum(0.0, self.ti_min - ti)
        ti_high_v = np.maximum(0.0, ti - self.ti_max)

        c12 = np.maximum(0.0, 0.3 - (ti[:, 0] - ti[:, 1]))
        c23 = np.maximum(0.0, 0.3 - (ti[:, 1] - ti[:, 2]))
        c24 = np.maximum(0.0, 0.3 - (ti[:, 1] - ti[:, 3]))
        c25 = np.maximum(0.0, 0.3 - (ti[:, 1] - ti[:, 4]))
        c26 = np.maximum(0.0, 0.3 - (ti[:, 1] - ti[:, 5]))

        penalty = (
            np.sum(td_low_v ** 2, axis=1)
            + np.sum(td_high_v ** 2, axis=1)
            + np.sum(ip_low_v ** 2, axis=1)
            + np.sum(ip_high_v ** 2, axis=1)
            + np.sum(ti_low_v ** 2, axis=1)
            + np.sum(ti_high_v ** 2, axis=1)
            + c12 ** 2 + c23 ** 2 + c24 ** 2 + c25 ** 2 + c26 ** 2
        )

        penalized_fitness = objective + self.penalty_weight * penalty

        feasible = (
            (np.sum(td_low_v, axis=1) <= self.tol)
            & (np.sum(td_high_v, axis=1) <= self.tol)
            & (np.sum(ip_low_v, axis=1) <= self.tol)
            & (np.sum(ip_high_v, axis=1) <= self.tol)
            & (np.sum(ti_low_v, axis=1) <= self.tol)
            & (np.sum(ti_high_v, axis=1) <= self.tol)
            & (c12 <= self.tol)
            & (c23 <= self.tol)
            & (c24 <= self.tol)
            & (c25 <= self.tol)
            & (c26 <= self.tol)
        )
        return penalized_fitness, objective, feasible, ti

    def check_constraints_from_ti(self, ti: np.ndarray) -> Dict[str, float | bool]:
        ti = np.asarray(ti, dtype=float)
        return {
            "t1_range": self.ti_min <= ti[0] <= self.ti_max,
            "t2_range": self.ti_min <= ti[1] <= self.ti_max,
            "t3_range": self.ti_min <= ti[2] <= self.ti_max,
            "t4_range": self.ti_min <= ti[3] <= self.ti_max,
            "t5_range": self.ti_min <= ti[4] <= self.ti_max,
            "t6_range": self.ti_min <= ti[5] <= self.ti_max,
            "t1_minus_t2": ti[0] - ti[1],
            "t2_minus_t3": ti[1] - ti[2],
            "t2_minus_t4": ti[1] - ti[3],
            "t2_minus_t5": ti[1] - ti[4],
            "t2_minus_t6": ti[1] - ti[5],
            "c12_ok": (ti[0] - ti[1]) >= 0.3,
            "c23_ok": (ti[1] - ti[2]) >= 0.3,
            "c24_ok": (ti[1] - ti[3]) >= 0.3,
            "c25_ok": (ti[1] - ti[4]) >= 0.3,
            "c26_ok": (ti[1] - ti[5]) >= 0.3,
        }


# =========================================================
# COMMON BASE
# =========================================================

class OptimizerBase:
    def __init__(self, problem: ORCProblem, algorithm_name: str, P: int = 100, G: int = 1000, seed: int = 123):
        self.problem = problem
        self.algorithm_name = algorithm_name
        self.P = P
        self.G = G
        self.D = problem.dim
        self.lb = problem.lb.copy()
        self.ub = problem.ub.copy()
        self.span = self.ub - self.lb
        self.rng = np.random.default_rng(seed)

        self.fe_count = 0
        self.gbest_X = np.zeros(self.D, dtype=float)
        self.gbest_penalized = np.inf
        self.gbest_objective = np.inf
        self.gbest_ti = None
        self.best_feasible_X = None
        self.best_feasible_obj = np.inf
        self.best_feasible_ti = None
        self.best_feasible_curve = np.full(G, np.nan, dtype=float)

    def clip(self, X: np.ndarray) -> np.ndarray:
        return np.clip(X, self.lb, self.ub)

    def evaluate_population(self, X: np.ndarray):
        penalized, objective, feasible, ti = self.problem.evaluate(X)
        self.fe_count += X.shape[0]
        return penalized, objective, feasible, ti

    def update_bests(self, X, penalized, objective, feasible, ti):
        idx_pen = np.argmin(penalized)
        if penalized[idx_pen] < self.gbest_penalized:
            self.gbest_penalized = penalized[idx_pen]
            self.gbest_objective = objective[idx_pen]
            self.gbest_X = X[idx_pen].copy()
            self.gbest_ti = ti[idx_pen].copy()
        feasible_idx = np.where(feasible)[0]
        if feasible_idx.size > 0:
            best_idx = feasible_idx[np.argmin(objective[feasible_idx])]
            if objective[best_idx] < self.best_feasible_obj:
                self.best_feasible_obj = objective[best_idx]
                self.best_feasible_X = X[best_idx].copy()
                self.best_feasible_ti = ti[best_idx].copy()

    def maybe_store_curve(self, g: int):
        if self.best_feasible_X is not None:
            self.best_feasible_curve[g] = self.best_feasible_obj
        elif np.isfinite(self.gbest_objective):
            self.best_feasible_curve[g] = self.gbest_objective

    def finalize_result(self) -> Dict:
        if self.best_feasible_X is not None:
            final_X = self.best_feasible_X.copy()
            final_ti = self.best_feasible_ti.copy()
            final_obj = self.best_feasible_obj
            feasible_flag = True
        else:
            final_X = self.gbest_X.copy()
            final_ti = self.gbest_ti.copy() if self.gbest_ti is not None else np.full(6, np.nan)
            final_obj = self.gbest_objective
            feasible_flag = False
        return {
            "x": final_X,
            "td": final_X[:6],
            "Ip": final_X[6:],
            "Ti": final_ti,
            "best_fitness": float(final_obj),
            "feasible": feasible_flag,
            "function_evaluations": int(self.fe_count),
            "feasible_curve": self.best_feasible_curve.copy(),
            "penalized_fitness": float(self.gbest_penalized),
            "algorithm": self.algorithm_name,
            "method_key": self.algorithm_name,
        }


# =========================================================
# WOA
# =========================================================

class WOA(OptimizerBase):
    def __init__(self, problem, P=100, G=1000, seed=123, b=1.0):
        super().__init__(problem, "WOA", P=P, G=G, seed=seed)
        self.b = b

    def opt(self):
        X = self.rng.uniform(self.lb, self.ub, size=(self.P, self.D))
        pen, obj, fea, ti = self.evaluate_population(X)
        self.update_bests(X, pen, obj, fea, ti)
        self.maybe_store_curve(0)

        for g in range(1, self.G):
            a = 2.0 - 2.0 * g / max(self.G - 1, 1)
            best = self.best_feasible_X if self.best_feasible_X is not None else self.gbest_X
            X_new = np.zeros_like(X)
            for i in range(self.P):
                r1 = self.rng.random(self.D)
                r2 = self.rng.random(self.D)
                A = 2 * a * r1 - a
                C = 2 * r2
                p = self.rng.random()
                l = self.rng.uniform(-1.0, 1.0)
                if p < 0.5:
                    if np.mean(np.abs(A)) < 1.0:
                        D = np.abs(C * best - X[i])
                        X_new[i] = best - A * D
                    else:
                        rand_idx = self.rng.integers(0, self.P)
                        X_rand = X[rand_idx]
                        D = np.abs(C * X_rand - X[i])
                        X_new[i] = X_rand - A * D
                else:
                    D = np.abs(best - X[i])
                    X_new[i] = D * np.exp(self.b * l) * np.cos(2 * np.pi * l) + best
            X = self.clip(X_new)
            pen, obj, fea, ti = self.evaluate_population(X)
            self.update_bests(X, pen, obj, fea, ti)
            self.maybe_store_curve(g)
        return self.finalize_result()


# =========================================================
# GWO
# =========================================================

class GWO(OptimizerBase):
    def __init__(self, problem, P=100, G=1000, seed=123):
        super().__init__(problem, "GWO", P=P, G=G, seed=seed)

    def opt(self):
        X = self.rng.uniform(self.lb, self.ub, size=(self.P, self.D))
        pen, obj, fea, ti = self.evaluate_population(X)
        self.update_bests(X, pen, obj, fea, ti)
        self.maybe_store_curve(0)

        for g in range(1, self.G):
            order = np.argsort(pen)
            alpha = X[order[0]].copy()
            beta = X[order[1]].copy()
            delta = X[order[2]].copy()
            a = 2.0 - 2.0 * g / max(self.G - 1, 1)

            X_new = np.zeros_like(X)
            for i in range(self.P):
                r1 = self.rng.random((3, self.D))
                r2 = self.rng.random((3, self.D))
                A1 = 2 * a * r1[0] - a
                C1 = 2 * r2[0]
                A2 = 2 * a * r1[1] - a
                C2 = 2 * r2[1]
                A3 = 2 * a * r1[2] - a
                C3 = 2 * r2[2]
                X1 = alpha - A1 * np.abs(C1 * alpha - X[i])
                X2 = beta - A2 * np.abs(C2 * beta - X[i])
                X3 = delta - A3 * np.abs(C3 * delta - X[i])
                X_new[i] = (X1 + X2 + X3) / 3.0
            X = self.clip(X_new)
            pen, obj, fea, ti = self.evaluate_population(X)
            self.update_bests(X, pen, obj, fea, ti)
            self.maybe_store_curve(g)
        return self.finalize_result()


# =========================================================
# GA
# =========================================================

class GA(OptimizerBase):
    def __init__(self, problem, P=100, G=1000, seed=123, crossover_rate=0.9, mutation_rate=None, elite_count=2):
        super().__init__(problem, "GA", P=P, G=G, seed=seed)
        self.crossover_rate = crossover_rate
        self.mutation_rate = mutation_rate if mutation_rate is not None else 1.0 / self.D
        self.elite_count = elite_count

    def tournament_select(self, X: np.ndarray, pen: np.ndarray, k: int = 3) -> np.ndarray:
        idx = self.rng.choice(len(X), size=k, replace=False)
        return X[idx[np.argmin(pen[idx])]].copy()

    def crossover(self, p1: np.ndarray, p2: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
        if self.rng.random() >= self.crossover_rate:
            return p1.copy(), p2.copy()
        alpha = self.rng.random(self.D)
        c1 = alpha * p1 + (1 - alpha) * p2
        c2 = alpha * p2 + (1 - alpha) * p1
        return c1, c2

    def mutate(self, x: np.ndarray, g: int) -> np.ndarray:
        sigma = 0.10 * (1 - g / max(self.G - 1, 1)) + 0.01
        mask = self.rng.random(self.D) < self.mutation_rate
        noise = self.rng.normal(0.0, sigma, self.D) * self.span
        y = x.copy()
        y[mask] += noise[mask]
        return self.clip(y)

    def opt(self):
        X = self.rng.uniform(self.lb, self.ub, size=(self.P, self.D))
        pen, obj, fea, ti = self.evaluate_population(X)
        self.update_bests(X, pen, obj, fea, ti)
        self.maybe_store_curve(0)

        for g in range(1, self.G):
            order = np.argsort(pen)
            elites = X[order[: self.elite_count]].copy()
            next_pop = [e.copy() for e in elites]

            while len(next_pop) < self.P:
                p1 = self.tournament_select(X, pen)
                p2 = self.tournament_select(X, pen)
                c1, c2 = self.crossover(p1, p2)
                c1 = self.mutate(c1, g)
                c2 = self.mutate(c2, g)
                next_pop.append(c1)
                if len(next_pop) < self.P:
                    next_pop.append(c2)
            X = np.array(next_pop, dtype=float)
            pen, obj, fea, ti = self.evaluate_population(X)
            self.update_bests(X, pen, obj, fea, ti)
            self.maybe_store_curve(g)
        return self.finalize_result()


# =========================================================
# BOA
# =========================================================

class BOA(OptimizerBase):
    def __init__(self, problem, P=100, G=1000, seed=123, sensory_modality=0.01, power_exponent=0.1, switch_probability=0.8):
        super().__init__(problem, "BOA", P=P, G=G, seed=seed)
        self.c = sensory_modality
        self.a = power_exponent
        self.p = switch_probability

    def opt(self):
        X = self.rng.uniform(self.lb, self.ub, size=(self.P, self.D))
        pen, obj, fea, ti = self.evaluate_population(X)
        self.update_bests(X, pen, obj, fea, ti)
        self.maybe_store_curve(0)

        for g in range(1, self.G):
            best = self.best_feasible_X if self.best_feasible_X is not None else self.gbest_X
            intensity = 1.0 / (1.0 + np.maximum(pen, 0.0))
            fragrance = self.c * np.power(intensity, self.a)
            X_new = np.zeros_like(X)
            for i in range(self.P):
                r = self.rng.random(self.D)
                if self.rng.random() < self.p:
                    step = (r * r) * best - X[i]
                    X_new[i] = X[i] + fragrance[i] * step
                else:
                    j, k = self.rng.choice(self.P, size=2, replace=False)
                    step = (r * r) * X[j] - X[k]
                    X_new[i] = X[i] + fragrance[i] * step
            X = self.clip(X_new)
            pen, obj, fea, ti = self.evaluate_population(X)
            self.update_bests(X, pen, obj, fea, ti)
            self.c = self.c + 0.025 / (self.c * self.G)
            self.maybe_store_curve(g)
        return self.finalize_result()


# =========================================================
# ABC
# =========================================================

class ABC(OptimizerBase):
    def __init__(self, problem, P=100, G=1000, seed=123, limit=None):
        super().__init__(problem, "ABC", P=P, G=G, seed=seed)
        self.limit = limit if limit is not None else max(20, self.D * 3)

    def opt(self):
        food = self.rng.uniform(self.lb, self.ub, size=(self.P, self.D))
        pen, obj, fea, ti = self.evaluate_population(food)
        self.update_bests(food, pen, obj, fea, ti)
        trial = np.zeros(self.P, dtype=int)
        self.maybe_store_curve(0)

        for g in range(1, self.G):
            # employed bees
            for i in range(self.P):
                k = self.rng.choice([idx for idx in range(self.P) if idx != i])
                j = self.rng.integers(0, self.D)
                phi = self.rng.uniform(-1.0, 1.0)
                candidate = food[i].copy()
                candidate[j] = food[i, j] + phi * (food[i, j] - food[k, j])
                candidate = self.clip(candidate)
                c_pen, c_obj, c_fea, c_ti = self.evaluate_population(candidate[None, :])
                self.update_bests(candidate[None, :], c_pen, c_obj, c_fea, c_ti)
                if c_pen[0] <= pen[i]:
                    food[i] = candidate
                    pen[i], obj[i], fea[i], ti[i] = c_pen[0], c_obj[0], c_fea[0], c_ti[0]
                    trial[i] = 0
                else:
                    trial[i] += 1

            # onlooker bees
            fitness = 1.0 / (1.0 + np.maximum(pen - np.min(pen), 0.0))
            probs = fitness / np.sum(fitness)
            for _ in range(self.P):
                i = self.rng.choice(self.P, p=probs)
                k = self.rng.choice([idx for idx in range(self.P) if idx != i])
                j = self.rng.integers(0, self.D)
                phi = self.rng.uniform(-1.0, 1.0)
                candidate = food[i].copy()
                candidate[j] = food[i, j] + phi * (food[i, j] - food[k, j])
                candidate = self.clip(candidate)
                c_pen, c_obj, c_fea, c_ti = self.evaluate_population(candidate[None, :])
                self.update_bests(candidate[None, :], c_pen, c_obj, c_fea, c_ti)
                if c_pen[0] <= pen[i]:
                    food[i] = candidate
                    pen[i], obj[i], fea[i], ti[i] = c_pen[0], c_obj[0], c_fea[0], c_ti[0]
                    trial[i] = 0
                else:
                    trial[i] += 1

            # scout bees
            scouts = np.where(trial >= self.limit)[0]
            for i in scouts:
                food[i] = self.rng.uniform(self.lb, self.ub, size=self.D)
                p_i, o_i, f_i, t_i = self.evaluate_population(food[i][None, :])
                self.update_bests(food[i][None, :], p_i, o_i, f_i, t_i)
                pen[i], obj[i], fea[i], ti[i] = p_i[0], o_i[0], f_i[0], t_i[0]
                trial[i] = 0

            self.maybe_store_curve(g)
        return self.finalize_result()


# =========================================================
# LCA (compact self-contained variant)
# =========================================================

class LCA(OptimizerBase):
    def __init__(self, problem, P=100, G=1000, seed=123, momentum=0.6, noise_scale=0.08):
        super().__init__(problem, "LCA", P=P, G=G, seed=seed)
        self.momentum = momentum
        self.noise_scale = noise_scale

    def opt(self):
        X = self.rng.uniform(self.lb, self.ub, size=(self.P, self.D))
        velocity = self.rng.normal(0.0, 0.05, size=(self.P, self.D)) * self.span
        pen, obj, fea, ti = self.evaluate_population(X)
        self.update_bests(X, pen, obj, fea, ti)
        self.maybe_store_curve(0)

        for g in range(1, self.G):
            order = np.argsort(pen)
            X = X[order]
            pen = pen[order]
            obj = obj[order]
            fea = fea[order]
            ti = ti[order]
            velocity = velocity[order]
            best = self.best_feasible_X if self.best_feasible_X is not None else self.gbest_X

            perm = self.rng.permutation(self.P)
            X_new = X.copy()
            vel_new = velocity.copy()
            for a, b in perm.reshape(-1, 2) if self.P % 2 == 0 else []:
                # original indices after permutation order mapping
                ia, ib = a, b
                if pen[ia] <= pen[ib]:
                    winner, loser = ia, ib
                else:
                    winner, loser = ib, ia
                learn = self.rng.random(self.D)
                noise = self.rng.normal(0.0, self.noise_scale, self.D) * self.span * (1 - g / max(self.G - 1, 1))
                vel_new[loser] = self.momentum * velocity[loser] + learn * (X[winner] - X[loser]) + noise
                X_new[loser] = X[loser] + vel_new[loser]
                vel_new[winner] = 0.5 * velocity[winner] + 0.3 * self.rng.random(self.D) * (best - X[winner])
                X_new[winner] = X[winner] + vel_new[winner]

            # if odd population, update last one toward best
            if self.P % 2 == 1:
                i = perm[-1]
                vel_new[i] = self.momentum * velocity[i] + 0.4 * self.rng.random(self.D) * (best - X[i])
                X_new[i] = X[i] + vel_new[i]

            X = self.clip(X_new)
            velocity = vel_new
            pen, obj, fea, ti = self.evaluate_population(X)
            self.update_bests(X, pen, obj, fea, ti)
            self.maybe_store_curve(g)
        return self.finalize_result()


# =========================================================
# RUNNER / REPORTING
# =========================================================

def format_relay_triplet(td: float, ip: float, ti: float) -> str:
    return f"{td:.4f} / {ip:.2f} / {ti:.4f}"


def summarize_results(all_results: List[Dict], method_key: str, display_name: str) -> Dict:
    feasible_results = [r for r in all_results if r["feasible"]]
    total_runs = len(all_results)
    if len(feasible_results) == 0:
        return {
            "method_key": method_key,
            "display_name": display_name,
            "feasible_runs": 0,
            "total_runs": total_runs,
            "mean_best": np.nan,
            "std_best": np.nan,
            "min_best": np.nan,
            "max_best": np.nan,
            "median_best": np.nan,
            "best_run": None,
            "all_feasible_values": np.array([], dtype=float),
        }
    vals = np.array([r["best_fitness"] for r in feasible_results], dtype=float)
    best_run = min(feasible_results, key=lambda x: x["best_fitness"])
    return {
        "method_key": method_key,
        "display_name": display_name,
        "feasible_runs": len(feasible_results),
        "total_runs": total_runs,
        "mean_best": float(np.mean(vals)),
        "std_best": float(np.std(vals)),
        "min_best": float(np.min(vals)),
        "max_best": float(np.max(vals)),
        "median_best": float(np.median(vals)),
        "best_run": best_run,
        "all_feasible_values": vals,
    }


def print_result_table(title: str, result: Dict):
    print("\n" + title)
    print(f"{'Relay':<10}{'td':<12}{'Ip(A)':<12}{'Ti(s)':<12}")
    for i in range(6):
        print(f"{'Relay-' + str(i+1):<10}{result['td'][i]:<12.6f}{result['Ip'][i]:<12.3f}{result['Ti'][i]:<12.6f}")


def print_constraint_check(problem: ORCProblem, result: Dict):
    cc = problem.check_constraints_from_ti(result["Ti"])
    print("\nConstraint Check:")
    print(f"1 <= t1 <= 2.2 : {cc['t1_range']}")
    print(f"1 <= t2 <= 2.2 : {cc['t2_range']}")
    print(f"1 <= t3 <= 2.2 : {cc['t3_range']}")
    print(f"1 <= t4 <= 2.2 : {cc['t4_range']}")
    print(f"1 <= t5 <= 2.2 : {cc['t5_range']}")
    print(f"1 <= t6 <= 2.2 : {cc['t6_range']}")
    print(f"t1 - t2 >= 0.3 : {cc['c12_ok']}   value = {cc['t1_minus_t2']:.6f}")
    print(f"t2 - t3 >= 0.3 : {cc['c23_ok']}   value = {cc['t2_minus_t3']:.6f}")
    print(f"t2 - t4 >= 0.3 : {cc['c24_ok']}   value = {cc['t2_minus_t4']:.6f}")
    print(f"t2 - t5 >= 0.3 : {cc['c25_ok']}   value = {cc['t2_minus_t5']:.6f}")
    print(f"t2 - t6 >= 0.3 : {cc['c26_ok']}   value = {cc['t2_minus_t6']:.6f}")


def print_summary_table(summary: Dict, label_col: str = "Algorithm"):
    print("\n" + "=" * 110)
    print("SUMMARY TABLE")
    print("=" * 110)
    print(f"{'Rank':<6}{label_col:<18}{'Feasible':<12}{'Mean':<14}{'Std':<14}{'Min':<14}{'Max':<14}")
    feasible_str = f"{summary['feasible_runs']}/{summary['total_runs']}"
    mean_str = f"{summary['mean_best']:.6f}" if np.isfinite(summary['mean_best']) else "nan"
    std_str = f"{summary['std_best']:.6f}" if np.isfinite(summary['std_best']) else "nan"
    min_str = f"{summary['min_best']:.6f}" if np.isfinite(summary['min_best']) else "nan"
    max_str = f"{summary['max_best']:.6f}" if np.isfinite(summary['max_best']) else "nan"
    print(f"{1:<6}{summary['display_name']:<18}{feasible_str:<12}{mean_str:<14}{std_str:<14}{min_str:<14}{max_str:<14}")


def build_publication_table(summary: Dict, first_col_name: str = "Algorithm") -> pd.DataFrame:
    best_run = summary.get("best_run", None)
    row = {first_col_name: summary["display_name"], "Total ti": np.nan if best_run is None else float(best_run["best_fitness"])}
    if best_run is None:
        for i in range(6):
            row[f"Relay-{i+1} (td, Ip, ti)"] = "NO FEASIBLE SOLUTION"
    else:
        for i in range(6):
            row[f"Relay-{i+1} (td, Ip, ti)"] = format_relay_triplet(best_run["td"][i], best_run["Ip"][i], best_run["Ti"][i])
    cols = [first_col_name] + [f"Relay-{i+1} (td, Ip, ti)" for i in range(6)] + ["Total ti"]
    return pd.DataFrame([row], columns=cols)


def print_publication_table(publication_df: pd.DataFrame, title_line: str, first_col_name: str = "Algorithm"):
    print("\n" + "=" * 180)
    print(title_line)
    print("=" * 180)
    header = (
        f"{first_col_name:<18}"
        f"{'Relay-1 (td, Ip, ti)':<28}"
        f"{'Relay-2 (td, Ip, ti)':<28}"
        f"{'Relay-3 (td, Ip, ti)':<28}"
        f"{'Relay-4 (td, Ip, ti)':<28}"
        f"{'Relay-5 (td, Ip, ti)':<28}"
        f"{'Relay-6 (td, Ip, ti)':<28}"
        f"{'Total ti':<12}"
    )
    print(header)
    print("-" * 180)
    for _, row in publication_df.iterrows():
        total_ti = row["Total ti"]
        total_str = f"{total_ti:.4f}" if np.isfinite(total_ti) else "nan"
        line = f"{row[first_col_name]:<18}"
        for i in range(6):
            line += f"{row[f'Relay-{i+1} (td, Ip, ti)']:<28}"
        line += total_str
        print(line)


def build_runs_dataframe(results: List[Dict]) -> pd.DataFrame:
    rows = []
    for r in results:
        row = {
            "run_id": r["run_id"],
            "seed": r["seed"],
            "feasible": r["feasible"],
            "best_fitness": r["best_fitness"],
            "penalized_fitness": r["penalized_fitness"],
            "function_evaluations": r["function_evaluations"],
        }
        for i in range(6):
            row[f"td_{i+1}"] = r["td"][i]
            row[f"Ip_{i+1}"] = r["Ip"][i]
            row[f"Ti_{i+1}"] = r["Ti"][i]
        rows.append(row)
    return pd.DataFrame(rows)


def save_convergence_csv(results: List[Dict], outdir: Path, algorithm_name: str):
    curves = np.array([r["feasible_curve"] for r in results], dtype=float)
    df = pd.DataFrame(curves.T)
    df.index = np.arange(1, curves.shape[1] + 1)
    df.index.name = "iteration"
    path = safe_output_path(f"{algorithm_name.lower()}_or_c_feasible_curves.csv", outdir)
    df.to_csv(path, encoding="utf-8-sig")
    return path


def plot_single_algorithm_distribution(results: List[Dict], algorithm_name: str, outdir: Path):
    vals = [r["best_fitness"] for r in results if r["feasible"]]
    if len(vals) == 0:
        return None
    fig, ax = plt.subplots(figsize=(7, 4.8))
    ax.boxplot([vals], tick_labels=[algorithm_name], showmeans=True)
    jitter = np.random.default_rng(123).normal(1.0, 0.02, size=len(vals))
    ax.scatter(jitter, vals, alpha=0.65, s=24)
    ax.set_ylabel("Best feasible objective (sum_ti)")
    ax.set_title(f"Distribution of ORC best fitness values over 30 runs - {algorithm_name}")
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    path = safe_output_path(f"{algorithm_name.lower()}_orc_distribution_boxplot.png", outdir)
    fig.savefig(path, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return path


def plot_single_algorithm_convergence(results: List[Dict], algorithm_name: str, outdir: Path):
    curves = np.array([r["feasible_curve"] for r in results], dtype=float)
    if curves.size == 0 or not np.any(np.isfinite(curves)):
        return None
    iterations = np.arange(1, curves.shape[1] + 1)
    mean_curve = np.nanmean(curves, axis=0)
    std_curve = np.nanstd(curves, axis=0)

    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    for c in curves:
        ax.plot(iterations, c, linewidth=0.8, alpha=0.15)
    ax.plot(iterations, mean_curve, linewidth=2.2, label="Mean feasible curve")
    lower = mean_curve - std_curve
    upper = mean_curve + std_curve
    ax.fill_between(iterations, lower, upper, alpha=0.2)
    ax.set_xlabel("Iteration")
    ax.set_ylabel("Best feasible objective (sum_ti)")
    ax.set_title(f"Mean convergence behavior over 30 runs - {algorithm_name}")
    ax.grid(True, alpha=0.3)
    ax.legend()
    plt.tight_layout()
    path = safe_output_path(f"{algorithm_name.lower()}_orc_mean_convergence.png", outdir)
    fig.savefig(path, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return path


def run_algorithm_experiments(
    optimizer_cls: Type[OptimizerBase],
    algorithm_name: str,
    n_runs: int = 30,
    P: int = 100,
    G: int = 1000,
    penalty_weight: float = 1e6,
    base_seed: int = 123,
    optimizer_kwargs: Dict | None = None,
):
    optimizer_kwargs = optimizer_kwargs or {}
    problem = ORCProblem(penalty_weight=penalty_weight)
    all_results: List[Dict] = []
    print("\n" + "=" * 100)
    print(f"{algorithm_name} ON THE ORC PROBLEM")
    print("=" * 100)
    for run_id in range(n_runs):
        seed = base_seed + run_id
        opt = optimizer_cls(problem=problem, P=P, G=G, seed=seed, **optimizer_kwargs)
        result = opt.opt()
        result["run_id"] = run_id + 1
        result["seed"] = seed
        all_results.append(result)
        if result["feasible"]:
            print(f"Run {run_id + 1:02d} | {algorithm_name} | Feasible: True | Best Fitness (sum_ti): {result['best_fitness']:.6f}")
        else:
            print(f"Run {run_id + 1:02d} | {algorithm_name} | Feasible: False | No feasible solution found")
    summary = summarize_results(all_results, method_key=algorithm_name, display_name=algorithm_name)
    return problem, all_results, summary


def save_single_algorithm_outputs(problem: ORCProblem, results: List[Dict], summary: Dict, algorithm_name: str, outdir_name: str | None = None):
    outdir = get_output_dir(outdir_name or f"outputs_{algorithm_name.lower()}")

    print_summary_table(summary, label_col="Algorithm")
    if summary.get("best_run") is not None:
        print_result_table(f"Best feasible run for {algorithm_name}", summary["best_run"])
        print_constraint_check(problem, summary["best_run"])
        print(f"\nBest fitness (sum_ti): {summary['best_run']['best_fitness']:.6f}")
        print(f"Function evaluations: {summary['best_run']['function_evaluations']}")

    publication_df = build_publication_table(summary, first_col_name="Algorithm")
    print_publication_table(
        publication_df,
        title_line=f"OPTIMAL RELAY COORDINATION PARAMETERS (TD, IP, TI) AND TOTAL RUN TIMES\nOBTAINED FROM THE {algorithm_name} ALGORITHM",
        first_col_name="Algorithm",
    )

    summary_df = pd.DataFrame([
        {
            "Algorithm": summary["display_name"],
            "Feasible": f"{summary['feasible_runs']}/{summary['total_runs']}",
            "Mean": summary["mean_best"],
            "Std": summary["std_best"],
            "Min": summary["min_best"],
            "Max": summary["max_best"],
            "Median": summary["median_best"],
        }
    ])
    runs_df = build_runs_dataframe(results)

    pub_path = safe_output_path(f"{algorithm_name.lower()}_orc_publication_table.csv", outdir)
    publication_df.to_csv(pub_path, index=False, encoding="utf-8-sig")
    summary_path = safe_output_path(f"{algorithm_name.lower()}_orc_summary_statistics.csv", outdir)
    summary_df.to_csv(summary_path, index=False, encoding="utf-8-sig")
    runs_path = safe_output_path(f"{algorithm_name.lower()}_orc_run_results.csv", outdir)
    runs_df.to_csv(runs_path, index=False, encoding="utf-8-sig")
    conv_csv_path = save_convergence_csv(results, outdir, algorithm_name)
    boxplot_path = plot_single_algorithm_distribution(results, algorithm_name, outdir)
    conv_fig_path = plot_single_algorithm_convergence(results, algorithm_name, outdir)

    print(f"\nSaved summary CSV: {summary_path.resolve()}")
    print(f"Saved run-level CSV: {runs_path.resolve()}")
    print(f"Saved publication CSV: {pub_path.resolve()}")
    print(f"Saved convergence CSV: {conv_csv_path.resolve()}")
    if boxplot_path is not None:
        print(f"Saved distribution figure: {boxplot_path.resolve()}")
    if conv_fig_path is not None:
        print(f"Saved convergence figure: {conv_fig_path.resolve()}")

    return {
        "outdir": outdir,
        "publication_csv": pub_path,
        "summary_csv": summary_path,
        "runs_csv": runs_path,
        "convergence_csv": conv_csv_path,
        "distribution_figure": boxplot_path,
        "convergence_figure": conv_fig_path,
    }
