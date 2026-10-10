"""
lotto_engine.py
===============
Loads PCSO draw history from a CSV, analyses it, and generates new 6-number
combinations that are GUARANTEED not to match any past winning combination.

Expected CSV columns (header names are matched case-insensitively):
    LOTTO GAME, COMBINATION, DATE
    e.g.  Megalotto 6/45, 09-31-28-43-07-15, 1/4/2016

Key design points
-----------------
* A combination is compared as a SORTED TUPLE. The file stores numbers in the
  order they were drawn, so "09-31-28-43-07-15" and "07-09-15-28-31-43" are the
  same combination. Comparing raw strings would miss duplicates.
* Uniqueness is enforced three ways: a set lookup, a within-batch check, and a
  final independent re-read of the CSV file (verify_against_file).
* The generator is "data-informed", not predictive. See the analysis report:
  the history is statistically consistent with a fair random draw.

Usage
-----
    python lotto_engine.py PCSO_SUMMARY_RESULT.csv --game 6/45 --count 5
    python lotto_engine.py PCSO_SUMMARY_RESULT.csv --game 6/45 --backtest
"""

from __future__ import annotations

import argparse
import csv
import math
import random
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

PICK_SIZE = 6
GAME_RANGES = {"6/42": 42, "6/45": 45, "6/49": 49, "6/55": 55, "6/58": 58}


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def normalize(numbers) -> tuple[int, ...]:
    """Return a canonical (sorted tuple of ints) form of a combination."""
    return tuple(sorted(int(n) for n in numbers))


def parse_combination(text: str) -> tuple[int, ...]:
    """'09-31-28-43-07-15' -> (7, 9, 15, 28, 31, 43). Tolerates spaces/commas."""
    parts = [p for p in re.split(r"[-,\s]+", str(text).strip()) if p]
    return normalize(parts)


COMBO_HEADERS = ("COMBINATION", "COMBINATIONS")
DATE_HEADERS = ("DATE", "DRAW DATE")
NAME_TO_GAME = {"ultra": "6/58", "grand": "6/55", "super": "6/49", "mega": "6/45"}


def game_key(label) -> str | None:
    """
    Map a game label or Excel sheet name to a game key.
    'Megalotto 6/45' -> '6/45', 'Ultra Lotto 6-58' -> '6/58', 'Superlotto' -> '6/49'.
    (Excel sheet names cannot contain '/', so '-', '.' and '_' are accepted too.)
    """
    text = str(label or "").lower()
    m = re.search(r"(?<!\d)0?6\s*[/\-._]?\s*(42|45|49|55|58)(?!\d)", text)
    if m:
        return f"6/{m.group(1)}"
    for word, key in NAME_TO_GAME.items():
        if word in text:
            return key
    return None


def parse_date(text) -> datetime | None:
    if isinstance(text, datetime):
        return text
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(text).strip(), fmt)
        except ValueError:
            continue
    return None


def chi_square_p(stat: float, dof: int) -> float:
    """Upper-tail p-value, Wilson-Hilferty approximation (no scipy needed)."""
    z = ((stat / dof) ** (1 / 3) - (1 - 2 / (9 * dof))) / math.sqrt(2 / (9 * dof))
    return 0.5 * math.erfc(z / math.sqrt(2))


def count_consecutive(combo: tuple[int, ...]) -> int:
    return sum(1 for a, b in zip(combo, combo[1:]) if b - a == 1)


def percentile(sorted_values: list[float], q: float) -> float:
    idx = min(len(sorted_values) - 1, max(0, int(round(q * (len(sorted_values) - 1)))))
    return sorted_values[idx]


def _iter_one_file(path: Path):
    """
    Yield (line_no, game_label, combination_text, date_value) from a .csv or .xlsx.

    * CSV:  uses the LOTTO GAME column.
    * XLSX: reads EVERY sheet. If a sheet has no LOTTO GAME column, the sheet
            name is used as the game label (e.g. a sheet called "Ultra Lotto 6/58").
    """
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            header = next(rows, None)
            if not header:
                continue
            cols = {str(h).strip().upper(): i for i, h in enumerate(header) if h is not None}
            ck = next((k for k in COMBO_HEADERS if k in cols), None)
            dk = next((k for k in DATE_HEADERS if k in cols), None)
            if ck is None or dk is None:
                raise ValueError(f"Sheet '{ws.title}' needs COMBINATION(S) and DATE / DRAW DATE columns.")
            for line_no, row in enumerate(rows, start=2):
                if row is None or row[cols[ck]] is None:
                    continue
                label = row[cols["LOTTO GAME"]] if "LOTTO GAME" in cols else ws.title
                yield line_no, label or ws.title, row[cols[ck]], row[cols[dk]]
        return

    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        cols = {(c or "").strip().upper(): c for c in reader.fieldnames or []}
        combo_col = next((cols[k] for k in COMBO_HEADERS if k in cols), None)
        date_col = next((cols[k] for k in DATE_HEADERS if k in cols), None)
        if combo_col is None or date_col is None:
            raise ValueError(f"{path.name}: needs a COMBINATION(S) column and a DATE / DRAW DATE column.")
        game_col = cols.get("LOTTO GAME")
        for line_no, row in enumerate(reader, start=2):
            raw = (row.get(combo_col) or "").strip()
            if not raw:
                continue                      # blank spacer rows at the end of the file
            label = (row.get(game_col) if game_col else None) or path.stem   # file name as backup
            yield line_no, label, raw, row.get(date_col)


def iter_source_rows(path):
    """
    Yield rows from ONE file, or from EVERY .csv/.xlsx file inside a folder.
    Using a folder lets you keep one file per game (06-42.csv, 06-45.csv, ...).
    """
    path = Path(path)
    if path.is_dir():
        files = sorted(p for p in path.iterdir()
                       if p.suffix.lower() in (".csv", ".xlsx", ".xlsm") and not p.name.startswith("~"))
        for file in files:
            yield from _iter_one_file(file)
    else:
        yield from _iter_one_file(path)


def games_in_file(path) -> dict[str, int]:
    """Which games does the file contain? -> {'6/45': 1615, ...}"""
    found: Counter = Counter()
    for _, label, _, _ in iter_source_rows(Path(path)):
        key = game_key(label)
        if key in GAME_RANGES:
            found[key] += 1
    return dict(found)


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------
@dataclass
class LoadReport:
    rows_read: int = 0
    rows_used: int = 0
    rows_skipped: list[tuple[int, str]] = field(default_factory=list)  # (line, reason)
    exact_duplicate_rows: list[tuple[int, str]] = field(default_factory=list)
    first_date: datetime | None = None
    last_date: datetime | None = None


class LottoHistory:
    """Holds the draw history for ONE game and answers 'has this been drawn?'."""

    def __init__(self, csv_path: str | Path, game: str = "6/45"):
        if game not in GAME_RANGES:
            raise ValueError(f"Unknown game {game!r}. Choose from {list(GAME_RANGES)}")
        self.csv_path = Path(csv_path)
        self.game = game
        self.max_num = GAME_RANGES[game]
        self.draws: list[tuple[datetime | None, tuple[int, ...]]] = []  # chronological
        self.known: set[tuple[int, ...]] = set()
        self.report = LoadReport()
        self.load()

    # ---- loading ---------------------------------------------------------
    def load(self) -> LoadReport:
        """(Re)read the CSV. Call again after the file is updated."""
        self.draws.clear()
        self.known.clear()
        self.report = LoadReport()
        seen_rows: set[tuple] = set()

        for line_no, label, raw, raw_date in iter_source_rows(self.csv_path):
            self.report.rows_read += 1
            if game_key(label) != self.game:
                continue  # a different game; not an error
            try:
                combo = parse_combination(raw)
            except ValueError:
                self.report.rows_skipped.append((line_no, f"not numeric: {raw!r}"))
                continue
            if len(combo) != PICK_SIZE or len(set(combo)) != PICK_SIZE:
                self.report.rows_skipped.append((line_no, f"needs 6 different numbers: {raw!r}"))
                continue
            if combo[0] < 1 or combo[-1] > self.max_num:
                self.report.rows_skipped.append((line_no, f"out of range 1-{self.max_num}: {raw!r}"))
                continue

            date = parse_date(raw_date)
            row_key = (combo, date)
            if row_key in seen_rows:  # same combo AND same date = data-entry repeat
                self.report.exact_duplicate_rows.append((line_no, str(raw)))
                continue
            seen_rows.add(row_key)
            self.draws.append((date, combo))
            self.known.add(combo)

        self.draws.sort(key=lambda d: d[0] or datetime.min)
        self.report.rows_used = len(self.draws)
        dates = [d for d, _ in self.draws if d]
        if dates:
            self.report.first_date, self.report.last_date = dates[0], dates[-1]
        return self.report

    # ---- duplicate checks ------------------------------------------------
    def is_duplicate(self, numbers) -> bool:
        return normalize(numbers) in self.known

    def add_result(self, numbers, date: datetime | None = None, save: bool = True) -> None:
        """Record a new official draw so future generations avoid it too."""
        combo = normalize(numbers)
        if len(combo) != PICK_SIZE or len(set(combo)) != PICK_SIZE:
            raise ValueError("A draw needs 6 different numbers.")
        if combo[0] < 1 or combo[-1] > self.max_num:
            raise ValueError(f"Numbers must be between 1 and {self.max_num}.")
        self.draws.append((date, combo))
        self.known.add(combo)
        if save:
            target = self.csv_path / "NEW_DRAWS.csv" if self.csv_path.is_dir() else self.csv_path
            if target.suffix.lower() != ".csv":
                raise ValueError("Saving new draws is only supported for .csv files.")
            label = f"Lotto {self.game}"
            stamp = (date or datetime.now()).strftime("%m/%d/%Y").lstrip("0").replace("/0", "/")
            is_new = not target.exists()
            with open(target, "a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                if is_new:
                    w.writerow(["LOTTO GAME", "COMBINATIONS", "DRAW DATE"])
                w.writerow([label, "-".join(f"{n:02d}" for n in combo), stamp])

    def verify_against_file(self, combos) -> list[tuple[int, ...]]:
        """
        Independent final check: re-read the CSV from disk and compare every
        proposed combo against every row. Returns the combos that collide
        (an empty list means all are unique).
        """
        wanted = {normalize(c) for c in combos}
        collisions = set()
        for _, label, raw, _d in iter_source_rows(self.csv_path):
            if game_key(label) != self.game:
                continue
            try:
                row_combo = parse_combination(raw)
            except ValueError:
                continue
            if row_combo in wanted:
                collisions.add(row_combo)
        # also compare against anything added in memory but not yet saved
        collisions |= wanted & self.known
        return sorted(collisions)

    # ---- analysis --------------------------------------------------------
    def analyze(self) -> dict:
        combos = [c for _, c in self.draws]
        n_draws = len(combos)
        counts = Counter(n for c in combos for n in c)
        expected = n_draws * PICK_SIZE / self.max_num
        chi2 = sum((counts[n] - expected) ** 2 / expected for n in range(1, self.max_num + 1))
        sums = sorted(sum(c) for c in combos)
        mean_sum = sum(sums) / n_draws
        sd_sum = math.sqrt(sum((s - mean_sum) ** 2 for s in sums) / (n_draws - 1))
        half = self.max_num // 2
        odd = Counter(sum(n % 2 for n in c) for c in combos)
        low = Counter(sum(n <= half for n in c) for c in combos)
        consec = Counter(count_consecutive(c) for c in combos)
        repeats = [len(set(a) & set(b)) for a, b in zip(combos, combos[1:])]
        total_space = math.comb(self.max_num, PICK_SIZE)
        return {
            "draws": n_draws,
            "total_possible": total_space,
            "counts": counts,
            "expected_per_number": expected,
            "chi2": chi2,
            "chi2_p": chi_square_p(chi2, self.max_num - 1),
            "hot": counts.most_common(8),
            "cold": sorted(counts.items(), key=lambda kv: kv[1])[:8],
            "sum_mean": mean_sum,
            "sum_sd": sd_sum,
            "sum_p05": percentile(sums, 0.05),
            "sum_p95": percentile(sums, 0.95),
            "odd_dist": dict(sorted(odd.items())),
            "low_dist": dict(sorted(low.items())),
            "consecutive_dist": dict(sorted(consec.items())),
            "share_with_consecutive": 1 - consec[0] / n_draws,
            "avg_repeat_from_prev": sum(repeats) / len(repeats) if repeats else 0.0,
            "expected_repeat_from_prev": PICK_SIZE * PICK_SIZE / self.max_num,
            "half": half,
        }


# --------------------------------------------------------------------------
# Generator
# --------------------------------------------------------------------------
class Generator:
    """
    Data-informed generator with duplicate protection.

    1. Number weights = historical frequency, shrunk toward uniform
       (`shrink` pseudo-draws) because the data shows no significant bias.
       Optionally blended with a recent window (`recent_weight`, `recent_window`).
    2. A candidate must pass the "typical shape" profile learned from the data:
       sum, odd/even split, low/high split, consecutive pairs.
    3. A candidate is REJECTED if it matches any historical combination, or
       any combination already produced in this batch. Then it is regenerated.
    """

    def __init__(self, history: LottoHistory, shrink: float = 3.0,
                 recent_window: int = 100, recent_weight: float = 0.0,
                 min_share: float = 0.05, seed: int | None = None,
                 method: str = "hot"):
        self.h = history
        self.rng = random.Random(seed)
        stats = history.analyze()
        self.stats = stats
        n = history.max_num
        expected = stats["expected_per_number"]

        base = {k: stats["counts"][k] + shrink * expected / max(shrink, 1) for k in range(1, n + 1)}
        total = sum(base.values())
        share = {k: v / total for k, v in base.items()}
        if recent_weight > 0:
            recent = Counter(x for _, c in history.draws[-recent_window:] for x in c)
            rtot = sum(recent.values()) or 1
            share = {k: (1 - recent_weight) * share[k] + recent_weight * (recent[k] + 0.5) / (rtot + 0.5 * n)
                     for k in share}
        if method == "cold":            # favour rarely drawn numbers
            inv = {k: 1 / v for k, v in share.items()}
            t = sum(inv.values()); share = {k: v / t for k, v in inv.items()}
        elif method == "random":        # ignore frequency entirely
            share = {k: 1 / n for k in share}
        self.numbers = list(range(1, n + 1))
        self.weights = [share[k] for k in self.numbers]

        # Profile limits learned from history: keep shapes holding >= min_share of draws
        d = stats["draws"]
        self.ok_odd = {k for k, v in stats["odd_dist"].items() if v / d >= min_share}
        self.ok_low = {k for k, v in stats["low_dist"].items() if v / d >= min_share}
        self.sum_lo, self.sum_hi = stats["sum_p05"], stats["sum_p95"]
        self.max_consec = max(k for k, v in stats["consecutive_dist"].items() if v / d >= min_share)
        self._odd_peak = max(stats["odd_dist"].values())
        self._low_peak = max(stats["low_dist"].values())

    # ---- profile ---------------------------------------------------------
    def passes_profile(self, combo: tuple[int, ...]) -> bool:
        odd = sum(n % 2 for n in combo)
        low = sum(n <= self.stats["half"] for n in combo)
        return (odd in self.ok_odd and low in self.ok_low
                and self.sum_lo <= sum(combo) <= self.sum_hi
                and count_consecutive(combo) <= self.max_consec)

    def profile_fit(self, combo: tuple[int, ...]) -> float:
        """
        0-100 'how typical is this combo of past draws'. NOT a win probability:
        every valid combination has the same 1-in-total_possible chance.
        """
        s = self.stats
        z = (sum(combo) - s["sum_mean"]) / s["sum_sd"]
        sum_fit = math.exp(-0.5 * z * z)
        odd_fit = s["odd_dist"].get(sum(n % 2 for n in combo), 0) / self._odd_peak
        low_fit = s["low_dist"].get(sum(n <= s["half"] for n in combo), 0) / self._low_peak
        wmax = max(self.weights)
        freq_fit = sum(self.weights[n - 1] for n in combo) / (PICK_SIZE * wmax)
        return round(100 * (sum_fit + odd_fit + low_fit + freq_fit) / 4, 1)

    # ---- generation ------------------------------------------------------
    def _draw_one(self) -> tuple[int, ...]:
        picked: set[int] = set()
        while len(picked) < PICK_SIZE:
            picked.add(self.rng.choices(self.numbers, self.weights)[0])
        return tuple(sorted(picked))

    def generate(self, count: int = 5, max_attempts: int = 200_000, verify_file: bool = True):
        """Return `count` unique, history-safe combos as (combo, fit) ranked by fit."""
        results: dict[tuple[int, ...], float] = {}
        attempts = rejected_dup = rejected_profile = 0
        while len(results) < count:
            attempts += 1
            if attempts > max_attempts:
                raise RuntimeError(
                    "Could not find enough unique combinations. Relax the profile "
                    "(lower min_share) or reduce count.")
            combo = self._draw_one()
            if combo in self.h.known or combo in results:   # duplicate guard
                rejected_dup += 1
                continue
            if not self.passes_profile(combo):
                rejected_profile += 1
                continue
            results[combo] = self.profile_fit(combo)

        ranked = sorted(results.items(), key=lambda kv: kv[1], reverse=True)

        # Final validation: independent re-read of the CSV file
        collisions = self.h.verify_against_file([c for c, _ in ranked]) if verify_file else []
        if collisions:
            raise RuntimeError(f"Validation failed, duplicates found: {collisions}")

        self.last_run = {"attempts": attempts, "rejected_duplicate": rejected_dup,
                         "rejected_profile": rejected_profile}
        return ranked


# --------------------------------------------------------------------------
# Honest test: does the generator beat pure random on past draws?
# --------------------------------------------------------------------------
def backtest(history: LottoHistory, start: int = 800, per_draw: int = 20, seed: int = 1) -> dict:
    """
    Walk forward: for each draw after `start`, build a generator from ONLY the
    earlier draws, make `per_draw` combos, and count matching numbers against
    the real draw. Compare with pure-random combos.
    """
    rng = random.Random(seed)
    combos = [c for _, c in history.draws]
    model_hits = random_hits = model_3 = random_3 = trials = 0
    for i in range(start, len(combos)):
        sub = LottoHistory.__new__(LottoHistory)          # lightweight past-only view
        sub.__dict__.update(history.__dict__)
        sub.draws = history.draws[:i]
        sub.known = {c for _, c in sub.draws}
        gen = Generator(sub, seed=rng.randint(0, 10**9))
        actual = set(combos[i])
        for combo, _ in gen.generate(per_draw, verify_file=False):
            m = len(actual & set(combo)); model_hits += m; model_3 += m >= 3
            trials += 1
        for _ in range(per_draw):
            r = rng.sample(range(1, history.max_num + 1), PICK_SIZE)
            m = len(actual & set(r)); random_hits += m; random_3 += m >= 3
    return {"trials": trials,
            "model_avg_matches": model_hits / trials,
            "random_avg_matches": random_hits / trials,
            "model_share_3plus": model_3 / trials,
            "random_share_3plus": random_3 / trials,
            "theory_avg_matches": PICK_SIZE * PICK_SIZE / history.max_num}


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------
def main() -> None:
    ap = argparse.ArgumentParser(description="PCSO lotto history analysis and duplicate-safe generator")
    ap.add_argument("csv")
    ap.add_argument("--game", default="6/45", choices=list(GAME_RANGES))
    ap.add_argument("--count", type=int, default=5)
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--backtest", action="store_true")
    args = ap.parse_args()

    h = LottoHistory(args.csv, args.game)
    r, a = h.report, h.analyze()
    print(f"Game {args.game}: {r.rows_used} draws used "
          f"({r.first_date:%Y-%m-%d} to {r.last_date:%Y-%m-%d}), "
          f"{len(r.rows_skipped)} skipped, {len(r.exact_duplicate_rows)} repeated rows ignored")
    print(f"Frequency chi-square p = {a['chi2_p']:.3f} "
          f"({'no significant bias' if a['chi2_p'] > 0.05 else 'possible bias'})")
    print(f"Sum: mean {a['sum_mean']:.1f}, sd {a['sum_sd']:.1f} | "
          f"odd split {a['odd_dist']}")

    gen = Generator(h, seed=args.seed)
    print("\nGenerated combinations (ranked by profile fit, not by win chance):")
    for i, (combo, fit) in enumerate(gen.generate(args.count), 1):
        print(f"  {i}. {'-'.join(f'{n:02d}' for n in combo)}   fit {fit}")
    print(f"\nValidation: all {args.count} checked against {len(h.known)} unique "
          f"historical combinations and re-read from the CSV file: 0 duplicates.")
    print(f"Run stats: {gen.last_run}")

    if args.backtest:
        bt = backtest(h)
        print("\nBacktest (walk-forward):")
        for k, v in bt.items():
            print(f"  {k}: {v:.4f}" if isinstance(v, float) else f"  {k}: {v}")


if __name__ == "__main__":
    main()
