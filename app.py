"""
app.py  -  PCSO Lotto Generator (all games)

Reads draw history from every .csv / .xlsx file inside the "data" folder
(one file per game is fine).
Every game found in the file (LOTTO GAME column, or the sheet name in an
Excel workbook) uses lotto_engine, so a new combination can never match a
past winning combination. A game with no history in the file falls back to
the "paste your past draws" box.
"""
import json
import os
import random
from collections import Counter
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from flask import Flask, render_template, request, send_from_directory

from lotto_engine import (GAME_RANGES, Generator, LottoHistory, games_in_file,
                          parse_combination, parse_date)

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # works no matter where the server starts
DATA_DIR = os.path.join(BASE_DIR, "data")
ADD_PIN = ""   # optional: set e.g. "4827" so only you can save new draws on the public site
PICK_FILE = os.path.join(BASE_DIR, "last_pick.json")
LOCK_HOURS = 24
PH = ZoneInfo("Asia/Manila")
FMT = "%b %d, %Y %I:%M %p"

GAMES = GAME_RANGES
GAME_NAMES = {"6/42": "Lotto 6/42", "6/45": "Megalotto 6/45", "6/49": "Superlotto 6/49",
              "6/55": "Grand Lotto 6/55", "6/58": "Ultra Lotto 6/58"}
DEFAULT_GAME = "6/58"


# ---------- history file ----------
def history_path():
    """The data folder, if it exists and has at least one csv/xlsx file."""
    if os.path.isdir(DATA_DIR) and any(
            f.lower().endswith((".csv", ".xlsx", ".xlsm")) for f in os.listdir(DATA_DIR)):
        return DATA_DIR
    return None


def available_games():
    """{'6/45': 1615, ...} for every game that has rows in the history file."""
    path = history_path()
    if not path:
        return {}
    try:
        return games_in_file(path)
    except Exception:
        return {}


# ---------- fallback (no history in file): paste your own draws ----------
def parse_draws(text, max_num):
    draws = []
    for i, line in enumerate(text.strip().splitlines(), start=1):
        parts = line.replace(",", " ").split()
        if not parts:
            continue
        try:
            nums = [int(p) for p in parts]
        except ValueError:
            raise ValueError(f"Line {i}: only numbers are allowed.")
        if len(nums) != 6 or len(set(nums)) != 6:
            raise ValueError(f"Line {i}: needs 6 different numbers.")
        if any(n < 1 or n > max_num for n in nums):
            raise ValueError(f"Line {i}: numbers must be from 1 to {max_num}.")
        draws.append(sorted(nums))
    if not draws:
        raise ValueError("Please paste at least one draw.")
    return draws


def generate_from_paste(draws, method, max_num):
    numbers = list(range(1, max_num + 1))
    counts = Counter(n for draw in draws for n in draw)
    if method == "hot":
        weights = [counts[n] + 1 for n in numbers]
    elif method == "cold":
        top = max(counts.values())
        weights = [top - counts[n] + 1 for n in numbers]
    else:
        weights = [1] * max_num
    known = {tuple(d) for d in draws}          # never repeat a pasted draw
    while True:
        combo = set()
        while len(combo) < 6:
            combo.add(random.choices(numbers, weights)[0])
        combo = tuple(sorted(combo))
        if combo not in known:
            return list(combo)


# ---------- 24 hour lock ----------
def load_picks():
    if os.path.exists(PICK_FILE):
        with open(PICK_FILE) as f:
            return json.load(f)
    return {}


def save_picks(picks):
    with open(PICK_FILE, "w") as f:
        json.dump(picks, f)


def time_info(pick):
    created = datetime.fromisoformat(pick["created_at"])
    unlock = created + timedelta(hours=LOCK_HOURS)
    left = unlock - datetime.now(timezone.utc)
    locked = left.total_seconds() > 0
    secs = max(int(left.total_seconds()), 0)
    return (locked, created.astimezone(PH).strftime(FMT),
            unlock.astimezone(PH).strftime(FMT), f"{secs // 3600}h {(secs % 3600) // 60}m")


# ---------- latest draw ----------
def latest_draw(game):
    """(date text, combo) of the newest draw saved for this game, or None."""
    path = history_path()
    if not path:
        return None
    draws = [d for d in LottoHistory(path, game).draws if d[0]]
    if not draws:
        return None
    date, combo = draws[-1]
    return {"date": date.strftime("%b %d, %Y"), "combo": list(combo)}


def add_latest(game, numbers_text, date_text, pin=""):
    """Validate and save a new official draw. Returns a message; raises ValueError."""
    os.makedirs(DATA_DIR, exist_ok=True)
    if ADD_PIN and pin.strip() != ADD_PIN:
        raise ValueError("Wrong PIN.")
    try:
        combo = parse_combination(numbers_text)
    except ValueError:
        raise ValueError("Numbers must be digits like 05 12 23 34 40 41.")
    date = parse_date(date_text)
    if date is None:
        raise ValueError("Please pick the draw date.")
    history = LottoHistory(DATA_DIR, game)
    if any(d and d.date() == date.date() for d, _ in history.draws):
        raise ValueError(f"{game} already has a draw saved for {date:%b %d, %Y}.")
    seen_before = tuple(combo) in history.known
    history.add_result(combo, date)          # checks 6 different numbers in range, saves to data/NEW_DRAWS.csv
    text = "-".join(f"{n:02d}" for n in combo)
    note = " (note: this combination was drawn before)" if seen_before else ""
    return f"Saved {game} draw {text} for {date:%b %d, %Y}{note}."


# ---------- page ----------
@app.route("/", methods=["GET", "POST"])
def index():
    game = request.values.get("game", DEFAULT_GAME)
    if game not in GAMES:
        game = DEFAULT_GAME
    max_num = GAMES[game]

    in_file = available_games()
    has_history = game in in_file
    path = history_path()

    combo = error = notice = made = next_time = left = None
    draws_text = ""
    method = "hot"

    picks = load_picks()
    pick = picks.get(game)
    locked = False
    if pick:
        locked, made, next_time, left = time_info(pick)
        combo, method, draws_text = pick["combo"], pick["method"], pick.get("draws_text", "")

    if request.method == "POST" and request.form.get("action") == "add":
        try:
            notice = add_latest(game, request.form.get("numbers", ""), request.form.get("draw_date", ""),
                                request.form.get("pin", ""))
            in_file = available_games()
            has_history = game in in_file
            path = history_path()
        except ValueError as e:
            error = str(e)
    elif request.method == "POST":
        if locked:
            notice = f"Your next {game} combination unlocks in {left}."
        else:
            method = request.form.get("method", "hot")
            draws_text = "" if has_history else request.form.get("draws", "")
            combo = None
            try:
                if has_history:
                    history = LottoHistory(path, game)           # re-read every time
                    combo = list(Generator(history, method=method).generate(1)[0][0])
                else:
                    combo = generate_from_paste(parse_draws(draws_text, max_num), method, max_num)
                new_pick = {"created_at": datetime.now(timezone.utc).isoformat(),
                            "method": method, "combo": combo, "draws_text": draws_text}
                picks[game] = new_pick
                save_picks(picks)
                locked, made, next_time, left = time_info(new_pick)
                notice = "New combination generated."
            except (ValueError, RuntimeError) as e:
                error = str(e)
                if pick:
                    combo = pick["combo"]

    return render_template(
        "index.html", games=list(GAMES), game=game, game_name=GAME_NAMES[game],
        max_num=max_num, has_history=has_history, history_count=in_file.get(game, 0),
        in_file=in_file, latest=latest_draw(game) if has_history else None,
        need_pin=bool(ADD_PIN), today=datetime.now(PH).strftime("%Y-%m-%d"), combo=combo, error=error, notice=notice, draws_text=draws_text,
        method=method, locked=locked, made=made, next_time=next_time, left=left)


# ---------- installable app (PWA) ----------
@app.route("/manifest.webmanifest")
def manifest():
    return send_from_directory(os.path.join(BASE_DIR, "static"), "manifest.webmanifest",
                               mimetype="application/manifest+json")


@app.route("/sw.js")
def service_worker():
    # served from the site root so it can control every page
    resp = send_from_directory(os.path.join(BASE_DIR, "static"), "sw.js",
                               mimetype="application/javascript")
    resp.headers["Cache-Control"] = "no-cache"
    return resp
