#!/usr/bin/env python3
"""Youth Soccer League Team Allocator

Author: ChatGPT
Date: May 2025

Usage (source):
    python team_allocator.py

Packaging:
    pip install pandas openpyxl
    pyinstaller --onefile --noconsole team_allocator.py

The resulting `team_allocator.exe` can be run directly with no Python installation.
"""

import math
import os
import random
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from typing import Dict, List

import pandas as pd


RNG = random.Random()


def choose_num_teams(n_players: int, min_size: int) -> int:
    """Return an **even** number ≥ 4 that satisfies the examples in the spec.

    Strategy:
        • Require each team to have at least *min_size* players.
        • Allow team size to exceed *min_size* by **up to 5** before new teams are created.
    """
    if n_players < min_size * 4:
        raise ValueError(
            f"{n_players} players cannot fill the required minimum of four teams with {min_size} per team."
        )

    for teams in range(4, n_players + 1):
        if teams % 2:
            continue
        avg = n_players / teams
        if avg < min_size:
            continue
        if avg <= min_size + 5:  # matches the 30‑, 47‑ and 90‑player examples for min_size 7
            return teams

    # Fallback – should not occur with realistic inputs
    teams = math.ceil(n_players / min_size)
    if teams % 2:
        teams += 1
    return teams


def distribute_capacities(n_players: int, n_teams: int) -> List[int]:
    """Return a list with the capacity of each team.

    Extra players (n mod teams) are spread one‑per‑team across the first teams.
    Example: 30 players, 4 teams → [8, 8, 7, 7] (order randomised later).
    """
    base = n_players // n_teams
    extras = n_players % n_teams
    capacities = [base + 1] * extras + [base] * (n_teams - extras)
    RNG.shuffle(capacities)  # randomise which teams are the larger ones
    return capacities


def group_girls(girl_idx: List[int]) -> List[List[int]]:
    """Return groups of 2–3 girls (to keep friends together)."""
    groups: List[List[int]] = []
    i = 0
    while i < len(girl_idx):
        remaining = len(girl_idx) - i
        if remaining == 1:
            # add the last single girl to the previous group (which must be ≥ 2)
            groups[-1].append(girl_idx[i])
            i += 1
        elif remaining == 2 or remaining % 3 == 0:
            groups.append(girl_idx[i : i + 2])
            i += 2
        else:
            groups.append(girl_idx[i : i + 3])
            i += 3
    return groups


def allocate_players(
    df: pd.DataFrame, capacities: List[int]
) -> List[List[int]]:
    """Return a list where each entry is a list of row indices for that team."""

    n_teams = len(capacities)
    teams: List[List[int]] = [[] for _ in range(n_teams)]
    remaining_capacity = capacities[:]

    # ---------- Step 1 – allocate girls in 2–3‑player pods ----------
    girls_idx = df[df["Gndr"].str.upper().str.startswith("F")].index.tolist()
    girl_groups = group_girls(girls_idx)
    team_cycle = RNG.sample(range(n_teams), k=n_teams)  # random starting order
    t_pointer = 0
    for grp in girl_groups:
        # advance to a team with enough free slots
        attempts = 0
        while remaining_capacity[team_cycle[t_pointer]] < len(grp) and attempts < n_teams:
            t_pointer = (t_pointer + 1) % n_teams
            attempts += 1
        if attempts == n_teams:
            # no single team fits; fall back to next phase
            break
        tidx = team_cycle[t_pointer]
        teams[tidx].extend(grp)
        remaining_capacity[tidx] -= len(grp)
        t_pointer = (t_pointer + 1) % n_teams

    # ---------- Step 2 – allocate by school ----------
    by_school: Dict[str, List[int]] = {}
    for idx, school in df["Answer"].fillna("Unknown").items():
        by_school.setdefault(school, []).append(idx)
    # sort largest school first so they are forced to split
    schools = sorted(by_school.items(), key=lambda kv: len(kv[1]), reverse=True)

    # repeatedly loop through schools, sprinkling one player on each pass
    exhausted = False
    while not exhausted:
        exhausted = True
        for school, idx_list in schools:
            if not idx_list:
                continue
            exhausted = False
            # pick the player (pop for O(1))
            p = idx_list.pop()
            # choose team that still has room and currently fewest from this school
            best_team = None
            min_school_count = math.inf
            for tidx in range(n_teams):
                if remaining_capacity[tidx] == 0:
                    continue
                school_count = sum(
                    1 for i in teams[tidx] if df.loc[i, "Answer"] == school
                )
                if school_count < min_school_count:
                    min_school_count = school_count
                    best_team = tidx
                    if school_count == 0:
                        break  # cannot do better
            if best_team is None:
                # should not happen; if it does, toss player at random free slot
                free_teams = [i for i, cap in enumerate(remaining_capacity) if cap > 0]
                best_team = RNG.choice(free_teams)
            teams[best_team].append(p)
            remaining_capacity[best_team] -= 1

    return teams


def save_to_csv(df: pd.DataFrame, teams: List[List[int]], in_path: str) -> str:
    # Map index → team name
    mapping = {}
    for t, members in enumerate(teams, 1):
        for idx in members:
            mapping[idx] = f"Team {t}"
    df_out = df.copy()
    df_out.insert(0, "Team", df_out.index.map(mapping).fillna("Unassigned"))
    df_out = df_out.sort_values("Team")
    out_path = os.path.join(
        os.path.dirname(in_path),
        os.path.splitext(os.path.basename(in_path))[0] + "_teams.csv",
    )
    df_out.to_csv(out_path, index=False)
    return out_path


def main():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Soccer Team Allocator",
        "Select the enrolment Excel file (with headers) to begin.",
    )
    in_path = filedialog.askopenfilename(
        title="Choose enrolment.xlsx",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
    )
    if not in_path:
        sys.exit()

    try:
        df = pd.read_excel(in_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read the Excel file:\n{e}")
        sys.exit(1)

    # --- Step 1 – filter to the school question ---
    df = df[df["Custom Question"] == "What is your school?"].reset_index(drop=True)

    if df.empty:
        messagebox.showerror(
            "Error", "No rows where ‘Custom Question’ = ‘What is your school?’"
        )
        sys.exit(1)

    # --- prompt for minimum team size ---
    min_size = simpledialog.askinteger(
        "Minimum team size",
        "Enter the minimum team size (5–7 recommended):",
        initialvalue=7,
        minvalue=3,
        maxvalue=15,
    )
    if not min_size:
        sys.exit()

    players = len(df)

    try:
        n_teams = choose_num_teams(players, min_size)
    except ValueError as e:
        messagebox.showerror("Not enough players", str(e))
        sys.exit(1)

    capacities = distribute_capacities(players, n_teams)

    while True:
        teams = allocate_players(df, capacities)
        avg = players / n_teams

        # --- confirmation window with 3 buttons ---
        top = tk.Toplevel()
        top.title("Confirm Teams")
        tk.Label(
            top,
            text=f"{n_teams} teams created with an average team size of {avg:.1f}.",
            padx=20,
            pady=10,
        ).pack()

        choice = tk.StringVar()

        def _set(val: str):
            choice.set(val)
            top.destroy()

        tk.Button(top, text="SAVE", width=10, command=lambda: _set("SAVE")).pack(
            side="left", padx=10, pady=10
        )
        tk.Button(
            top, text="REGENERATE", width=10, command=lambda: _set("REGEN")
        ).pack(side="left", padx=10, pady=10)
        tk.Button(top, text="QUIT", width=10, command=lambda: _set("QUIT")).pack(
            side="left", padx=10, pady=10
        )
        top.wait_variable(choice)

        if choice.get() == "QUIT":
            sys.exit()
        elif choice.get() == "REGEN":
            RNG.shuffle(capacities)  # reshuffle capacity assignment
            continue
        elif choice.get() == "SAVE":
            try:
                out_csv = save_to_csv(df, teams, in_path)
                messagebox.showinfo(
                    "Saved", f"Team allocation saved to:\n{out_csv}"
                )
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV:\n{e}")
            break


if __name__ == "__main__":
    main()
