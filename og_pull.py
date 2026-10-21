#!/usr/bin/env python3
"""
og_pull.py — Offensive guard data pull for the Upper Edge OG breakdown deck.

Pulls, for every NFL team (or one team):
  * Current LG / RG depth chart (plus the starting center for context)
  * Roster status (active, IR, practice squad) and age
  * 2026 offensive snaps to date
  * Active contract: type, length, total value, guarantees, APY
  * Year-by-year base salary / cap hit for the chart years, in deck format

Data source: nflverse (via the nflreadpy package). Depth charts and rosters
are refreshed daily. Contract figures come from OverTheCap, NOT Spotrac —
spot-check dollar figures against Spotrac before they go on a slide.

Usage:
    pip install nflreadpy polars openpyxl pandas
    python og_pull.py                 # all 32 teams -> og_pull.xlsx
    python og_pull.py --team ARI      # one team, also printed to screen
    python og_pull.py --season 2026 --years 2026 2027 2028 --out guards.xlsx
"""

import argparse
from datetime import date

import nflreadpy as nfl
import polars as pl

# nflverse depth charts use abbreviations; OverTheCap contracts use nicknames.
TEAM_NICKNAMES = {
    "ARI": "Cardinals", "ATL": "Falcons", "BAL": "Ravens", "BUF": "Bills",
    "CAR": "Panthers", "CHI": "Bears", "CIN": "Bengals", "CLE": "Browns",
    "DAL": "Cowboys", "DEN": "Broncos", "DET": "Lions", "GB": "Packers",
    "HOU": "Texans", "IND": "Colts", "JAX": "Jaguars", "KC": "Chiefs",
    "LV": "Raiders", "LAC": "Chargers", "LA": "Rams", "MIA": "Dolphins",
    "MIN": "Vikings", "NE": "Patriots", "NO": "Saints", "NYG": "Giants",
    "NYJ": "Jets", "PHI": "Eagles", "PIT": "Steelers", "SF": "49ers",
    "SEA": "Seahawks", "TB": "Buccaneers", "TEN": "Titans", "WAS": "Commanders",
}

STATUS_LABELS = {
    "ACT": "Active", "RES": "IR/Reserve", "DEV": "Practice Squad",
    "INA": "Inactive", "SUS": "Suspended", "PUP": "PUP", "NON": "NFI",
}

POSITIONS = ["LG", "RG", "C"]  # C is pulled for context only (rank 1)


def money(m):
    """OTC figures are in millions. 1.215 -> '$1.215M', 0.885 -> '$885K'."""
    if m is None:
        return "?"
    if m < 1:
        return f"${round(m * 1000):,}K"
    return f"${m:.3f}".rstrip("0").rstrip(".") + "M"


def age_on(birth, today):
    if birth is None:
        return None
    return today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))


def load_data(season):
    print("Loading depth charts, rosters, snaps and contracts from nflverse...")
    depth = nfl.load_depth_charts(seasons=season)
    latest = depth["dt"].max()
    depth = depth.filter(pl.col("dt") == latest)

    rosters = nfl.load_rosters_weekly(seasons=season)
    # Most recent roster entry per player. Teams on bye have no row for the
    # current week, so filtering to max(week) would drop them.
    rosters = rosters.sort("week").group_by("gsis_id").last()

    snaps = nfl.load_snap_counts(seasons=season)
    snaps = (
        snaps.filter(pl.col("game_type") == "REG")
        .group_by(["pfr_player_id", "player", "team"])
        .agg(
            pl.col("offense_snaps").sum().alias("off_snaps"),
            (pl.col("offense_pct") >= 0.5).sum().alias("games_50pct"),
        )
    )

    contracts = nfl.load_contracts().filter(pl.col("is_active") == True)
    return depth, latest, rosters, snaps, contracts


def contract_cells(con_row, years):
    """Return (summary dict, {year: chart text}) for one active contract."""
    if con_row is None:
        return {}, {y: "No contract found" for y in years}

    end_year = con_row["year_signed"] + con_row["years"] - 1
    history = {int(s["year"]): s for s in (con_row["season_history"] or [])
               if str(s.get("year", "")).isdigit()}   # skip the "Total" summary row

    cells = {}
    fa_written = False
    for y in years:
        if y <= end_year:
            s = history.get(y)
            if s and s.get("cap_number") is not None:
                cells[y] = f"{money(s.get('base_salary'))}/{money(s.get('cap_number'))}"
            else:
                cells[y] = "Under contract (no breakdown)"
        elif not fa_written:
            cells[y] = "Free Agent"      # first year after expiry only
            fa_written = True
        else:
            cells[y] = ""                # trailing FA entries dropped (deck standard)

    last = (con_row["contract_history"] or [{}])[-1]
    summary = {
        "contract_type": last.get("contract_type"),
        "signed": con_row["year_signed"],
        "length_yrs": con_row["years"],
        "through": end_year,
        "total_value": money(con_row["value"]),
        "guaranteed": money(con_row["guaranteed"]),
        "apy": money(con_row["apy"]),
    }
    return summary, cells


def build(team_filter, season, years):
    today = date.today()
    depth, latest, rosters, snaps, contracts = load_data(season)

    # gsis_id -> pfr id for the snap join
    ids = nfl.load_players().select(["gsis_id", "pfr_id"])

    roster_info = rosters.select(
        ["gsis_id", "team", "status", "birth_date", "entry_year", "draft_number"]
    )

    teams = [team_filter] if team_filter else sorted(TEAM_NICKNAMES)
    rows = []
    for team in teams:
        d = depth.filter(
            (pl.col("team") == team) & pl.col("pos_abb").is_in(POSITIONS)
        ).sort(["pos_abb", "pos_rank"])

        for r in d.iter_rows(named=True):
            # Center: context only, starter only
            if r["pos_abb"] == "C" and r["pos_rank"] != 1:
                continue

            gid = r["gsis_id"]
            ri = roster_info.filter(pl.col("gsis_id") == gid)
            ri = ri.row(0, named=True) if ri.height else {}

            pfr = ids.filter(pl.col("gsis_id") == gid)
            pfr = pfr["pfr_id"][0] if pfr.height else None
            sn = snaps.filter(pl.col("pfr_player_id") == pfr) if pfr else snaps.clear()
            if not sn.height:  # fall back to name + team when the id link is missing
                sn = snaps.filter((pl.col("player") == r["player_name"]) & (pl.col("team") == team))
            sn = {"off_snaps": sn["off_snaps"].sum(), "games_50pct": sn["games_50pct"].sum()} \
                if sn.height else {}

            con = contracts.filter(pl.col("gsis_id") == gid).sort("year_signed", descending=True)
            con = con.row(0, named=True) if con.height else None
            summary, cells = contract_cells(con, years)

            age = age_on(ri.get("birth_date"), today)
            chart_line = f"{r['player_name']} ({age if age is not None else '?'})- " + \
                "     ".join(f"{y}={cells[y]}" for y in years if cells[y])

            rows.append({
                "team": team,
                "pos": r["pos_abb"],
                "depth": r["pos_rank"],
                "player": r["player_name"],
                "age": age,
                "status": STATUS_LABELS.get(ri.get("status"), ri.get("status") or "Not on roster"),
                "entry_year": ri.get("entry_year"),
                "draft_pick": ri.get("draft_number"),
                f"{season}_off_snaps": sn.get("off_snaps", 0),
                f"{season}_games_50pct": sn.get("games_50pct", 0),
                **summary,
                **{str(y): cells[y] for y in years},
                "chart_line": chart_line,
                "contract_otc_team": con["team"] if con else None,
            })

    df = pl.DataFrame(rows)

    # Flags worth a human look before anything goes on a slide
    flags = []
    for row in df.iter_rows(named=True):
        f = []
        if row["status"] != "Active":
            f.append(row["status"])
        if row["contract_otc_team"] and row["contract_otc_team"] != TEAM_NICKNAMES[row["team"]]:
            f.append(f"contract listed with {row['contract_otc_team']}")
        if row.get("through") is None:
            f.append("no contract record")
        elif any(row[str(y)] == "Under contract (no breakdown)" for y in years):
            f.append("no year-by-year $ — fill from Spotrac")
        flags.append("; ".join(f))
    df = df.with_columns(pl.Series("flags", flags)).drop("contract_otc_team")
    return df, latest


def main():
    p = argparse.ArgumentParser(description="Pull OG depth charts and contracts for the deck.")
    p.add_argument("--team", help="Team abbreviation, e.g. ARI. Omit for all 32.")
    p.add_argument("--season", type=int, default=2026)
    p.add_argument("--years", type=int, nargs="+", default=[2026, 2027, 2028],
                   help="Contract chart columns (default 2026 2027 2028)")
    p.add_argument("--out", default="og_pull.xlsx", help="Output .xlsx (or .csv)")
    a = p.parse_args()

    team = a.team.upper() if a.team else None
    if team and team not in TEAM_NICKNAMES:
        raise SystemExit(f"Unknown team '{a.team}'. Use one of: {', '.join(sorted(TEAM_NICKNAMES))}")

    df, latest = build(team, a.season, a.years)

    if a.out.endswith(".csv"):
        df.write_csv(a.out)
    else:
        import pandas as pd
        with pd.ExcelWriter(a.out, engine="openpyxl") as xw:
            pdf = pd.DataFrame(df.to_dicts(), columns=df.columns)
            pdf.to_excel(xw, sheet_name="Guards", index=False)
            ws = xw.sheets["Guards"]
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 70)
            ws.freeze_panes = "E2"
            pd.DataFrame({
                "note": [
                    f"Depth charts as of {latest} (nflverse/ESPN)",
                    "Contracts from OverTheCap via nflverse. Spot-check $ figures against Spotrac.",
                    "Chart cells are base salary / cap hit. 'Free Agent' shown once after expiry; later years left blank.",
                    "C rows are the starting center, included for context only.",
                    "PFF grades and line rankings are not included (no public API) — add manually.",
                ]
            }).to_excel(xw, sheet_name="Read me", index=False)

    print(f"\nDepth charts as of {latest}. Wrote {df.height} rows to {a.out}")
    if team:
        with pl.Config(tbl_rows=50, tbl_cols=30, tbl_width_chars=250, fmt_str_lengths=120):
            print(df.select(["pos", "depth", "player", "age", "status", "through",
                             "total_value", "guaranteed", "flags"]))
            print()
            for line in df["chart_line"]:
                print(line)


if __name__ == "__main__":
    main()