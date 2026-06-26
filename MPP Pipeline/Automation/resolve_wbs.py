"""
Stage D — WBS Resolver and Grouping
====================================
Applies the structural resolver to the Buyout WBS and produces the six
grouping fields per leaf activity: Section, Group, Phase, Category,
Sub-Category, Activity.

This is the analytically subtle stage. The WBS is NOT uniform:
  - Most packages sit at  Phase (L4) -> Material (L5)
  - A few "override" scopes sit directly at the phase level (L4), with a
    generic workflow summary occupying the L5 material slot.
A fixed outline-depth rule mislabels the overrides. The fix is structural:

    category(leaf) = the first ancestor A (climbing up from the leaf) whose
                     parent is a Group node or a Phase node.

Bottom-up traversal guarantees the material/trade wins over the phase
wrapper. A package under a phase resolves to the material; a package at the
phase level resolves to itself; generic workflow summaries never qualify.

Inputs:
  - project_config.json  (wbs_resolver section + schedule.buyout_task_id_range_end)
  - Stage C output        (output_root/stage_c/snapshots/<snapshot>.parquet)

Outputs (written to output_root/stage_d/):
  - buyout_grouping.parquet  — one row per buyout LEAF, with six grouping fields
  - buyout_packages.parquet  — one row per package (resolved category summary)
  - resolver_report.json     — node counts, bucket distribution, Harrison validation

Runs on a single snapshot (default: latest by filename) because the buyout
grouping is structural and Stage H analyses the current schedule.

No Java required — reads the canonical Stage C parquet.

Usage:
  python resolve_wbs.py [--config project_config.json] [--snapshot <stem>]
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


# Harrison reference distribution for the built-in validation check.
HARRISON_TARGETS = {
    "packages_total": 84,
    "leaves_total": 2151,
    "buckets": {
        ("Buyout Work", "Procurement"):    {"packages": 11, "leaves": 1561},
        ("Buyout Work", "Subcontracting"): {"packages": 29, "leaves": 485},
        ("Lead Time", "Procurement"):      {"packages": 11, "leaves": 46},
        ("Lead Time", "Subcontracting"):   {"packages": 33, "leaves": 59},
    },
}


# ---------------------------------------------------------------------------
# Config + snapshot selection
# ---------------------------------------------------------------------------

def load_config(config_path: str) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def find_snapshot(cfg: dict, snapshot_arg: str | None) -> Path:
    """Return the parquet path for the snapshot to resolve."""
    output_root = Path(cfg["paths"]["output_root"])
    snap_dir = output_root / "stage_c" / "snapshots"
    if not snap_dir.exists():
        raise FileNotFoundError(
            f"Stage C output not found at {snap_dir}. Run extract_snapshots.py first."
        )

    parquets = sorted(snap_dir.glob("*.parquet"), key=lambda p: p.stem)
    if not parquets:
        raise FileNotFoundError(f"No snapshot parquet files in {snap_dir}.")

    if snapshot_arg:
        match = snap_dir / f"{snapshot_arg}.parquet"
        if not match.exists():
            raise FileNotFoundError(
                f"Requested snapshot '{snapshot_arg}' not found in {snap_dir}.\n"
                f"Available: {[p.stem for p in parquets]}"
            )
        return match

    # Default: latest by sorted filename stem
    return parquets[-1]


# ---------------------------------------------------------------------------
# Outline parsing
# ---------------------------------------------------------------------------

def parse_segments(outline_number) -> list[str]:
    """Split an MS Project outline number ('1.1.2.3') into segments."""
    if outline_number is None:
        return []
    s = str(outline_number).strip()
    if s == "" or s.lower() == "none" or s.lower() == "nan":
        return []
    return s.split(".")


# ---------------------------------------------------------------------------
# Core resolver  (pure function — fully unit-testable)
# ---------------------------------------------------------------------------

def build_grouping(tasks_df: pd.DataFrame, cfg: dict):
    """
    Apply the structural resolver to the buyout-scoped tasks.

    Parameters
    ----------
    tasks_df : DataFrame
        One snapshot's tasks. Required columns:
        uid, name, outline_number, parent_uid, is_summary
    cfg : dict
        The project_config dict.

    Returns
    -------
    (grouping_df, packages_df, report) :
        grouping_df — one row per buyout leaf with the six grouping fields
        packages_df — one row per resolved package (category summary)
        report      — dict of node counts + bucket distribution
    """
    wr = cfg["wbs_resolver"]
    buyout_end_id = cfg["schedule"]["buyout_task_id_range_end"]

    # ── Derive depths from the configured outline prefixes ────────────────
    bw_segments = wr["buyout_work_outline_prefix"].split(".")   # e.g. ["1","1"]
    lt_segments = wr["lead_time_outline_prefix"].split(".")     # e.g. ["1","2"]
    proc_suffix = str(wr["procurement_outline_suffix"])          # "1"
    sub_suffix  = str(wr["subcontracting_outline_suffix"])       # "2"

    section_depth = len(bw_segments)        # 2
    group_depth   = section_depth + 1       # 3  (Procurement / Subcontracting)
    phase_depth   = section_depth + 2       # 4  (Phase wrapper / override material)
    group_idx     = section_depth           # 0-based index of the group segment (2)

    override_names = {
        o["task_name"].strip().lower() for o in wr.get("bucket_overrides", [])
    }

    # ── Scope to buyout tasks ─────────────────────────────────────────────
    df = tasks_df[tasks_df["uid"] <= buyout_end_id].copy()

    # ── Build a uid -> record map for O(1) climbs ─────────────────────────
    records = {}
    for row in df.itertuples(index=False):
        uid = int(row.uid)
        puid = getattr(row, "parent_uid")
        puid = int(puid) if pd.notna(puid) else None
        records[uid] = {
            "uid":        uid,
            "name":       str(row.name) if pd.notna(row.name) else "",
            "parent_uid": puid,
            "segments":   parse_segments(getattr(row, "outline_number")),
            "is_summary": bool(row.is_summary),
        }

    # ── Classify node types ───────────────────────────────────────────────
    # Group nodes:  summaries at group_depth (3 segments)  -> Procurement / Subcontracting
    # Phase nodes:  summaries at phase_depth (4 segments) whose name is NOT an override
    group_node_uids = set()
    phase_node_uids = set()
    for uid, rec in records.items():
        if not rec["is_summary"]:
            continue
        depth = len(rec["segments"])
        if depth == group_depth:
            group_node_uids.add(uid)
        elif depth == phase_depth and rec["name"].strip().lower() not in override_names:
            phase_node_uids.add(uid)

    boundary_uids = group_node_uids | phase_node_uids

    # ── Section / Group classifier (from a task's own outline segments) ───
    def classify_section_group(segments: list[str]):
        section = ""
        group = ""
        if segments[:section_depth] == bw_segments:
            section = "Buyout Work"
        elif segments[:section_depth] == lt_segments:
            section = "Lead Time"
        if len(segments) > group_idx:
            seg = segments[group_idx]
            if seg == proc_suffix:
                group = "Procurement"
            elif seg == sub_suffix:
                group = "Subcontracting"
        return section, group

    # ── Resolve the category (package) for one leaf via the climb ─────────
    def resolve_category(leaf_uid: int):
        """
        Climb the ancestors of the leaf. Return the first ancestor A whose
        parent is a boundary (group or phase) node. That A is the package.
        """
        leaf = records[leaf_uid]
        a_uid = leaf["parent_uid"]            # start at the immediate parent
        visited = set()
        while a_uid is not None and a_uid in records and a_uid not in visited:
            visited.add(a_uid)
            a = records[a_uid]
            if a["parent_uid"] in boundary_uids:
                return a_uid                  # A sits directly under a group/phase node
            a_uid = a["parent_uid"]
        return None                           # unresolved (malformed branch)

    # ── Walk every leaf and assemble the six grouping fields ──────────────
    grouping_rows = []
    for uid, rec in records.items():
        if rec["is_summary"]:
            continue  # leaves only

        section, group = classify_section_group(rec["segments"])
        category_uid = resolve_category(uid)

        if category_uid is not None:
            cat = records[category_uid]
            category_name = cat["name"]
            cat_parent_uid = cat["parent_uid"]
            # Phase = the category's parent IF that parent is a phase node;
            # blank when the category is a stand-alone (parent is a group node).
            if cat_parent_uid in phase_node_uids:
                phase_name = records[cat_parent_uid]["name"]
            else:
                phase_name = ""
        else:
            category_name = ""
            phase_name = ""

        # Sub-Category = the leaf's immediate parent summary name
        parent_uid = rec["parent_uid"]
        sub_category = records[parent_uid]["name"] if parent_uid in records else ""

        grouping_rows.append({
            "uid":          uid,
            "section":      section,
            "group":        group,
            "phase":        phase_name,
            "category":     category_name,
            "sub_category": sub_category,
            "activity":     rec["name"],
            "package_uid":  category_uid if category_uid is not None else pd.NA,
            "is_lead_time": (section == "Lead Time"),
        })

    grouping_df = pd.DataFrame(grouping_rows)

    # ── Structural package set (independent of leaf resolution) ───────────
    # A package is a summary whose parent is a group/phase node and which is
    # not itself a phase node (phase wrappers are not packages).
    structural_packages = set()
    for uid, rec in records.items():
        if (rec["is_summary"]
                and rec["parent_uid"] in boundary_uids
                and uid not in phase_node_uids):
            structural_packages.add(uid)

    # ── Package-level table (from the leaves that resolved to each package)─
    pkg_rows = []
    if not grouping_df.empty:
        resolved = grouping_df.dropna(subset=["package_uid"])
        for pkg_uid, grp in resolved.groupby("package_uid"):
            first = grp.iloc[0]
            pkg_rows.append({
                "package_uid":  int(pkg_uid),
                "package_name": records[int(pkg_uid)]["name"],
                "section":      first["section"],
                "group":        first["group"],
                "phase":        first["phase"],
                "leaf_count":   len(grp),
            })
    packages_df = pd.DataFrame(pkg_rows)

    # ── Bucket distribution (Section x Group) ─────────────────────────────
    bucket_dist = {}
    if not grouping_df.empty:
        for (sec, grp), sub in grouping_df.groupby(["section", "group"]):
            if sec == "" or grp == "":
                continue
            pkgs = sub["package_uid"].dropna().nunique()
            bucket_dist[f"{sec} / {grp}"] = {
                "packages": int(pkgs),
                "leaves":   int(len(sub)),
            }

    report = {
        "buyout_task_count":        len(df),
        "summary_count":            int(df["is_summary"].sum()),
        "leaf_count":               int((~df["is_summary"]).sum()),
        "group_node_count":         len(group_node_uids),
        "phase_node_count":         len(phase_node_uids),
        "structural_package_count": len(structural_packages),
        "resolved_category_count":  int(grouping_df["package_uid"].dropna().nunique())
                                    if not grouping_df.empty else 0,
        "unresolved_leaf_count":    int(grouping_df["package_uid"].isna().sum())
                                    if not grouping_df.empty else 0,
        "bucket_distribution":      bucket_dist,
        "override_check":           _check_overrides(records, structural_packages, wr),
    }

    return grouping_df, packages_df, report


def _check_overrides(records, structural_packages, wr):
    """Confirm every configured override material was found as a package."""
    pkg_names = {records[u]["name"].strip().lower() for u in structural_packages}
    result = []
    for o in wr.get("bucket_overrides", []):
        name = o["task_name"]
        result.append({
            "task_name": name,
            "found_as_package": name.strip().lower() in pkg_names,
        })
    return result


# ---------------------------------------------------------------------------
# Validation against Harrison reference
# ---------------------------------------------------------------------------

def print_harrison_validation(report: dict):
    """Print the computed distribution alongside the Harrison targets."""
    print("\n  Harrison reference validation")
    print("  " + "-" * 70)
    print(f"  {'Bucket':<34}{'Packages (got/exp)':<20}{'Leaves (got/exp)':<18}")
    print("  " + "-" * 70)

    dist = report["bucket_distribution"]
    total_pkg_got = total_pkg_exp = 0
    total_leaf_got = total_leaf_exp = 0
    all_match = True

    for (sec, grp), tgt in HARRISON_TARGETS["buckets"].items():
        key = f"{sec} / {grp}"
        got = dist.get(key, {"packages": 0, "leaves": 0})
        pkg_got, pkg_exp = got["packages"], tgt["packages"]
        leaf_got, leaf_exp = got["leaves"], tgt["leaves"]
        total_pkg_got += pkg_got;   total_pkg_exp += pkg_exp
        total_leaf_got += leaf_got; total_leaf_exp += leaf_exp
        ok = (pkg_got == pkg_exp and leaf_got == leaf_exp)
        all_match &= ok
        flag = "OK " if ok else "XX "
        print(f"  {flag}{key:<31}{pkg_got:>4} / {pkg_exp:<13}{leaf_got:>5} / {leaf_exp}")

    print("  " + "-" * 70)
    print(f"     {'TOTAL':<31}{total_pkg_got:>4} / {total_pkg_exp:<13}"
          f"{total_leaf_got:>5} / {total_leaf_exp}")
    print("  " + "-" * 70)
    if all_match:
        print("  All buckets match the Harrison reference exactly.")
    else:
        print("  Distribution differs from Harrison reference (expected for a different project,")
        print("  or a signal to investigate the WBS structure / bucket_overrides config).")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Stage D — WBS resolver and grouping")
    parser.add_argument("--config", default="project_config.json")
    parser.add_argument("--snapshot", default=None,
                        help="Snapshot stem to resolve (default: latest)")
    args = parser.parse_args()

    cfg = load_config(args.config)
    project_name = cfg["project"]["name"]
    output_root = Path(cfg["paths"]["output_root"])
    stage_dir = output_root / "stage_d"
    stage_dir.mkdir(parents=True, exist_ok=True)

    snapshot_path = find_snapshot(cfg, args.snapshot)

    print(f"\n{'='*60}")
    print(f"  Stage D — WBS Resolver and Grouping")
    print(f"  Project  : {project_name}")
    print(f"  Snapshot : {snapshot_path.stem}")
    print(f"  Output   : {stage_dir}")
    print(f"{'='*60}\n")

    tasks_df = pd.read_parquet(snapshot_path)

    grouping_df, packages_df, report = build_grouping(tasks_df, cfg)

    # ── Persist ───────────────────────────────────────────────────────────
    grouping_out = stage_dir / "buyout_grouping.parquet"
    packages_out = stage_dir / "buyout_packages.parquet"
    grouping_df.to_parquet(grouping_out, index=False)
    packages_df.to_parquet(packages_out, index=False)

    report["generated"] = datetime.now().isoformat(timespec="seconds")
    report["project"] = project_name
    report["snapshot"] = snapshot_path.stem
    report_out = stage_dir / "resolver_report.json"
    with open(report_out, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, default=str)

    # ── Console summary ───────────────────────────────────────────────────
    print(f"  Buyout tasks      : {report['buyout_task_count']:>6}  "
          f"(summary {report['summary_count']}, leaf {report['leaf_count']})")
    print(f"  Group nodes       : {report['group_node_count']:>6}")
    print(f"  Phase nodes       : {report['phase_node_count']:>6}")
    print(f"  Packages (struct) : {report['structural_package_count']:>6}")
    print(f"  Packages (leaves) : {report['resolved_category_count']:>6}")
    print(f"  Unresolved leaves : {report['unresolved_leaf_count']:>6}")

    print("\n  Bucket distribution (Section / Group):")
    for bucket, vals in report["bucket_distribution"].items():
        print(f"    {bucket:<34}{vals['packages']:>4} packages   {vals['leaves']:>5} leaves")

    print("\n  Override check:")
    for chk in report["override_check"]:
        flag = "OK " if chk["found_as_package"] else "XX MISSING"
        print(f"    {flag} {chk['task_name']}")

    print_harrison_validation(report)

    print(f"\n{'='*60}")
    print(f"  Grouping : {grouping_out}")
    print(f"  Packages : {packages_out}")
    print(f"  Report   : {report_out}")
    print(f"{'='*60}\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
