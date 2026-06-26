"""
Stage C — Environment Setup & Data Extraction
==============================================
Reads every XML snapshot from the folder defined in project_config.json,
extracts the canonical task-level fields required by Stages D-H, and
persists each snapshot as a parquet file.

Outputs (written to <output_root>/stage_c/):
  snapshots/
    <snapshot_date>.parquet   — one file per snapshot, all tasks
  predecessors/
    <snapshot_date>.parquet   — predecessor relationships per snapshot
  extraction_report.json      — file count, field validation, UID stats,
                                ActualDuration == Duration check

Usage:
  python extract_snapshots.py [--config path/to/project_config.json]

Requirements:
  Python 3.12, JPype1, Java 21, mpxj (org.mpxj v16.4.0),
  pandas, pyarrow, numpy

Known traps (from playbook):
  - .xml extension files are actually native MPP binaries — MPXJ
    UniversalProjectReader handles this transparently.
  - Use org.mpxj namespace, NOT net.sf.mpxj (MPXJ v16.x).
  - Never name this file inspect.py — shadows stdlib inspect module
    and breaks JPype.
  - Duration unit conversion: hours / 8, weeks * 5 -> working days.
  - Java 21 required; older JVMs fail at startup.
"""

import argparse
import json
import os
import sys
import traceback
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------

def load_config(config_path: str) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# JVM startup
# ---------------------------------------------------------------------------

def start_jvm(mpxj_jar_path: str | None = None):
    """
    Start JPype JVM with the MPXJ jar on the classpath.
    If mpxj_jar_path is None, attempts to locate it via the mpxj
    Python package (pip install mpxj installs the jar automatically).
    """
    import jpype
    import jpype.imports

    if jpype.isJVMStarted():
        return

    # Locate MPXJ jar via the mpxj Python package if no explicit path given
    if mpxj_jar_path is None:
        try:
            import mpxj as mpxj_pkg
            jar = Path(mpxj_pkg.__file__).parent / "mpxj.jar"
            if not jar.exists():
                # Fallback: search site-packages for any mpxj jar
                import site
                for sp in site.getsitepackages():
                    candidates = list(Path(sp).rglob("mpxj*.jar"))
                    if candidates:
                        jar = candidates[0]
                        break
            mpxj_jar_path = str(jar)
        except ImportError:
            raise RuntimeError(
                "mpxj Python package not found. Install with: pip install mpxj\n"
                "Or pass --mpxj-jar /path/to/mpxj.jar explicitly."
            )

    print(f"  JVM classpath: {mpxj_jar_path}")
    jpype.startJVM(classpath=[mpxj_jar_path], convertStrings=True)
    print(f"  JVM started (Java {jpype.getJVMVersion()})")


# ---------------------------------------------------------------------------
# Duration conversion to working days
# ---------------------------------------------------------------------------

def to_working_days(duration_obj) -> float | None:
    """
    Convert an MPXJ Duration object to a float in working days.
    MPXJ stores durations with a TimeUnit; we normalise to days.
    Returns None if the duration object is None or zero-equivalent.
    """
    if duration_obj is None:
        return None

    try:
        # Get the numeric value and time unit
        value = float(duration_obj.getDuration())
        unit_str = str(duration_obj.getUnits()).upper()

        if unit_str in ("HOURS", "ELAPSED_HOURS"):
            return value / 8.0
        elif unit_str in ("DAYS", "ELAPSED_DAYS"):
            return value
        elif unit_str in ("WEEKS", "ELAPSED_WEEKS"):
            return value * 5.0
        elif unit_str in ("MINUTES", "ELAPSED_MINUTES"):
            return value / 480.0
        elif unit_str == "MONTHS":
            return value * 20.0  # approximate
        else:
            # Default: treat as days
            return value
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Date extraction helper
# ---------------------------------------------------------------------------

def to_pydate(mpxj_date) -> date | None:
    """Convert an MPXJ LocalDateTime / Date to a Python date."""
    if mpxj_date is None:
        return None
    try:
        # MPXJ v16 returns java.time.LocalDateTime for most date fields
        s = str(mpxj_date)
        # LocalDateTime.toString() -> "2026-04-09T00:00" or "2026-04-09"
        return datetime.fromisoformat(s[:10]).date()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Single-file extraction
# ---------------------------------------------------------------------------

def extract_file(xml_path: Path, snapshot_label: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Parse one MPP/XML file and return:
      tasks_df       — one row per task, all canonical fields
      preds_df       — one row per predecessor relationship
    """
    from org.mpxj.reader import UniversalProjectReader  # type: ignore

    reader = UniversalProjectReader()
    project = reader.read(str(xml_path))

    tasks_rows = []
    preds_rows = []

    for task in project.getTasks():
        uid = task.getUniqueID()
        if uid is None:
            continue

        uid_int = int(uid)

        # ── Core identity fields ────────────────────────────────────────
        name          = str(task.getName()) if task.getName() else ""
        outline_level = int(task.getOutlineLevel()) if task.getOutlineLevel() else 0
        outline_num   = str(task.getOutlineNumber()) if task.getOutlineNumber() else ""

        # Parent: MPXJ provides getParentTask()
        parent_task   = task.getParentTask()
        parent_uid    = int(parent_task.getUniqueID()) if parent_task and parent_task.getUniqueID() else None

        # Summary flag
        is_summary    = bool(task.getSummary()) if task.getSummary() is not None else False
        is_milestone  = bool(task.getMilestone()) if task.getMilestone() is not None else False

        # ── Baseline 0 fields ───────────────────────────────────────────
        baseline_start    = to_pydate(task.getBaselineStart())
        baseline_finish   = to_pydate(task.getBaselineFinish())
        baseline_duration = to_working_days(task.getBaselineDuration())

        # ── Current schedule fields ─────────────────────────────────────
        actual_start  = to_pydate(task.getActualStart())
        actual_finish = to_pydate(task.getActualFinish())
        duration      = to_working_days(task.getDuration())
        actual_dur    = to_working_days(task.getActualDuration())

        # ── Status / critical path ──────────────────────────────────────
        pct_complete  = float(task.getPercentageComplete()) if task.getPercentageComplete() is not None else 0.0
        is_critical   = bool(task.getCritical()) if task.getCritical() is not None else False
        total_slack   = to_working_days(task.getTotalSlack())

        # ── Scheduled (current) start/finish ───────────────────────────
        # For incomplete tasks these are the forecast dates
        sched_start   = to_pydate(task.getStart())
        sched_finish  = to_pydate(task.getFinish())

        # ── Resource assignments (names only) ───────────────────────────
        resource_names = []
        for asgn in task.getResourceAssignments():
            res = asgn.getResource()
            if res and res.getName():
                resource_names.append(str(res.getName()))
        resources_str = "; ".join(resource_names) if resource_names else ""

        tasks_rows.append({
            "snapshot":           snapshot_label,
            "uid":                uid_int,
            "name":               name,
            "outline_level":      outline_level,
            "outline_number":     outline_num,
            "parent_uid":         parent_uid,
            "is_summary":         is_summary,
            "is_milestone":       is_milestone,
            "baseline_start":     baseline_start,
            "baseline_finish":    baseline_finish,
            "baseline_duration":  baseline_duration,
            "actual_start":       actual_start,
            "actual_finish":      actual_finish,
            "duration":           duration,
            "actual_duration":    actual_dur,
            "pct_complete":       pct_complete,
            "is_critical":        is_critical,
            "total_slack":        total_slack,
            "sched_start":        sched_start,
            "sched_finish":       sched_finish,
            "resources":          resources_str,
        })

        # ── Predecessor relationships ────────────────────────────────────
        for pred in task.getPredecessors():
            pred_uid     = int(pred.getUniqueID()) if pred.getUniqueID() else None
            pred_task_id = int(pred.getTask().getUniqueID()) if pred.getTask() and pred.getTask().getUniqueID() else None
            rel_type     = str(pred.getType()) if pred.getType() else "FS"
            lag_wd       = to_working_days(pred.getLag())
            is_driving   = bool(pred.getDriving()) if hasattr(pred, "getDriving") and pred.getDriving() is not None else None

            preds_rows.append({
                "snapshot":       snapshot_label,
                "task_uid":       uid_int,
                "pred_task_uid":  pred_task_id,
                "rel_type":       rel_type,
                "lag_wd":         lag_wd if lag_wd else 0.0,
                "is_driving":     is_driving,
            })

    tasks_df = pd.DataFrame(tasks_rows)
    preds_df = pd.DataFrame(preds_rows)
    return tasks_df, preds_df


# ---------------------------------------------------------------------------
# Validation checks
# ---------------------------------------------------------------------------

def validate_snapshot(df: pd.DataFrame, snapshot_label: str) -> dict:
    """
    Run the playbook-mandated validation for one snapshot:
      1. Duration == ActualDuration for completed leaf tasks
      2. Count summary vs leaf tasks
    Returns a dict of validation results for the report.
    """
    leaves = df[~df["is_summary"]]
    complete_leaves = leaves[leaves["pct_complete"] >= 100.0].copy()

    if len(complete_leaves) == 0:
        dur_match_pct = None
        dur_mismatch_count = 0
    else:
        # Compare duration and actual_duration; allow 0.01 wd tolerance
        both_not_null = complete_leaves.dropna(subset=["duration", "actual_duration"])
        if len(both_not_null) == 0:
            dur_match_pct = None
            dur_mismatch_count = 0
        else:
            mismatches = both_not_null[
                (both_not_null["duration"] - both_not_null["actual_duration"]).abs() > 0.01
            ]
            dur_mismatch_count = len(mismatches)
            dur_match_pct = round(
                (1 - dur_mismatch_count / len(both_not_null)) * 100, 2
            )

    return {
        "snapshot":              snapshot_label,
        "total_tasks":           len(df),
        "summary_tasks":         int(df["is_summary"].sum()),
        "leaf_tasks":            int((~df["is_summary"]).sum()),
        "complete_leaves":       len(complete_leaves),
        "dur_match_pct":         dur_match_pct,
        "dur_mismatch_count":    dur_mismatch_count,
        "zero_baseline_tasks":   int(df["baseline_duration"].isna().sum()),
        "critical_task_count":   int(df["is_critical"].sum()),
    }


# ---------------------------------------------------------------------------
# UID persistence check (across all snapshots)
# ---------------------------------------------------------------------------

def check_uid_persistence(all_snapshots: list[str], snapshot_dir: Path) -> dict:
    """
    Load all snapshot parquet files and compute UID persistence:
    what % of UIDs from the first snapshot appear in every subsequent one.
    Also detects renamed tasks (same UID, different name across snapshots).
    """
    frames = []
    for snap in all_snapshots:
        p = snapshot_dir / f"{snap}.parquet"
        if p.exists():
            df = pd.read_parquet(p, columns=["snapshot", "uid", "name"])
            frames.append(df)

    if len(frames) < 2:
        return {"uid_persistence_pct": None, "rename_count": None}

    combined = pd.concat(frames, ignore_index=True)
    first_snap = all_snapshots[0]
    baseline_uids = set(combined.loc[combined["snapshot"] == first_snap, "uid"])

    # Persistence: UIDs present in ALL snapshots / UIDs in first snapshot
    uid_snapshot_counts = combined.groupby("uid")["snapshot"].nunique()
    n_snapshots = len(all_snapshots)
    persistent_uids = uid_snapshot_counts[uid_snapshot_counts == n_snapshots].index
    persistence_pct = round(len(persistent_uids) / len(baseline_uids) * 100, 2) if baseline_uids else None

    # Rename detection: same UID, multiple distinct names
    uid_name_counts = combined.groupby("uid")["name"].nunique()
    renamed_uids = uid_name_counts[uid_name_counts > 1]

    return {
        "baseline_uid_count":   len(baseline_uids),
        "persistent_uid_count": len(persistent_uids),
        "uid_persistence_pct":  persistence_pct,
        "rename_count":         int(len(renamed_uids)),
        "renamed_uid_examples": list(renamed_uids.head(10).index.astype(int)),
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Stage C — MPP/XML snapshot extraction")
    parser.add_argument(
        "--config", default="project_config.json",
        help="Path to project_config.json (default: ./project_config.json)"
    )
    parser.add_argument(
        "--mpxj-jar", default=None,
        help="Explicit path to mpxj.jar (optional; auto-detected via pip mpxj package)"
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Re-extract even if output parquet already exists (skip cached files by default)"
    )
    args = parser.parse_args()

    # ── Load config ─────────────────────────────────────────────────────
    cfg = load_config(args.config)
    xml_folder   = Path(cfg["paths"]["xml_snapshots_folder"])
    output_root  = Path(cfg["paths"]["output_root"])
    project_name = cfg["project"]["name"]

    stage_dir    = output_root / "stage_c"
    snap_dir     = stage_dir / "snapshots"
    pred_dir     = stage_dir / "predecessors"
    snap_dir.mkdir(parents=True, exist_ok=True)
    pred_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  Stage C — Data Extraction")
    print(f"  Project : {project_name}")
    print(f"  Source  : {xml_folder}")
    print(f"  Output  : {stage_dir}")
    print(f"{'='*60}\n")

    # ── Discover XML files ───────────────────────────────────────────────
    if not xml_folder.exists():
        print(f"ERROR: xml_snapshots_folder does not exist: {xml_folder}")
        sys.exit(1)

    # Collect all .xml files (may also be .mpp — MPXJ handles both)
    xml_files = sorted(
        [f for f in xml_folder.iterdir() if f.suffix.lower() in (".xml", ".mpp")],
        key=lambda f: f.stem
    )

    if not xml_files:
        print(f"ERROR: No .xml or .mpp files found in {xml_folder}")
        sys.exit(1)

    print(f"  Found {len(xml_files)} file(s) to process.\n")

    # ── Start JVM ────────────────────────────────────────────────────────
    print("  Starting JVM...")
    start_jvm(args.mpxj_jar)
    print()

    # ── Extract each snapshot ─────────────────────────────────────────────
    validation_results = []
    processed_snapshots = []
    errors = []

    for i, xml_path in enumerate(xml_files, start=1):
        # Use the filename stem as the snapshot label
        # MPP-Pipeline names files with dates; use stem directly
        snapshot_label = xml_path.stem
        snap_out  = snap_dir / f"{snapshot_label}.parquet"
        pred_out  = pred_dir / f"{snapshot_label}.parquet"

        print(f"  [{i:>3}/{len(xml_files)}] {xml_path.name}", end="  ")

        # Skip if already extracted and --force not set
        if snap_out.exists() and not args.force:
            print("(cached, skipping)")
            processed_snapshots.append(snapshot_label)
            # Load existing for validation tally
            existing = pd.read_parquet(snap_out)
            validation_results.append(validate_snapshot(existing, snapshot_label))
            continue

        try:
            tasks_df, preds_df = extract_file(xml_path, snapshot_label)

            # ── Unit validation: durations should be reasonable wd values
            # Flag any duration > 500 wd as a likely unit-conversion issue
            suspect = tasks_df[tasks_df["duration"].fillna(0) > 500]
            if len(suspect) > 0:
                print(f"\n    WARNING: {len(suspect)} tasks have duration > 500 wd — "
                      f"possible unit conversion issue. Check: {suspect['name'].head(3).tolist()}")

            # ── Persist ─────────────────────────────────────────────────
            tasks_df.to_parquet(snap_out, index=False)
            if len(preds_df) > 0:
                preds_df.to_parquet(pred_out, index=False)

            # ── Validate ─────────────────────────────────────────────────
            val = validate_snapshot(tasks_df, snapshot_label)
            validation_results.append(val)
            processed_snapshots.append(snapshot_label)

            print(
                f"{val['total_tasks']:>5} tasks  |  "
                f"leaves: {val['leaf_tasks']:>4}  |  "
                f"dur_match: {val['dur_match_pct']}%"
            )

        except Exception as e:
            msg = f"ERROR extracting {xml_path.name}: {e}"
            print(f"\n    {msg}")
            errors.append({"file": xml_path.name, "error": str(e), "trace": traceback.format_exc()})

    # ── UID persistence across all snapshots ─────────────────────────────
    print("\n  Running UID persistence check across all snapshots...")
    uid_stats = check_uid_persistence(processed_snapshots, snap_dir)
    print(
        f"  Baseline UIDs: {uid_stats.get('baseline_uid_count')}  |  "
        f"Persistent: {uid_stats.get('persistent_uid_count')}  |  "
        f"Persistence: {uid_stats.get('uid_persistence_pct')}%  |  "
        f"Renamed UIDs: {uid_stats.get('rename_count')}"
    )

    # ── Warn if persistence below threshold ──────────────────────────────
    warn_threshold = cfg.get("qc", {}).get("uid_persistence_warn_threshold_pct", 1.0)
    persistence = uid_stats.get("uid_persistence_pct")
    if persistence is not None and persistence < (100.0 - warn_threshold):
        print(f"\n  WARNING: UID persistence {persistence}% is below the "
              f"{100 - warn_threshold}% threshold defined in qc.uid_persistence_warn_threshold_pct")

    # ── Write extraction report ───────────────────────────────────────────
    report = {
        "generated":         datetime.now().isoformat(timespec="seconds"),
        "project":           project_name,
        "config_path":       str(args.config),
        "xml_folder":        str(xml_folder),
        "files_found":       len(xml_files),
        "files_processed":   len(processed_snapshots),
        "files_errored":     len(errors),
        "snapshots":         processed_snapshots,
        "uid_persistence":   uid_stats,
        "validation":        validation_results,
        "errors":            errors,
    }

    report_path = stage_dir / "extraction_report.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, default=str)

    # ── Summary ──────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"  Extraction complete.")
    print(f"  Snapshots processed : {len(processed_snapshots)}")
    print(f"  Errors              : {len(errors)}")
    print(f"  Parquet output      : {snap_dir}")
    print(f"  Report              : {report_path}")

    if errors:
        print(f"\n  ERRORS ({len(errors)}):")
        for e in errors:
            print(f"    {e['file']}: {e['error']}")

    # Playbook validation summary
    print(f"\n  Playbook validation (Duration == ActualDuration for completed leaves):")
    issues = [v for v in validation_results if v.get("dur_mismatch_count", 0) > 0]
    if not issues:
        print("  ✓ All snapshots: zero discrepancies.")
    else:
        print(f"  ✗ {len(issues)} snapshot(s) have Duration/ActualDuration mismatches:")
        for v in issues:
            print(f"    {v['snapshot']}: {v['dur_mismatch_count']} mismatches "
                  f"({v['dur_match_pct']}% match)")
        print("  → Investigate: schedulers may be using manual duration overrides.")

    print(f"{'='*60}\n")

    return 0 if not errors else 1


if __name__ == "__main__":
    sys.exit(main())
