#!/usr/bin/env python3
"""Install this skill for Codex and/or Claude Code using only Python's standard library."""
import argparse
import hashlib
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

NAME = "consulting-pptx-to-html"
EXCLUDED = {"__pycache__", ".DS_Store", ".git"}


def manifest(root):
    result = {}
    for path in sorted(root.rglob("*")):
        rel = path.relative_to(root)
        if any(part in EXCLUDED for part in rel.parts) or path.suffix == ".pyc":
            continue
        if path.is_file():
            result[rel.as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    return result


def ignore(directory, names):
    return [name for name in names if name in EXCLUDED or name.endswith(".pyc")]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--target", choices=("all", "codex", "claude"), default="all")
    parser.add_argument("--home", type=Path, default=Path.home(), help="Target user's home directory; defaults to the current user's home")
    parser.add_argument("--dry-run", action="store_true", help="Inspect and print the plan without writing anything")
    parser.add_argument("--replace", action="store_true", help="Replace a different installed copy after backing it up outside the discovery directories")
    args = parser.parse_args()
    source = Path(__file__).resolve().parents[1]
    if not (source / "SKILL.md").is_file() or source.name != NAME:
        raise ValueError("Keep the complete consulting-pptx-to-html directory intact before running this installer")
    home = args.home.expanduser().resolve()
    locations = {
        "codex": home / ".agents" / "skills" / NAME,
        "claude": home / ".claude" / "skills" / NAME,
    }
    targets = list(locations) if args.target == "all" else [args.target]
    expected = manifest(source)
    plan = []
    for target in targets:
        dest = locations[target]
        if dest.resolve() == source:
            action = "already-installed"
        elif dest.is_symlink():
            raise ValueError("Existing destination is a symlink; inspect it manually before replacing: " + str(dest))
        elif dest.exists():
            if not dest.is_dir():
                raise ValueError("Destination is not a directory: " + str(dest))
            action = "already-installed" if manifest(dest) == expected else "replace"
            if action == "replace" and not args.replace:
                raise ValueError("Different skill already exists; use --replace to back it up and update: " + str(dest))
        else:
            action = "install"
        plan.append({"target": target, "path": str(dest), "action": action})
    if args.dry_run:
        print(json.dumps({"dry_run": True, "plan": plan}, ensure_ascii=False, indent=2))
        return 0
    for item in plan:
        if item["action"] == "already-installed":
            continue
        dest = Path(item["path"])
        backup = None
        if item["action"] == "replace":
            stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
            backup = home / ".skill-backups" / NAME / item["target"] / stamp
            backup.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(dest), str(backup))
            item["backup"] = str(backup)
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copytree(source, dest, ignore=ignore)
            if manifest(dest) != expected:
                raise RuntimeError("Copied file hashes do not match the source")
        except Exception:
            if dest.exists() and dest.is_dir() and not dest.is_symlink():
                shutil.rmtree(dest)
            if backup is not None and backup.exists():
                shutil.move(str(backup), str(dest))
            raise
    print(json.dumps({"installed": True, "plan": plan, "note": "Open Codex or Claude Code and confirm that consulting-pptx-to-html appears in its skills list."}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        raise SystemExit(main())
    except (OSError, ValueError, RuntimeError) as exc:
        print("Install failed: " + str(exc), file=sys.stderr)
        raise SystemExit(1)
