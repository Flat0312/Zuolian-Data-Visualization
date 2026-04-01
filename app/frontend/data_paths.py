from __future__ import annotations

from pathlib import Path

STANDARD_DATA_FILES = (
    "persons.csv",
    "organizations.csv",
    "places.csv",
    "events.csv",
    "person_relations.csv",
    "org_memberships.csv",
    "event_participants.csv",
    "sources.csv",
)
LEGACY_DATA_FILES = ("nodes.csv", "edges.csv", "events.csv")
CORE_DATA_FILES = STANDARD_DATA_FILES


def _dedupe(paths: list[Path]) -> list[Path]:
    unique: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        key = str(path.resolve())
        if key in seen:
            continue
        seen.add(key)
        unique.append(path)
    return unique


def candidate_data_dirs(base_dir: Path | None = None) -> list[Path]:
    app_dir = (base_dir or Path(__file__).resolve().parent).resolve()
    project_root = app_dir.parent.parent
    return _dedupe([project_root / "data" / "processed"])


def resolve_data_dir(
    base_dir: Path | None = None,
    required_files: tuple[str, ...] = CORE_DATA_FILES,
) -> Path:
    for candidate in candidate_data_dirs(base_dir):
        if all((candidate / filename).exists() for filename in required_files):
            return candidate
        if all((candidate / filename).exists() for filename in LEGACY_DATA_FILES):
            return candidate
    return candidate_data_dirs(base_dir)[0]


def format_candidate_paths(paths: list[Path]) -> str:
    return "\n".join(f"- {path}" for path in paths)
