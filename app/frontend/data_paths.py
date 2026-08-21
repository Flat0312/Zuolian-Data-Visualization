from __future__ import annotations

from pathlib import Path

STANDARD_DATA_FILES = (
    "persons.csv",
    "organizations.csv",
    "places.csv",
    "events.csv",
    "person_relations.csv",
    "org_memberships.csv",
    "org_membership_evidences.csv",
    "fact_evidences.csv",
    "event_participants.csv",
    "sources.csv",
)
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


def candidate_data_dirs(base_dir: Path | None = None, mode: str = "research") -> list[Path]:
    app_dir = (base_dir or Path(__file__).resolve().parent).resolve()
    project_root = app_dir.parent.parent
    folder = "publish" if mode == "public" else "processed"
    return _dedupe([project_root / "data" / folder])


def resolve_data_dir(
    base_dir: Path | None = None,
    required_files: tuple[str, ...] = CORE_DATA_FILES,
    mode: str = "research",
) -> Path:
    for candidate in candidate_data_dirs(base_dir, mode=mode):
        if all((candidate / filename).exists() for filename in required_files):
            return candidate
    return candidate_data_dirs(base_dir, mode=mode)[0]


def format_candidate_paths(paths: list[Path]) -> str:
    return "\n".join(f"- {path}" for path in paths)
