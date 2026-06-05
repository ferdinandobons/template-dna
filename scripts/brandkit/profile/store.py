# SPDX-License-Identifier: MIT
"""Dual-store resolution, save, load, and listing of Brand Profiles.

A Brand Profile is a self-contained, copyable directory ``brand-kit/<name>/``
(see §3.1). It can live in two places:

- **project store**: ``./brand-kit/<name>`` (relative to the current working
  directory). **Project wins** — a project profile shadows a global one of the
  same name.
- **global store**: ``~/.claude/brand-kit/<name>``.

This module is the only code that knows those two paths. It also owns the
``provenance.sha256`` drift/tamper detection on the shell: the saved profile
records the shell's SHA-256, and :func:`load_profile` re-hashes the on-disk
shell and reports a mismatch (a hand-edited / corrupted shell) without refusing
to load — the caller decides what to do with the warning.

On-disk layout written by :func:`save_profile` (subset present depends on what
the extractor produced)::

    brand-kit/<name>/
    ├─ profile.json
    ├─ template/shell.<ext>
    ├─ assets/...            (optional)
    ├─ components/...        (optional)
    ├─ sections/...          (optional)
    └─ provenance.sha256     (the shell hash, also mirrored in profile.json)
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional, Union

from brandkit.profile import schema

PathLike = Union[str, os.PathLike]

# ---------------------------------------------------------------------------
# Store roots
# ---------------------------------------------------------------------------
PROJECT_STORE_DIRNAME = "brand-kit"
GLOBAL_STORE_SUBPATH = (".claude", "brand-kit")
PROFILE_JSON = "profile.json"
SHELL_HASH_FILE = "provenance.sha256"


def project_store_root(cwd: Optional[PathLike] = None) -> Path:
    """Return the project store root (``<cwd>/brand-kit``)."""
    base = Path(cwd) if cwd is not None else Path.cwd()
    return base / PROJECT_STORE_DIRNAME


def global_store_root() -> Path:
    """Return the global store root (``~/.claude/brand-kit``)."""
    return Path.home().joinpath(*GLOBAL_STORE_SUBPATH)


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------
@dataclass
class ProfileLocation:
    """Where a profile directory resolved to."""

    name: str
    scope: str          # "project" | "global"
    directory: Path     # the brand-kit/<name> dir
    profile_json: Path  # the profile.json inside it


@dataclass
class LoadedProfile:
    """A loaded profile plus its provenance check.

    Attributes:
        name: the profile name.
        scope: ``"project"`` or ``"global"``.
        directory: the profile root directory.
        profile: the parsed ``profile.json`` dict.
        shell_path: absolute path to the binary shell (may not exist if the
            profile is malformed; check ``shell_exists``).
        shell_drift: True if the on-disk shell's SHA-256 differs from the value
            recorded in the profile (tamper / corruption signal).
        recorded_sha256: the SHA-256 the profile claims for the shell.
        actual_sha256: the SHA-256 freshly computed from disk (None if missing).
        validation_problems: structural problems from :func:`schema.validate`.
    """

    name: str
    scope: str
    directory: Path
    profile: dict
    shell_path: Path
    shell_drift: bool = False
    recorded_sha256: Optional[str] = None
    actual_sha256: Optional[str] = None
    validation_problems: list[str] = field(default_factory=list)

    @property
    def shell_exists(self) -> bool:
        return self.shell_path.is_file()

    @property
    def kind(self) -> str:
        return self.profile.get("kind", "")


class ProfileNotFoundError(FileNotFoundError):
    """Raised when a profile name cannot be resolved in any store."""


class ProfileStoreError(RuntimeError):
    """Raised on a malformed or unreadable profile on disk."""


# ---------------------------------------------------------------------------
# Hashing
# ---------------------------------------------------------------------------
def sha256_file(path: PathLike, *, chunk: int = 1 << 20) -> str:
    """Return the hex SHA-256 of a file (streamed)."""
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def sha256_bytes(data: bytes) -> str:
    """Return the hex SHA-256 of a byte string."""
    return hashlib.sha256(data).hexdigest()


# ---------------------------------------------------------------------------
# Resolution
# ---------------------------------------------------------------------------
def resolve_profile_dir(
    name: str,
    scope: str = "auto",
    *,
    cwd: Optional[PathLike] = None,
) -> ProfileLocation:
    """Resolve a profile ``name`` to its on-disk directory.

    Args:
        name: the profile name (directory name under a store).
        scope: ``"auto"`` (project first, then global — **project wins**),
            ``"project"`` (only the project store), or ``"global"`` (only the
            global store).
        cwd: override the working directory used for the project store.

    Returns:
        A :class:`ProfileLocation` for the first store that contains the
        profile (a directory with a ``profile.json``).

    Raises:
        ValueError: if ``scope`` is not one of the three accepted values.
        ProfileNotFoundError: if no store in scope holds the profile.
    """
    if scope not in ("auto", "project", "global"):
        raise ValueError(f"scope must be auto|project|global, got {scope!r}")

    candidates: list[tuple[str, Path]] = []
    if scope in ("auto", "project"):
        candidates.append(("project", project_store_root(cwd) / name))
    if scope in ("auto", "global"):
        candidates.append(("global", global_store_root() / name))

    for store_scope, directory in candidates:
        pj = directory / PROFILE_JSON
        if pj.is_file():
            return ProfileLocation(
                name=name, scope=store_scope, directory=directory, profile_json=pj
            )

    searched = ", ".join(str(d) for _, d in candidates)
    raise ProfileNotFoundError(f"no profile {name!r} found (searched: {searched})")


def target_dir_for_save(
    name: str,
    scope: str = "project",
    *,
    cwd: Optional[PathLike] = None,
) -> Path:
    """Return the directory a save with ``scope`` would write to.

    ``scope`` of ``"auto"`` resolves to the **project** store for writes (the
    same "project wins" bias). ``"project"`` / ``"global"`` are explicit.
    """
    if scope in ("auto", "project"):
        return project_store_root(cwd) / name
    if scope == "global":
        return global_store_root() / name
    raise ValueError(f"scope must be auto|project|global, got {scope!r}")


# ---------------------------------------------------------------------------
# Save
# ---------------------------------------------------------------------------
def save_profile(
    directory: PathLike,
    profile: dict,
    shell_bytes: bytes,
    *,
    assets: Optional[dict] = None,
    extra_files: Optional[dict] = None,
    overwrite: bool = True,
) -> Path:
    """Write a complete profile directory to ``directory``.

    Re-extraction policy (M1) is **overwrite + provenance drift detection**:
    when ``overwrite`` is True the target ``profile.json`` and shell are
    replaced in place (other user files like ``PROFILE.md`` are left untouched);
    no history dir is kept (drift is detected later via the recorded SHA-256).

    Args:
        directory: the ``brand-kit/<name>`` dir to write (created if absent).
        profile: the ``profile.json`` dict. Its ``provenance.shell.sha256`` is
            (re)stamped here from ``shell_bytes`` so the recorded hash always
            matches what is written — callers need not pre-compute it.
        shell_bytes: the byte-for-byte template shell to write at
            ``template/shell.<ext>`` (ext from ``profile["kind"]``).
        assets: optional ``{relative_path: bytes}`` to write under the profile
            dir (e.g. ``{"assets/logo.primary.png": b"..."}``).
        extra_files: optional ``{relative_path: bytes|str}`` for any other files
            (e.g. ``components/kpi.card.xml``, ``PROFILE.md``). ``str`` values
            are UTF-8 encoded.
        overwrite: if False and ``directory`` already holds a ``profile.json``,
            raises :class:`ProfileStoreError`.

    Returns:
        The ``Path`` to the written ``profile.json``.

    Raises:
        ValueError: if ``profile["kind"]`` is missing/unknown.
        ProfileStoreError: if the target exists and ``overwrite`` is False.
    """
    kind = profile.get("kind")
    if kind not in schema.KIND_EXTENSION:
        raise ValueError(f"profile['kind'] missing or unknown: {kind!r}")
    ext = schema.KIND_EXTENSION[kind]

    root = Path(directory)
    pj = root / PROFILE_JSON
    if pj.exists() and not overwrite:
        raise ProfileStoreError(f"profile already exists at {root} (overwrite=False)")
    root.mkdir(parents=True, exist_ok=True)

    # 1) Write the shell verbatim and stamp its hash into provenance.
    shell_rel = f"template/shell.{ext}"
    shell_abs = root / shell_rel
    shell_abs.parent.mkdir(parents=True, exist_ok=True)
    shell_abs.write_bytes(shell_bytes)
    shell_hash = sha256_bytes(shell_bytes)

    prov = profile.setdefault("provenance", {})
    shell_block = prov.setdefault("shell", {})
    shell_block["path"] = shell_rel
    shell_block["sha256"] = shell_hash

    # 2) Assets and extra files.
    for rel, data in (assets or {}).items():
        _write_under(root, rel, data)
    for rel, data in (extra_files or {}).items():
        _write_under(root, rel, data)

    # 3) The profile.json itself (deterministic, sorted keys, trailing newline).
    pj.write_text(
        json.dumps(profile, indent=2, ensure_ascii=False, sort_keys=False) + "\n",
        encoding="utf-8",
    )

    # 4) The standalone shell-hash sidecar (mirrors provenance.shell.sha256).
    (root / SHELL_HASH_FILE).write_text(shell_hash + "\n", encoding="utf-8")

    return pj


def _write_under(root: Path, rel: str, data: Union[bytes, str]) -> None:
    """Write ``data`` to ``root/rel`` safely (no escaping the profile dir)."""
    root_resolved = root.resolve()
    target = (root / rel).resolve()
    try:
        target.relative_to(root_resolved)
    except ValueError:
        raise ProfileStoreError(f"refusing to write outside profile dir: {rel!r}")
    target.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(data, str):
        target.write_text(data, encoding="utf-8")
    else:
        target.write_bytes(data)


# ---------------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------------
def load_profile(
    name: str,
    scope: str = "auto",
    *,
    cwd: Optional[PathLike] = None,
) -> LoadedProfile:
    """Resolve, read, and provenance-check a profile by name.

    Reads ``profile.json``, re-hashes the on-disk shell, compares it to the
    recorded ``provenance.shell.sha256``, and runs :func:`schema.validate`.
    Never raises on drift or validation problems — those are returned on the
    :class:`LoadedProfile` so the caller (verify/generate) decides severity.

    Args:
        name: profile name.
        scope: ``"auto"`` (project wins) | ``"project"`` | ``"global"``.
        cwd: override the project-store working directory.

    Returns:
        A :class:`LoadedProfile`.

    Raises:
        ProfileNotFoundError: if no store in scope holds the profile.
        ProfileStoreError: if ``profile.json`` is unreadable / not valid JSON.
    """
    loc = resolve_profile_dir(name, scope, cwd=cwd)
    try:
        profile = json.loads(loc.profile_json.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProfileStoreError(f"cannot read {loc.profile_json}: {exc}") from exc

    kind = profile.get("kind")
    ext = schema.KIND_EXTENSION.get(kind, kind or "")
    shell_path = loc.directory / "template" / f"shell.{ext}"

    recorded = (
        (profile.get("provenance") or {}).get("shell") or {}
    ).get("sha256")
    actual = sha256_file(shell_path) if shell_path.is_file() else None
    drift = bool(recorded and actual and recorded != actual)

    return LoadedProfile(
        name=name,
        scope=loc.scope,
        directory=loc.directory,
        profile=profile,
        shell_path=shell_path,
        shell_drift=drift,
        recorded_sha256=recorded,
        actual_sha256=actual,
        validation_problems=schema.validate(profile),
    )


def profile_exists(name: str, scope: str = "auto", *, cwd: Optional[PathLike] = None) -> bool:
    """Return True if a profile ``name`` is resolvable in ``scope``."""
    try:
        resolve_profile_dir(name, scope, cwd=cwd)
        return True
    except ProfileNotFoundError:
        return False


# ---------------------------------------------------------------------------
# Listing
# ---------------------------------------------------------------------------
@dataclass
class ProfileSummary:
    """A lightweight listing entry (no shell hashing)."""

    name: str
    scope: str          # "project" | "global"
    directory: Path
    kind: Optional[str]
    display_name: Optional[str]
    verification_status: Optional[str]
    shadowed: bool = False  # True if a project profile hides a global of same name


def list_profiles(*, cwd: Optional[PathLike] = None) -> list[ProfileSummary]:
    """List every profile across both stores (project entries first).

    When the same name exists in both stores, the project entry is returned and
    the global one is included with ``shadowed=True`` so a UI can show the user
    that a global profile is being overridden.

    Returns:
        A list of :class:`ProfileSummary`, project store first then global,
        each store's entries sorted by name.
    """
    out: list[ProfileSummary] = []
    project_names: set[str] = set()

    for scope, root in (("project", project_store_root(cwd)), ("global", global_store_root())):
        if not root.is_dir():
            continue
        for child in sorted(root.iterdir(), key=lambda p: p.name):
            pj = child / PROFILE_JSON
            if not pj.is_file():
                continue
            name = child.name
            summary = _summarize(name, scope, child, pj)
            if scope == "project":
                project_names.add(name)
            elif name in project_names:
                summary.shadowed = True
            out.append(summary)
    return out


def _summarize(name: str, scope: str, directory: Path, pj: Path) -> ProfileSummary:
    """Build a :class:`ProfileSummary`, tolerating a partially malformed file."""
    kind = display = ver = None
    try:
        data = json.loads(pj.read_text(encoding="utf-8"))
        kind = data.get("kind")
        display = (data.get("identity") or {}).get("display_name")
        ver = (data.get("verification") or {}).get("status")
    except (OSError, json.JSONDecodeError):
        pass
    return ProfileSummary(
        name=name, scope=scope, directory=directory, kind=kind,
        display_name=display, verification_status=ver,
    )


def copy_profile(
    src_dir: PathLike,
    dest_dir: PathLike,
    *,
    overwrite: bool = False,
) -> Path:
    """Copy a whole profile directory (the self-contained-copy guarantee).

    Because every internal path in a profile is relative to its root, copying
    the directory tree is a complete move between stores. Useful for promoting a
    project profile to the global store (or vice versa).

    Raises:
        ProfileStoreError: if ``dest_dir`` exists and ``overwrite`` is False.
    """
    src = Path(src_dir)
    dest = Path(dest_dir)
    if dest.exists():
        if not overwrite:
            raise ProfileStoreError(f"destination exists: {dest} (overwrite=False)")
        shutil.rmtree(dest)
    shutil.copytree(src, dest)
    return dest
