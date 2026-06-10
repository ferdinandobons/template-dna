# SPDX-License-Identifier: MIT
"""Dual-store resolution, save, load, and listing of Brand Profiles.

A Brand Profile is a self-contained, copyable directory ``brand-kit/<name>/``
(see §3.1). It can live in two places:

- **project store**: ``./brand-kit/<name>`` (relative to the current working
  directory). **Project wins** - a project profile shadows a global one of the
  same name.
- **global store**: ``~/.claude/brand-kit/<name>``.

This module is the only code that knows those two paths. It also owns the
``provenance.sha256`` drift/tamper detection on the shell: the saved profile
records the shell's SHA-256, and :func:`load_profile` re-hashes the on-disk
shell and reports a mismatch (a hand-edited / corrupted shell) without refusing
to load - the caller decides what to do with the warning.

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
from typing import Optional, Union

from brandkit.profile import schema

PathLike = Union[str, os.PathLike]

# ---------------------------------------------------------------------------
# Store roots
# ---------------------------------------------------------------------------
PROJECT_STORE_DIRNAME = "brand-kit"
GLOBAL_STORE_SUBPATH = (".claude", "brand-kit")
PROFILE_JSON = "profile.json"
SHELL_HASH_FILE = "provenance.sha256"
# Content-addressed secondary-shell namespace (multi-template blending): every
# blended donor binary lives at ``template/blend-<sha256[:12]>.<ext>`` next to the
# primary ``template/shell.<ext>``. Only this prefix is ever pruned.
BLEND_SHELL_PREFIX = "blend-"


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
    scope: str  # "project" | "global"
    directory: Path  # the brand-kit/<name> dir
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

    @property
    def comprehension_present(self) -> bool:
        """True iff a current, sha-bound comprehension is available (see
        :func:`comprehension_is_present`)."""
        return comprehension_is_present(self.profile)

    @property
    def blended_shells(self) -> list[dict]:
        """A sorted copy of ``provenance.blended_shells`` (``[]`` when absent)."""
        prov = self.profile.get("provenance")
        entries = prov.get("blended_shells") if isinstance(prov, dict) else None
        if not isinstance(entries, list):
            return []
        return sorted(
            (dict(e) for e in entries if isinstance(e, dict)),
            key=lambda e: str(e.get("sha256")),
        )

    def blended_shell_drift(self) -> list[str]:
        """Re-hash every recorded secondary (blend) shell ON DEMAND.

        Deliberately NOT part of :func:`load_profile` (no per-load hashing of N
        secondaries on the generate hot path): the fail-closed QA check
        ``check_blend_shell_provenance`` is the enforcement; this is its
        programmatic mirror for tests/UI. Returns one problem string per
        missing / escaping / unreadable / hash-drifted entry (``[]`` == clean).
        """
        problems: list[str] = []
        root_resolved = Path(self.directory).resolve()
        for entry in self.blended_shells:
            rel = entry.get("path")
            recorded = entry.get("sha256")
            if not isinstance(rel, str) or not rel or not isinstance(recorded, str):
                problems.append(f"provenance.blended_shells: malformed entry {entry!r}")
                continue
            target = Path(self.directory) / rel
            try:
                target.resolve().relative_to(root_resolved)
            except ValueError:
                problems.append(f"blend shell path escapes the profile dir: {rel!r}")
                continue
            if not target.is_file():
                problems.append(f"recorded blend shell is missing: {rel}")
                continue
            try:
                actual = sha256_file(target)
            except OSError as exc:
                problems.append(f"could not hash blend shell {rel}: {exc}")
                continue
            if actual != recorded:
                problems.append(
                    f"blend shell hash drifted: recorded {recorded}, "
                    f"actual {actual} ({rel})"
                )
        return problems


class ProfileNotFoundError(FileNotFoundError):
    """Raised when a profile name cannot be resolved in any store."""


class ProfileStoreError(RuntimeError):
    """Raised on a malformed or unreadable profile on disk."""


# ---------------------------------------------------------------------------
# Name validation (path-traversal defense)
# ---------------------------------------------------------------------------
def _safe_name_segment(name: str) -> str:
    """Validate ``name`` is a single, safe path segment for a store subdir.

    A profile name maps directly to ``brand-kit/<name>``; it must therefore be
    exactly one path component with no way to climb out of the store. We reject
    (rather than silently rewrite) so a name always round-trips between
    :func:`target_dir_for_save` and :func:`resolve_profile_dir`.

    Refused: empty/whitespace, ``.`` / ``..``, any ``/`` or ``\\`` or
    :data:`os.sep` / :data:`os.altsep`, NUL bytes, and absolute or
    drive-qualified names. Everything else (e.g. ``acme``, ``deck-2024``) is
    returned unchanged.

    Raises:
        ProfileStoreError: if ``name`` is not a safe single segment.
    """
    if not isinstance(name, str) or not name.strip():
        raise ProfileStoreError(f"invalid profile name: {name!r}")
    bad = {"/", "\\", os.sep}
    if os.altsep:
        bad.add(os.altsep)
    if any(ch in name for ch in bad) or "\x00" in name:
        raise ProfileStoreError(f"profile name must be a single path segment: {name!r}")
    if name in (".", ".."):
        raise ProfileStoreError(f"profile name must be a single path segment: {name!r}")
    # An absolute or drive-qualified name (e.g. ``/etc`` or ``C:\\x``) collapses
    # the join and escapes the store; Path tells us if it has more than one part.
    p = Path(name)
    if p.is_absolute() or p.drive or len(p.parts) != 1 or p.parts[0] != name:
        raise ProfileStoreError(f"profile name must be a single path segment: {name!r}")
    return name


def _assert_within(target: Path, root: Path, name: str) -> Path:
    """Return ``target`` if it resolves inside ``root``; else raise (belt-and-braces)."""
    try:
        target.resolve().relative_to(root.resolve())
    except ValueError:
        raise ProfileStoreError(f"profile name escapes the store root: {name!r}")
    return target


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
# Comprehension cache binding (sha-bound; the cache key is the shell hash)
# ---------------------------------------------------------------------------
def comprehension_is_present(profile: dict) -> bool:
    """Return True iff the profile carries a *valid, current* comprehension.

    The comprehension cache is bound to the shell IN CODE (not prose): it counts
    as present only when its ``status`` is ``present`` AND its
    ``source_shell_sha256`` equals the live ``provenance.shell.sha256``. A drifted
    shell (re-extract) leaves a stale comprehension whose recorded sha no longer
    matches, so generation correctly falls back to the deterministic path until
    ``comprehend`` is re-run. This closes the stale-reuse-on-drift hole.
    """
    comp = profile.get("comprehension")
    if not isinstance(comp, dict):
        return False
    if comp.get("status") != "present":
        return False
    recorded = comp.get("source_shell_sha256")
    live = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    return bool(recorded and live and recorded == live)


def overrides_are_present(profile: dict) -> bool:
    """Return True iff the profile carries *valid, current, non-empty* overrides.

    The learned-overrides block (``rules.overrides``, Cluster B) is sha-bound to the
    shell IN CODE exactly like the comprehension cache: it counts as present (and the
    resolver is allowed to consume it) only when its ``status`` is ``present`` AND its
    ``source_shell_sha256`` equals the live ``provenance.shell.sha256`` AND at least
    one of its closed-kind containers (``reroute_roles`` / ``number_format_swaps`` /
    ``demo_clears``) is non-empty.

    A drifted shell (re-extract re-stamps ``provenance.shell.sha256`` at
    :func:`save_profile`) leaves a stale lesson whose recorded sha no longer matches,
    so the resolver correctly ignores it and reverts to today's deterministic,
    byte-identical path until ``learn`` is re-run. This mirrors
    :func:`comprehension_is_present` one-for-one.
    """
    rules = profile.get("rules")
    if not isinstance(rules, dict):
        return False
    overrides = rules.get("overrides")
    if not isinstance(overrides, dict):
        return False
    if overrides.get("status") != "present":
        return False
    recorded = overrides.get("source_shell_sha256")
    live = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    if not (recorded and live and recorded == live):
        return False
    return bool(
        overrides.get("reroute_roles")
        or overrides.get("number_format_swaps")
        or overrides.get("demo_clears")
    )


def blend_is_present(profile: dict) -> bool:
    """Return True iff the profile carries a *current, sha-bound* blend ledger.

    Mirrors :func:`comprehension_is_present` one-for-one: the blend ledger counts
    as present only when its ``status`` is ``present`` AND its
    ``source_shell_sha256`` equals the live ``provenance.shell.sha256``. A drifted
    primary shell (re-extract overwrites profile.json wholesale, so the ledger is
    gone anyway) can never present a stale ledger as current.
    """
    block = profile.get("blend")
    if not isinstance(block, dict):
        return False
    if block.get("status") != "present":
        return False
    recorded = block.get("source_shell_sha256")
    live = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    return bool(recorded and live and recorded == live)


# ---------------------------------------------------------------------------
# Secondary (blend) shell binaries - content-addressed under template/
# ---------------------------------------------------------------------------
def blend_shell_relpath(kind: str, sha256: str) -> str:
    """The content-addressed relative path for a blended secondary shell.

    ``template/blend-<sha256[:12]>.<ext>`` - minted from the sha (never from any
    input filename), so a crafted donor name can never steer the write path.

    Raises:
        ValueError: if ``kind`` is not a recognized :class:`schema.Kind`.
    """
    if kind not in schema.KIND_EXTENSION:
        raise ValueError(f"unknown kind {kind!r}")
    return f"template/{BLEND_SHELL_PREFIX}{sha256[:12]}.{schema.KIND_EXTENSION[kind]}"


def save_blend_shell(
    directory: PathLike, kind: str, shell_bytes: bytes
) -> tuple[str, str]:
    """Write a blended secondary shell binary under the profile dir.

    Content-addressed (same bytes -> same path, idempotent) and routed through
    :func:`_write_under` so the containment guard applies. Returns
    ``(relative_path, sha256)``.
    """
    sha = sha256_bytes(shell_bytes)
    rel = blend_shell_relpath(kind, sha)
    _write_under(Path(directory), rel, shell_bytes)
    return rel, sha


def prune_blend_shells(
    directory: PathLike, kind: str, keep_sha256s: set[str]
) -> list[str]:
    """Remove blend-shell binaries not referenced by ``keep_sha256s``.

    Only the ``template/blend-*`` namespace is ever touched (the primary shell and
    user files are structurally out of reach). Returns the removed relative paths,
    sorted. Unknown ``kind`` raises like :func:`blend_shell_relpath`.
    """
    if kind not in schema.KIND_EXTENSION:
        raise ValueError(f"unknown kind {kind!r}")
    ext = schema.KIND_EXTENSION[kind]
    root = Path(directory)
    keep_prefixes = {s[:12] for s in keep_sha256s}
    removed: list[str] = []
    for path in (root / "template").glob(f"{BLEND_SHELL_PREFIX}*.{ext}"):
        stem = path.name[len(BLEND_SHELL_PREFIX) : -(len(ext) + 1)]
        if stem in keep_prefixes:
            continue
        try:
            path.unlink()
        except OSError:
            continue
        removed.append(f"template/{path.name}")
    return sorted(removed)


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
        scope: ``"auto"`` (project first, then global - **project wins**),
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
    safe = _safe_name_segment(name)

    candidates: list[tuple[str, Path]] = []
    if scope in ("auto", "project"):
        root = project_store_root(cwd)
        candidates.append(("project", _assert_within(root / safe, root, name)))
    if scope in ("auto", "global"):
        root = global_store_root()
        candidates.append(("global", _assert_within(root / safe, root, name)))

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
    safe = _safe_name_segment(name)
    if scope in ("auto", "project"):
        root = project_store_root(cwd)
        return _assert_within(root / safe, root, name)
    if scope == "global":
        root = global_store_root()
        return _assert_within(root / safe, root, name)
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
            matches what is written - callers need not pre-compute it.
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
    if root.is_symlink():
        raise ProfileStoreError(f"refusing to write profile through symlink: {root}")
    pj = root / PROFILE_JSON
    if pj.exists() and not overwrite:
        raise ProfileStoreError(f"profile already exists at {root} (overwrite=False)")
    root.mkdir(parents=True, exist_ok=True)

    # 1) Write the shell verbatim and stamp its hash into provenance.
    shell_rel = f"template/shell.{ext}"
    shell_hash = sha256_bytes(shell_bytes)
    _write_under(root, shell_rel, shell_bytes)

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
    write_profile_json(root, profile)

    # 4) The standalone shell-hash sidecar (mirrors provenance.shell.sha256).
    write_shell_hash(root, shell_hash)

    return pj


def write_profile_json(directory: PathLike, profile: dict) -> Path:
    """Write ``profile.json`` under a profile dir using the safe writer.

    Keys are sorted so the on-disk form is deterministic regardless of dict
    insertion order (matching the "deterministic, sorted keys" guarantee its caller
    documents): re-running extract/comprehend yields a byte-identical profile.json.
    """
    root = Path(directory)
    data = json.dumps(profile, indent=2, ensure_ascii=False, sort_keys=True) + "\n"
    _write_under(root, PROFILE_JSON, data)
    return root / PROFILE_JSON


def write_shell_hash(directory: PathLike, shell_hash: str) -> Path:
    """Write the provenance sidecar under a profile dir using the safe writer."""
    root = Path(directory)
    _write_under(root, SHELL_HASH_FILE, shell_hash + "\n")
    return root / SHELL_HASH_FILE


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
    Never raises on drift or validation problems - those are returned on the
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

    # Forward-migration seam (additive-forever): identity today, so this is
    # behaviour-identical, but a future cross-major migration would normalise the
    # profile here BEFORE validate sees it.
    profile = schema.migrate(profile)

    kind = profile.get("kind")
    ext = schema.KIND_EXTENSION.get(kind, kind or "")
    shell_path = loc.directory / "template" / f"shell.{ext}"

    recorded = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
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


def profile_exists(
    name: str, scope: str = "auto", *, cwd: Optional[PathLike] = None
) -> bool:
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
    scope: str  # "project" | "global"
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

    for scope, root in (
        ("project", project_store_root(cwd)),
        ("global", global_store_root()),
    ):
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
        name=name,
        scope=scope,
        directory=directory,
        kind=kind,
        display_name=display,
        verification_status=ver,
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
