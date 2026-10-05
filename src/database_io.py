"""Serialized, crash-safe permanent SEC writes. History is never cleaned."""
from __future__ import annotations

import hashlib
import json
import os
import re
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from database import DatabaseError, checked_path, read_marker

_mutex = threading.RLock()
_local = threading.local()
_TICKER = re.compile(r'[A-Z0-9][A-Z0-9.-]{0,19}\Z')
_ACCESSION = re.compile(r'\d{10}-\d{2}-\d{6}\Z')


@contextmanager
def database_lock(root: Path, *, timeout: float = 30, validate: bool = True):
    root = checked_path(root)
    if validate:
        read_marker(root)
    with _mutex:
        held = getattr(_local, 'held', {})
        key = str(root)
        if key in held:
            yield
            return
        stage = checked_path(root / 'staging')
        stage.mkdir(parents=True, exist_ok=True)
        path = checked_path(stage / 'database.lock')
        with path.open('a+b') as handle:
            handle.seek(0, os.SEEK_END)
            if not handle.tell():
                handle.write(b'0')
                handle.flush()
            acquired = False
            deadline = time.monotonic() + timeout
            while not acquired:
                try:
                    handle.seek(0)
                    if os.name == 'nt':
                        import msvcrt
                        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                    else:
                        import fcntl
                        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    acquired = True
                except OSError as exc:
                    if time.monotonic() >= deadline:
                        raise DatabaseError('Database is busy; try again after the current operation') from exc
                    time.sleep(.05)
            held[key] = True
            _local.held = held
            try:
                yield
            finally:
                del held[key]
                handle.seek(0)
                if os.name == 'nt':
                    import msvcrt
                    msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _preserve(path: Path, content: bytes) -> None:
    path = checked_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        if path.read_bytes() != content:
            raise DatabaseError('History verification failed; current filing was not replaced')
        return
    # A partial history left by a crash is detected and blocks any later replacement.
    with path.open('xb') as handle:
        handle.write(content)
        handle.flush()
        os.fsync(handle.fileno())
    if path.read_bytes() != content:
        raise DatabaseError('Cannot verify preserved history')


def _meaningful(content: bytes):
    try:
        value = json.loads(content)
        if isinstance(value, dict):
            value.pop('cached_at', None)
        return value
    except (ValueError, UnicodeError):
        return content


def _replace(root: Path, target: Path, content: bytes) -> None:
    target = checked_path(target)
    if not target.is_relative_to(root):
        raise DatabaseError('Database write escapes root')
    tmp = checked_path(root / 'staging' / (uuid.uuid4().hex + '.tmp'))
    try:
        with tmp.open('xb') as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        target.parent.mkdir(parents=True, exist_ok=True)
        os.replace(tmp, target)
    finally:
        tmp.unlink(missing_ok=True)


def commit_filing(root: Path, ticker: str, accession: str, entry: dict) -> None:
    root = checked_path(root)
    ticker = str(ticker).strip().upper()
    if not _TICKER.fullmatch(ticker) or not _ACCESSION.fullmatch(str(accession)):
        raise DatabaseError('Invalid SEC filing path')
    if (not isinstance(entry, dict) or entry.get('accession_no') != accession
            or not isinstance(entry.get('cik'), int) or entry['cik'] <= 0
            or not isinstance(entry.get('schema_version'), int)
            or type(entry.get('has_financials')) is not bool
            or not entry.get('edgartools_version')):
        raise DatabaseError('Invalid SEC filing payload')
    try:
        content = json.dumps(entry, ensure_ascii=False).encode('utf-8')
        with database_lock(root):
            target = checked_path(root / 'filings' / ticker / (accession + '.json'))
            if target.exists():
                old = target.read_bytes()
                if _meaningful(old) == _meaningful(content):
                    return
                digest = hashlib.sha256(old).hexdigest()
                _preserve(root / 'history' / ticker / accession / (digest + '.json'), old)
            _replace(root, target, content)
    except (OSError, ValueError, TypeError) as exc:
        raise DatabaseError('Cannot persist SEC filing; previous data is retained') from exc


def write_metadata(root: Path, path: Path, value: dict) -> None:
    root = checked_path(root)
    target = checked_path(path)
    relative = target.relative_to(root) if target.is_relative_to(root) else None
    allowed = (relative is not None and (relative.parts[0] == 'metadata'
               or (len(relative.parts) == 3 and relative.parts[0] == 'filings'
                   and relative.name in ('_meta.json', '_sixk_probe.json'))))
    if not allowed:
        raise DatabaseError('Metadata writer cannot replace permanent filings or history')
    try:
        content = json.dumps(value, ensure_ascii=False).encode('utf-8')
        with database_lock(root):
            _replace(root, target, content)
    except (OSError, TypeError, ValueError) as exc:
        raise DatabaseError('Cannot save database management state') from exc


def clean_staging(root: Path, *, older_than_seconds: int = 3600) -> int:
    root = checked_path(root)
    count = 0
    with database_lock(root):
        for path in (root / 'staging').glob('*.tmp'):
            checked_path(path)
            if path.stat().st_mtime < time.time() - max(3600, older_than_seconds):
                path.unlink()
                count += 1
    return count
