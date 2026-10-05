"""Permanent SEC database identity and connection. Resolving never creates a database."""
from __future__ import annotations

import ctypes
import hashlib
import json
import os
import stat
import tempfile
import uuid
from datetime import datetime
from pathlib import Path


class DatabaseError(RuntimeError):
    """A database could not be connected or safely persisted."""


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def default_database_path() -> Path:
    documents = Path.home() / 'Documents'
    if os.name == 'nt':
        class GUID(ctypes.Structure):
            _fields_ = [('a', ctypes.c_ulong), ('b', ctypes.c_ushort),
                        ('c', ctypes.c_ushort), ('d', ctypes.c_ubyte * 8)]
        guid = GUID.from_buffer_copy(uuid.UUID('fdd39ad0-238f-46af-adb4-6c85480369c7').bytes_le)
        out = ctypes.c_wchar_p()
        result = ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None,
                                                          ctypes.byref(out))
        if result != 0:
            raise DatabaseError('Cannot locate Windows Documents folder')
        try:
            documents = Path(out.value)
        finally:
            ctypes.windll.ole32.CoTaskMemFree(ctypes.cast(out, ctypes.c_void_p))
    return documents / 'SEC財報資料庫'


def checked_path(path: Path) -> Path:
    """Reject links/junctions before resolving, including ancestors."""
    path = Path(os.path.abspath(path))
    for item in (path, *path.parents):
        try:
            info = item.lstat()
        except FileNotFoundError:
            continue
        except OSError as exc:
            raise DatabaseError('Cannot inspect database path') from exc
        if stat.S_ISLNK(info.st_mode) or getattr(info, 'st_file_attributes', 0) & 0x400:
            raise DatabaseError('Database paths cannot contain links or junctions')
    return path.resolve()


def read_json(path: Path) -> dict:
    try:
        data = json.loads(checked_path(path).read_text(encoding='utf-8-sig'))
    except (OSError, ValueError) as exc:
        raise DatabaseError('Cannot read database connection or metadata') from exc
    if not isinstance(data, dict):
        raise DatabaseError('Invalid database metadata')
    return data


def atomic_json(path: Path, value: dict) -> None:
    """For marker/config only; database contents use database_io locking."""
    path = checked_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        with tmp.open('x', encoding='utf-8') as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, path)
    except (OSError, TypeError, ValueError) as exc:
        raise DatabaseError('Cannot save database connection or metadata') from exc
    finally:
        tmp.unlink(missing_ok=True)


def read_marker(root: Path) -> dict:
    root = checked_path(root)
    data = read_json(root / 'database.json')
    try:
        uuid.UUID(data['database_id'])
    except (ValueError, TypeError, KeyError, AttributeError) as exc:
        raise DatabaseError('Invalid database identity') from exc
    if data.get('format_version') != 1 or type(data.get('test_mode')) is not bool:
        raise DatabaseError('Unsupported database identity format')
    if not (root / 'filings').is_dir():
        raise DatabaseError('Database filings directory is missing')
    checked_path(root / 'filings')
    return data


def validate_destination(root: Path, *, test_mode: bool = False) -> Path:
    root = checked_path(root)
    if not test_mode and root.is_relative_to(project_root()):
        raise DatabaseError('Permanent database must be outside the program directory')
    if test_mode and not root.is_relative_to(checked_path(Path(tempfile.gettempdir()))):
        raise DatabaseError('Test database must be inside the OS temporary directory')
    if root.exists() and (not root.is_dir() or any(root.iterdir())):
        raise DatabaseError('Destination must be new or empty')
    return root


def create_database(root: Path, *, test_mode: bool = False) -> dict:
    root = validate_destination(root, test_mode=test_mode)
    for name in ('filings', 'history', 'metadata', 'staging', 'snapshots'):
        (root / name).mkdir(parents=True, exist_ok=True)
    marker = dict(format_version=1, database_id=str(uuid.uuid4()),
                  created_at=datetime.now().astimezone().isoformat(), test_mode=test_mode)
    atomic_json(root / 'metadata/update_list.json',
                dict(format_version=1, database_id=marker['database_id'], tickers=[]))
    atomic_json(root / 'database.json', marker)
    return marker


def _config_path(path: Path | None) -> Path:
    if path is not None:
        return Path(path)
    from config import CONFIG_PATH
    return Path(os.environ.get('SEC_CONFIG_PATH') or CONFIG_PATH)


def connection_config(config_path: Path | None = None) -> dict:
    path = _config_path(config_path)
    return read_json(path) if path.exists() else {}


def connect_database(root: Path, *, config_path: Path | None = None) -> dict:
    root = checked_path(root)
    marker = read_marker(root)
    path = checked_path(_config_path(config_path))
    if path.exists():
        raw = path.read_bytes()
        try:
            cfg = json.loads(raw.decode('utf-8-sig'))
            if not isinstance(cfg, dict):
                raise ValueError('Invalid settings object')
        except (ValueError, UnicodeError):
            backup = path.with_name(path.name + '.corrupt-' + hashlib.sha256(raw).hexdigest() + '.bak')
            if backup.exists():
                if backup.read_bytes() != raw:
                    raise DatabaseError('Cannot verify preserved corrupt configuration')
            else:
                with backup.open('xb') as handle:
                    handle.write(raw)
                    handle.flush()
                    os.fsync(handle.fileno())
            cfg = {}
    else:
        cfg = {}
    cfg.update(database_path=str(root), database_id=marker['database_id'])
    atomic_json(_config_path(config_path), cfg)
    return marker


def database_root(*, config_path: Path | None = None) -> Path:
    override = os.environ.get('SEC_LOCAL_DB_ROOT') if config_path is None else None
    if override:
        root = checked_path(Path(override))
        read_marker(root)
        return root
    cfg = connection_config(config_path)
    location, expected_id = cfg.get('database_path'), cfg.get('database_id')
    if location or expected_id:
        if not location or not expected_id:
            raise DatabaseError('Database connection is incomplete; reconnect the database')
        root = checked_path(Path(location))
        if read_marker(root)['database_id'] != expected_id:
            raise DatabaseError('Database identity differs from registered connection')
        return root
    root = checked_path(default_database_path())
    read_marker(root)
    # Discovery is read-only. GUI/CLI can explicitly register this identity.
    return root


def is_test_database(root: Path) -> bool:
    root = checked_path(root)
    return (read_marker(root)['test_mode'] is True
            and root.is_relative_to(checked_path(Path(tempfile.gettempdir()))))
