"""Copy-only SEC migration and verified snapshots. Never delete a source database."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import uuid
from datetime import datetime
from pathlib import Path, PurePosixPath

from database import (DatabaseError, atomic_json, checked_path, connect_database,
                      connection_config, read_json, read_marker, validate_destination)
from database_io import database_lock


def _hash(path: Path) -> dict:
    checked_path(path)
    digest = hashlib.sha256()
    with path.open('rb') as handle:
        for chunk in iter(lambda: handle.read(1024*1024), b''):
            digest.update(chunk)
    return {'size': path.stat().st_size, 'sha256': digest.hexdigest()}


def inventory(root: Path) -> dict[str, dict]:
    root = checked_path(root)
    if not root.is_dir():
        raise DatabaseError('Source directory is unavailable')
    files = {}
    def scan(directory):
        for path in sorted(directory.iterdir()):
            checked_path(path)
            if path.is_dir():
                scan(path)
            elif path.is_file():
                files[path.relative_to(root).as_posix()] = _hash(path)
            else:
                raise DatabaseError('Unsupported source file')
    scan(root)
    return files


def _safe_relative(text: str) -> Path:
    path = PurePosixPath(text)
    if (not isinstance(text, str) or not text or path.is_absolute()
            or any(x in ('..', '.') for x in path.parts)
            or '\\' in text or ':' in text or path.as_posix() != text):
        raise DatabaseError('Invalid path in snapshot or migration manifest')
    return Path(*path.parts)


def _copy_verified(source: Path, target: Path, expected: dict) -> None:
    checked_path(source)
    checked_path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists():
        # Temporary partial copies are not published under the final file name.
        tmp = target.with_name(target.name + '.' + uuid.uuid4().hex + '.tmp')
        try:
            shutil.copy2(source, tmp)
            if _hash(tmp) != expected:
                raise DatabaseError('Copied file failed verification')
            tmp.replace(target)
        finally:
            tmp.unlink(missing_ok=True)
    if _hash(target) != expected:
        raise DatabaseError('Destination file differs; it was not overwritten')


def _dirs(root: Path):
    for name in ('filings','history','metadata','staging','snapshots'):
        checked_path(root/name).mkdir(parents=True, exist_ok=True)


def migrate_database(source: Path, destination: Path, tickers: list[str], *,
                     config_path: Path | None = None) -> dict:
    source, destination = checked_path(source), checked_path(destination)
    if source.is_relative_to(destination) or destination.is_relative_to(source):
        raise DatabaseError('Source and destination overlap')
    connection_config(config_path)  # Fail on corrupted settings before copying anything.
    manifest_path = destination/'staging/migration.json'
    if not manifest_path.exists():
        validate_destination(destination)
    with database_lock(source.parent, validate=False):
        before = inventory(source)
        if manifest_path.exists():
            migration = read_json(manifest_path)
            if (migration.get('source') != str(source) or migration.get('files') != before
                    or migration.get('destination') != str(destination)):
                raise DatabaseError('Migration resume does not match source and destination')
        else:
            migration = dict(format_version=1, database_id=str(uuid.uuid4()),
                             source=str(source), destination=str(destination), files=before,
                             tickers=list(dict.fromkeys(str(t).strip().upper() for t in tickers if t)),
                             created_at=datetime.now().astimezone().isoformat())
            _dirs(destination)
            atomic_json(manifest_path, migration)
        for relative, expected in before.items():
            name = _safe_relative(relative)
            _copy_verified(source/name, destination/'filings'/name, expected)
        if inventory(source) != before or inventory(destination/'filings') != before:
            raise DatabaseError('Source changed or copied database differs; connection was not switched')
        marker = dict(format_version=1, database_id=migration['database_id'],
                      created_at=migration['created_at'], test_mode=False)
        if (destination/'database.json').exists() and read_marker(destination) != marker:
            raise DatabaseError('Existing destination identity differs from migration')
        if (destination/'database.json').exists() and (destination/'metadata/migration.json').exists():
            completed = read_json(destination/'metadata/migration.json')
            if completed.get('verified') is True and completed.get('files') == before:
                connect_database(destination, config_path=config_path)
                return completed
        atomic_json(destination/'metadata/update_list.json',
                    dict(format_version=1, database_id=marker['database_id'], tickers=migration['tickers']))
        report = dict(migration, verified=True,
                      verified_at=datetime.now().astimezone().isoformat(),
                      file_count=len(before), filing_count=sum(bool(re.fullmatch(r'\d{10}-\d{2}-\d{6}\.json', Path(n).name)) for n in before),
                      total_bytes=sum(item['size'] for item in before.values()))
        atomic_json(destination/'metadata/migration.json', report)
        atomic_json(destination/'database.json', marker)
        connect_database(destination, config_path=config_path)
        note = source.parent/'DATABASE-MOVED.txt'
        note.write_text('SEC database migrated by verified copy. Original files retained.\n'
                        f'Use the updated application and database: {destination}\n'
                        'Do not run old application versions against this retained copy.\n', encoding='utf-8')
        return report


def _database_inventory(root: Path) -> dict:
    out = {'database.json': _hash(root/'database.json')}
    for name in ('filings','history','metadata'):
        for relative, value in inventory(root/name).items():
            out[name+'/'+relative] = value
    return out


def create_snapshot(root: Path, destination: Path | None = None) -> Path:
    root = checked_path(root)
    marker = read_marker(root)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:8]
    destination = checked_path(destination or root/'snapshots'/stamp)
    if (root.is_relative_to(destination) or
            (destination.is_relative_to(root) and not destination.is_relative_to(root/'snapshots'))):
        raise DatabaseError('Snapshot destination overlaps live database')
    if destination.exists() and any(destination.iterdir()):
        raise DatabaseError('Snapshot destination must be empty')
    with database_lock(root):
        files = _database_inventory(root)
        destination.mkdir(parents=True, exist_ok=True)
        for name in ('filings', 'history', 'metadata'):
            (destination/name).mkdir(exist_ok=True)
        for relative, expected in files.items():
            name = _safe_relative(relative)
            _copy_verified(root/name, destination/name, expected)
        if _database_inventory(root) != files:
            raise DatabaseError('Database changed during snapshot')
        atomic_json(destination/'snapshot.json', dict(format_version=1, verified=True,
                    database_id=marker['database_id'], created_at=datetime.now().astimezone().isoformat(),
                    files=files))
        verify_snapshot(destination)
    return destination


def verify_snapshot(snapshot: Path) -> dict:
    snapshot = checked_path(snapshot)
    data = read_json(snapshot/'snapshot.json')
    files = data.get('files')
    if data.get('format_version') != 1 or data.get('verified') is not True or not isinstance(files, dict):
        raise DatabaseError('Snapshot is incomplete')
    for name, expected in files.items():
        relative = _safe_relative(name)
        if relative.parts[0] not in ('database.json','filings','history','metadata'):
            raise DatabaseError('Unexpected snapshot file')
        if _hash(snapshot/relative) != expected:
            raise DatabaseError('Snapshot content failed verification')
    actual = inventory(snapshot)
    actual.pop('snapshot.json', None)
    if actual != files or read_marker(snapshot)['database_id'] != data.get('database_id'):
        raise DatabaseError('Snapshot file list or identity differs')
    # Ensure empty-but-required directories still exist in every snapshot.
    return data


def restore_snapshot(snapshot: Path, destination: Path) -> dict:
    snapshot, destination = checked_path(snapshot), checked_path(destination)
    data = verify_snapshot(snapshot)
    if destination.is_relative_to(snapshot) or snapshot.is_relative_to(destination):
        raise DatabaseError('Restore destination overlaps snapshot')
    validate_destination(destination)
    _dirs(destination)
    for name, expected in data['files'].items():
        if name == 'database.json':
            continue
        relative = _safe_relative(name)
        _copy_verified(snapshot/relative, destination/relative, expected)
    if verify_snapshot(snapshot) != data:
        raise DatabaseError('Snapshot changed during restore')
    _copy_verified(snapshot/'database.json', destination/'database.json', data['files']['database.json'])
    if _database_inventory(destination) != data['files']:
        raise DatabaseError('Restored database differs from snapshot')
    return dict(verified=True, root=str(destination), database_id=data['database_id'])
