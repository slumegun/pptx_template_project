import hashlib
import os
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Protocol

from .config import get_settings


@dataclass(frozen=True)
class StoredFile:
    key: str
    sha256: str
    size_bytes: int


class StorageAdapter(Protocol):
    def put_file(self, source: Path, key: str) -> StoredFile: ...
    def path(self, key: str) -> Path: ...
    def exists(self, key: str) -> bool: ...
    def delete(self, key: str) -> None: ...


class LocalStorage:
    """Immutable files in a mounted volume. Storage keys never come from filenames."""

    def __init__(self, root: Path):
        self.root = root.resolve()
        self.root.mkdir(parents=True, exist_ok=True)

    def path(self, key: str) -> Path:
        candidate = (self.root / key).resolve()
        if not candidate.is_relative_to(self.root) or candidate == self.root:
            raise ValueError("Invalid storage key")
        return candidate

    def put_file(self, source: Path, key: str) -> StoredFile:
        destination = self.path(key)
        if destination.exists():
            raise FileExistsError(key)
        destination.parent.mkdir(parents=True, exist_ok=True)
        digest = hashlib.sha256()
        size = 0
        staged_path = None
        try:
            with NamedTemporaryFile(dir=destination.parent, delete=False) as staged:
                staged_path = Path(staged.name)
                with source.open("rb") as input_file:
                    for chunk in iter(lambda: input_file.read(1024 * 1024), b""):
                        staged.write(chunk)
                        digest.update(chunk)
                        size += len(chunk)
                staged.flush()
                os.fsync(staged.fileno())
            # A hard link publishes the complete file atomically and refuses to
            # overwrite an existing key, including another writer's result.
            os.link(staged_path, destination)
        finally:
            if staged_path is not None:
                staged_path.unlink(missing_ok=True)
        return StoredFile(key, digest.hexdigest(), size)

    def exists(self, key: str) -> bool:
        return self.path(key).is_file()

    def delete(self, key: str) -> None:
        self.path(key).unlink(missing_ok=True)


def get_storage() -> LocalStorage:
    return LocalStorage(get_settings().storage_root)

