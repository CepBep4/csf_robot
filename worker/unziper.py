"""Модуль для распаковки ZIP-архивов. Кроссплатформенно (Windows, macOS, Linux)."""

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


def unzip_to_dir(file_path: str | Path) -> Path:
    """
    Распаковывает ZIP-архив в директорию, имя которой совпадает с именем файла.

    Папка создаётся рядом с архивом (в том же каталоге). Например, для
    ``C:\\data\\55-004741`` содержимое окажется в ``C:\\data\\55-004741\\``.

    Args:
        file_path: Путь к ZIP-файлу (расширение может отсутствовать).

    Returns:
        Путь к директории с распакованным содержимым.

    Raises:
        zipfile.BadZipFile: Если файл не является корректным ZIP-архивом.
        FileNotFoundError: Если файл не найден.
    """
    path = Path(file_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    # Имя директории 1:1 соответствует имени файла (без расширения .zip)
    name = path.stem if path.suffix.lower() == ".zip" else path.name
    extract_dir = path.parent / name

    # Если архив без расширения (файл и папка — одно имя), сначала распаковываем во временную папку
    if extract_dir.resolve() == path.resolve():
        tmp_dir = Path(tempfile.mkdtemp(dir=path.parent))
        try:
            open_kw: dict = {}
            if sys.version_info >= (3, 11):
                open_kw["metadata_encoding"] = "utf-8"
            with zipfile.ZipFile(path, "r", **open_kw) as zf:
                zf.extractall(tmp_dir)
            path.unlink()
            tmp_dir.rename(extract_dir)
        except Exception:
            if tmp_dir.exists():
                shutil.rmtree(tmp_dir, ignore_errors=True)
            raise
        return extract_dir

    extract_dir.mkdir(parents=True, exist_ok=True)
    open_kw = {}
    if sys.version_info >= (3, 11):
        open_kw["metadata_encoding"] = "utf-8"
    with zipfile.ZipFile(path, "r", **open_kw) as zf:
        zf.extractall(extract_dir)
    return extract_dir


def unzip_all_in_dir(dir_path: str | Path) -> tuple[list[Path], list[tuple[Path, str]]]:
    """
    Находит все файлы в указанной папке (только верхний уровень) и каждый
    распаковывает через unzip_to_dir. Файлы, не являющиеся ZIP, пропускаются.

    Args:
        dir_path: Путь к папке с архивами.

    Returns:
        (processed, errors): список путей к распакованным директориям и список
        (файл, сообщение_об_ошибке) для файлов, которые не удалось обработать.
    """
    root = Path(dir_path).resolve()
    if not root.is_dir():
        raise NotADirectoryError(f"Не папка: {root}")

    processed: list[Path] = []
    errors: list[tuple[Path, str]] = []

    for item in root.iterdir():
        if not item.is_file():
            continue
        try:
            extract_dir = unzip_to_dir(item)
            processed.append(extract_dir)
        except zipfile.BadZipFile:
            errors.append((item, "Не ZIP-архив"))
        except FileNotFoundError as e:
            errors.append((item, str(e)))
        except Exception as e:
            errors.append((item, str(e)))

    return processed, errors
