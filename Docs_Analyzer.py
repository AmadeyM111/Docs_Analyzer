from openpyxl import load_workbook
from pathlib import Path
import zipfile
import shutil
import mimetypes
import argparse
import logging
import json
import sys
import hashlib
import os

try:
    from PIL import Image, ExifTags  # type: ignore
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False


logger = logging.getLogger(__name__)


def _compute_sha256(file_path: Path) -> str:
    hasher = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def _extract_exif_dict(pil_image) -> dict:
    # Returns a flat EXIF dictionary with human-readable keys
    exif: dict = {}
    try:
        raw = pil_image.getexif()
        if raw:
            tag_map = {ExifTags.TAGS.get(k, k): v for k, v in raw.items()}
            # Small subset of the most useful fields
            exif = {
                "exif_make": tag_map.get("Make"),
                "exif_model": tag_map.get("Model"),
                "exif_date_time_original": tag_map.get("DateTimeOriginal"),
                "exif_orientation": tag_map.get("Orientation"),
            }
            # Full EXIF dump stored under a separate key
            exif["exif_json"] = tag_map
    except Exception as e:
        logger.debug("EXIF извлечь не удалось: %s", e)
    return exif


def inspect_image_file(file_path_str: str) -> dict:
    """
    Return a dictionary with image information.
    Safely handles non-image files (returns only basic fields).
    """
    file_path = Path(file_path_str)
    info: dict = {
        "extracted_path": str(file_path),
        "file_size_bytes": None,
        "sha256": None,
        "mime": None,
        "extension": file_path.suffix.lower(),
        "format": None,
        "color_mode": None,
        "width_px": None,
        "height_px": None,
        "dpi_x": None,
        "dpi_y": None,
    }

    try:
        info["file_size_bytes"] = os.path.getsize(file_path)
    except Exception:
        pass
    try:
        info["sha256"] = _compute_sha256(file_path)
    except Exception:
        pass
    try:
        mime_guess, _ = mimetypes.guess_type(str(file_path))
        info["mime"] = mime_guess
    except Exception:
        pass

    if not PIL_AVAILABLE:
        return info

    try:
        with Image.open(file_path) as im:
            info["format"] = im.format
            info["color_mode"] = im.mode
            width, height = im.size
            info["width_px"] = width
            info["height_px"] = height
            dpi = im.info.get("dpi")
            if isinstance(dpi, tuple) and len(dpi) >= 2:
                info["dpi_x"] = dpi[0]
                info["dpi_y"] = dpi[1]
            # EXIF
            info.update(_extract_exif_dict(im))
    except Exception as e:
        logger.debug("Файл не распознан как изображение или ошибка чтения: %s", e)

    return info


def extract_images_xlsx(xlsx_path, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    logger.info("Начало извлечения изображений из файла: %s", xlsx_path)
    logger.debug("Директория для сохранения: %s", out.resolve())

    # 1) Try via openpyxl API
    wb = load_workbook(xlsx_path, data_only=True)
    results = []

    for worksheet in wb.worksheets:
        logger.info("Обработка листа: %s", worksheet.title)
        images = getattr(worksheet, "_images", [])
        logger.debug("Найдено изображений на листе %s: %d", worksheet.title, len(images))

        for index_in_sheet, image in enumerate(images, start=1):
            # File name and extension
            extension = None
            if hasattr(image, "mime"):
                extension = mimetypes.guess_extension(image.mime) or ".bin"
            elif hasattr(image, "format"):
                extension = f".{image.format.lower()}"
            else:
                extension = ".bin"

            # Anchor address (top-left cell)
            cell_coordinate = None
            if hasattr(image, "anchor") and hasattr(image.anchor, "from"):
                from_anchor = image.anchor.from_
                cell_coordinate = f"{worksheet.cell(row=from_anchor.row+1, column=from_anchor.col+1).coordinate}"
            elif hasattr(image, "anchor") and isinstance(image.anchor, str):
                cell_coordinate = image.anchor
            else:
                cell_coordinate = "unknown"

            # Raw bytes retrieval
            raw_bytes = None
            if hasattr(image, "_data"):
                raw_bytes = image._data()  # bytes

            file_name = f"{worksheet.title}_{cell_coordinate}_{index_in_sheet}{extension}"
            out_file = out / file_name

            if raw_bytes:
                out_file.write_bytes(raw_bytes)
                logger.debug("Сохранено изображение: %s", out_file)
            else:
                logger.warning("Не удалось получить байты изображения для %s", file_name)

            results.append({
                "sheet": worksheet.title,
                "cell": cell_coordinate,
                "file": str(out_file),
                "ok": bool(raw_bytes),
            })

    # 2) Fallback: unpack ZIP and extract media/* (in case API missed something)
    with zipfile.ZipFile(xlsx_path) as archive:
        media_files = [name for name in archive.namelist() if name.startswith("xl/media/")]
        logger.info("Файлов в xl/media/: %d", len(media_files))
        for media_name in media_files:
            target = out / Path(media_name).name
            if not target.exists():
                with archive.open(media_name) as source, open(target, "wb") as destination:
                    shutil.copyfileobj(source, destination)
                logger.debug("Добавлен медиа-файл из архива: %s", target)
                results.append({"sheet": None, "cell": None, "file": str(target), "ok": True})
            else:
                logger.debug("Медиа-файл уже существует, пропуск: %s", target)

    logger.info("Завершено. Всего записей в результатах: %d", len(results))
    return results


def parse_args(argv):
    parser = argparse.ArgumentParser(description="Извлечение изображений из XLSX.")
    parser.add_argument("--xlsx", required=True, help="Путь к XLSX файлу")
    parser.add_argument("--out", required=True, help="Директория для сохранения изображений")
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["CRITICAL", "ERROR", "WARNING", "INFO", "DEBUG"],
        help="Уровень логирования (по умолчанию: INFO)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Печатать результаты в формате JSON",
    )
    parser.add_argument(
        "--out-json",
        help="Сохранить результаты в указанный JSON-файл",
    )
    return parser.parse_args(argv)


def configure_logging(level_name: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level_name),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    logger.debug("Логирование настроено. Уровень: %s", level_name)


if __name__ == "__main__":
    args = parse_args(sys.argv[1:])
    configure_logging(args.log_level)

    xlsx_path = args.xlsx
    output_dir = args.out

    results = extract_images_xlsx(xlsx_path, output_dir)

    # Дополнительная инспекция изображений
    enriched_results = []
    for item in results:
        file_path = item.get("file")
        if file_path and Path(file_path).exists():
            meta = inspect_image_file(file_path)
            merged = {**item, **meta}
        else:
            merged = dict(item)
        enriched_results.append(merged)

    total = len(results)
    total_ok = sum(1 for item in results if item.get("ok"))
    total_fail = total - total_ok
    print(f"Итог: найдено {total} элементов, сохранено успешно: {total_ok}, ошибок: {total_fail}")

    # Запись в JSON-файл, если указан --out-json
    if getattr(args, "out_json", None):
        out_json_path = Path(args.out_json)
        out_json_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_json_path, "w", encoding="utf-8") as f:
            json.dump(enriched_results, f, ensure_ascii=False, indent=2)
        print(f"JSON сохранён: {out_json_path}")

    if args.json:
        print(json.dumps(enriched_results, ensure_ascii=False, indent=2))
    else:
        # Краткий табличный вывод
        print("sheet\tcell\tfile\tok")
        for item in enriched_results:
            print(f"{item.get('sheet')}\t{item.get('cell')}\t{item.get('file')}\t{item.get('ok')}")