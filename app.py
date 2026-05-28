from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from collections import defaultdict
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Annotated, AsyncIterator

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from xlsx_tools import read_xlsx, write_xlsx


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
WORK_DIR = BASE_DIR / "work"
WORK_DIR.mkdir(exist_ok=True)
TASK_TTL_SECONDS = int(os.getenv("TASK_TTL_SECONDS", "7200"))
CLEANUP_INTERVAL_SECONDS = int(os.getenv("CLEANUP_INTERVAL_SECONDS", "600"))

PRODUCT_NAME_ALIASES = ["商品名称", "商品名", "产品名称", "产品名", "name", "product name"]
IMAGE_COLUMN_MAP = {
    "主图": "商品主图",
    "主图文件名": "商品主图",
    "商品主图": "商品主图",
    "详情图": "商品详情页图",
    "详情图文件名": "商品详情页图",
    "详情页图": "商品详情页图",
    "商品详情页图": "商品详情页图",
    "商品信息图": "商品信息",
    "信息图": "商品信息",
    "颜色图": "颜色图",
    "颜色图文件名": "颜色图",
}
IMAGE_FOLDER_NAMES = ["商品主图", "商品详情页图", "商品信息", "颜色图"]
DATA_CODE_ALIASES = ["资料编码", "商品编码", "编码", "款号", "货号", "code"]
HEADER_ROW_COUNT = 2


def _cleanup_expired_jobs() -> None:
    now = time.time()
    for path in WORK_DIR.iterdir():
        if not path.is_dir():
            continue
        try:
            if now - path.stat().st_mtime > TASK_TTL_SECONDS:
                shutil.rmtree(path)
        except FileNotFoundError:
            continue


def _cleanup_loop(stop_event: threading.Event) -> None:
    while not stop_event.wait(CLEANUP_INTERVAL_SECONDS):
        _cleanup_expired_jobs()


@asynccontextmanager
async def lifespan(_: FastAPI) -> AsyncIterator[None]:
    _cleanup_expired_jobs()
    stop_event = threading.Event()
    worker = threading.Thread(target=_cleanup_loop, args=(stop_event,), daemon=True)
    worker.start()
    try:
        yield
    finally:
        stop_event.set()
        worker.join(timeout=1)


app = FastAPI(title="Excel Splitter MVP", lifespan=lifespan)


@app.get("/health")
def health() -> dict[str, str | int]:
    return {
        "status": "ok",
        "task_ttl_seconds": TASK_TTL_SECONDS,
        "cleanup_interval_seconds": CLEANUP_INTERVAL_SECONDS,
    }


def _clean_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", value.strip())
    value = re.sub(r"\s+", " ", value).strip(". ")
    return value[:80] or "未命名商品"


def _parse_workbook(file_path: Path) -> tuple[list[list[str]], list[str], list[list[str]]]:
    try:
        rows = read_xlsx(file_path)
    except Exception as exc:  # pragma: no cover - surfaced as API error
        raise HTTPException(status_code=400, detail=f"Excel 解析失败：{exc}") from exc

    if len(rows) < HEADER_ROW_COUNT:
        raise HTTPException(status_code=400, detail="Excel 内容为空")
    header_rows = [[str(item).strip() for item in row] for row in rows[:HEADER_ROW_COUNT]]
    headers = header_rows[-1]
    body = rows[HEADER_ROW_COUNT:]
    if not all(any(row) for row in header_rows):
        raise HTTPException(status_code=400, detail="前两行需要是表头")
    return header_rows, headers, body


def _find_product_column(headers: list[str]) -> int:
    normalized = {header.strip().lower(): index for index, header in enumerate(headers)}
    for alias in PRODUCT_NAME_ALIASES:
        if alias.lower() in normalized:
            return normalized[alias.lower()]
    for index, header in enumerate(headers):
        if "商品" in header and "名称" in header:
            return index
    raise HTTPException(status_code=400, detail="找不到“商品名称”列，请在 Excel 第二行添加商品名称")


def _find_optional_column(headers: list[str], aliases: list[str]) -> int | None:
    normalized = {header.strip().lower(): index for index, header in enumerate(headers)}
    for alias in aliases:
        index = normalized.get(alias.lower())
        if index is not None:
            return index
    return None


def _cell(row: list[str], index: int) -> str:
    return row[index].strip() if index < len(row) else ""


def _group_rows(headers: list[str], rows: list[list[str]]) -> dict[str, list[list[str]]]:
    product_index = _find_product_column(headers)
    grouped: dict[str, list[list[str]]] = defaultdict(list)
    current_name = ""
    for row in rows:
        if not any(str(value).strip() for value in row):
            continue
        name = _cell(row, product_index)
        if name:
            current_name = name
        if current_name:
            grouped[current_name].append(row)
    return dict(grouped)


def _preview_payload(headers: list[str], rows: list[list[str]]) -> dict:
    grouped = _group_rows(headers, rows)
    total_rows = sum(len(items) for items in grouped.values())
    products = [
        {
            "index": index,
            "name": name,
            "rows": len(items),
            "preview": _row_to_dict(headers, items[0]) if items else {},
        }
        for index, (name, items) in enumerate(grouped.items(), start=1)
    ]
    return {
        "headers": headers,
        "summary": {
            "product_count": len(grouped),
            "sku_rows": total_rows,
            "error_rows": len(rows) - total_rows,
            "folder_count": len(grouped),
        },
        "products": products,
    }


def _row_to_dict(headers: list[str], row: list[str]) -> dict[str, str]:
    return {header: _cell(row, index) for index, header in enumerate(headers)}


def _extract_images(images_zip: UploadFile | None, target_dir: Path) -> dict[str, Path]:
    image_index: dict[str, Path] = {}
    if images_zip is None or not images_zip.filename:
        return image_index

    archive_path = target_dir / "images.zip"
    with archive_path.open("wb") as output:
        shutil.copyfileobj(images_zip.file, output)

    extract_dir = target_dir / "images"
    extract_dir.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(archive_path) as archive:
            for member in archive.infolist():
                destination = (extract_dir / member.filename).resolve()
                if not str(destination).startswith(str(extract_dir.resolve())):
                    continue
                archive.extract(member, extract_dir)
    except zipfile.BadZipFile as exc:
        raise HTTPException(status_code=400, detail="图片包必须是 ZIP 格式") from exc

    for path in extract_dir.rglob("*"):
        if path.is_file():
            image_index[path.name.lower()] = path
    return image_index


def _image_columns(headers: list[str]) -> list[tuple[int, str]]:
    result: list[tuple[int, str]] = []
    for index, header in enumerate(headers):
        folder = IMAGE_COLUMN_MAP.get(header.strip())
        if folder:
            result.append((index, folder))
    return result


def _copy_images(product_rows: list[list[str]], columns: list[tuple[int, str]], image_index: dict[str, Path], product_dir: Path) -> int:
    copied = 0
    for row in product_rows:
        for column_index, folder_name in columns:
            raw = _cell(row, column_index)
            if not raw:
                continue
            names = [part.strip() for part in re.split(r"[,，;；\n]", raw) if part.strip()]
            for name in names:
                source = image_index.get(Path(name).name.lower())
                if source is None:
                    continue
                destination_dir = product_dir / folder_name
                destination_dir.mkdir(exist_ok=True)
                destination = destination_dir / source.name
                if not destination.exists():
                    shutil.copy2(source, destination)
                    copied += 1
    return copied


def _data_code(headers: list[str], rows: list[list[str]], fallback: int) -> str:
    code_index = _find_optional_column(headers, DATA_CODE_ALIASES)
    if code_index is not None:
        for row in rows:
            code = _cell(row, code_index)
            if code:
                return _clean_filename(code)
    return str(fallback).zfill(4)


def _zip_dir(source_dir: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in source_dir.rglob("*"):
            relative_path = path.relative_to(source_dir)
            if path.is_dir():
                archive.writestr(f"{relative_path.as_posix()}/", b"")
            elif path.is_file():
                archive.write(path, relative_path)


@app.post("/api/preview")
async def preview(excel: Annotated[UploadFile, File()]) -> dict:
    if not excel.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传 .xlsx 文件")
    with tempfile.TemporaryDirectory(dir=WORK_DIR) as temp_name:
        temp_dir = Path(temp_name)
        excel_path = temp_dir / "upload.xlsx"
        with excel_path.open("wb") as output:
            shutil.copyfileobj(excel.file, output)
        _, headers, rows = _parse_workbook(excel_path)
        return _preview_payload(headers, rows)


@app.post("/api/generate")
async def generate(
    excel: Annotated[UploadFile, File()],
    options: Annotated[str, Form()] = "{}",
) -> FileResponse:
    if not excel.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传 .xlsx 文件")

    job_id = uuid.uuid4().hex
    job_dir = WORK_DIR / job_id
    source_dir = job_dir / "source"
    output_dir = job_dir / "output"
    source_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    excel_path = source_dir / "upload.xlsx"
    with excel_path.open("wb") as output:
        shutil.copyfileobj(excel.file, output)

    header_rows, headers, rows = _parse_workbook(excel_path)
    grouped = _group_rows(headers, rows)
    if not grouped:
        raise HTTPException(status_code=400, detail="没有可生成的商品数据")

    copied_images = 0
    for index, (product_name, product_rows) in enumerate(grouped.items(), start=1):
        product_dir = output_dir / _clean_filename(product_name)
        product_dir.mkdir(parents=True, exist_ok=True)
        for folder_name in IMAGE_FOLDER_NAMES:
            (product_dir / folder_name).mkdir(exist_ok=True)
        data_code = _data_code(headers, product_rows, index)
        write_xlsx(product_dir / f"商品数据-{data_code}.xlsx", header_rows, product_rows)

    manifest = {
        "source_excel": excel.filename,
        "options": json.loads(options or "{}"),
        "product_count": len(grouped),
        "sku_rows": sum(len(items) for items in grouped.values()),
        "copied_images": copied_images,
    }
    (output_dir / "生成说明.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    zip_path = job_dir / "商品数据包.zip"
    _zip_dir(output_dir, zip_path)
    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename="商品数据包.zip",
        background=None,
    )


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
