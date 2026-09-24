#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量把 Word/PDF 等文档转换为 Markdown，尽量保证内容准确。

核心策略：
1. 数字原生 PDF：直接用 PyMuPDF/pymupdf4llm 抽取文本，不做 OCR，
   因此不会产生 OCR 识别错字；程序还会用原始文本层回比 Markdown 文本。
2. 扫描型 PDF：优先调用 MinerU 做版面/表格/公式/OCR 解析；MinerU 不可用时，
   默认拒绝输出，避免把低质量 OCR 结果当成准确结果。
   只有显式加 --allow-basic-ocr 时才用 RapidOCR 生成“需人工复核”的 md。
3. DOCX：优先 Pandoc 转 GFM；Pandoc 不可用时用 Mammoth/python-docx 回退。
4. 旧版 DOC：先用 LibreOffice(soffice) 或 Word COM 转成 DOCX，再走 DOCX 流程。
5. 所有输出：UTF-8、目录结构镜像、逐文件 QA 报告、待复核清单。

注意：任何 OCR 都不可能 100% 保证无误。程序会标记所有 OCR/扫描件输出，
并把低置信页面、可疑字符、原文覆盖度不足的文件放进 review_queue.csv。
请把这些文件作为人工复核对象。
"""

from __future__ import annotations

import argparse
import csv
import html
import importlib.util
import itertools
import logging
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import urllib.parse
from collections import Counter
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

VERSION = "1.0.0"
DEFAULT_SOURCE = r"D:\chemical_design\文档\标准2(1)\标准2"
SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".doc", ".txt", ".md", ".html", ".htm"}

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
KEEP_FOR_COMPARE_RE = re.compile(
    r"[0-9A-Za-z\u0370-\u03ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]"
)
SUSPICIOUS_RE = re.compile(r"[\ufffd\u25a1\u25a0\ue000-\uf8ff]")
CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
IMAGE_MD_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")

LOG = logging.getLogger("convert_to_md")


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def setup_logging(log_file: Path) -> None:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    LOG.setLevel(logging.INFO)
    LOG.handlers.clear()
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    LOG.addHandler(sh)
    fh = logging.FileHandler(str(log_file), encoding="utf-8")
    fh.setFormatter(fmt)
    LOG.addHandler(fh)


def has_module(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False


def which(name: str) -> Optional[Path]:
    if not name:
        return None
    p = shutil.which(name)
    if p:
        return Path(p)
    pp = Path(name)
    if pp.exists() and pp.is_file():
        return pp
    if name == "soffice":
        for c in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(c).exists():
                return Path(c)
    if name == "pandoc":
        for c in (
            r"C:\Program Files\Pandoc\pandoc.exe",
            r"C:\Program Files (x86)\Pandoc\pandoc.exe",
            str(Path.home() / "AppData" / "Local" / "Pandoc" / "pandoc.exe"),
        ):
            if Path(c).exists():
                return Path(c)
    return None


def find_mineru_module() -> Optional[List[str]]:
    """优先用 python -m mineru.cli.main，避免 Windows 上 .exe 启动器路径问题。"""
    try:
        if importlib.util.find_spec("mineru.cli.main") is not None:
            return [sys.executable, "-m", "mineru.cli.main"]
    except Exception:
        pass
    return None


def find_mineru() -> Optional[List[str]]:
    module_cmd = find_mineru_module()
    if module_cmd:
        return module_cmd
    exe = which("mineru")
    if exe:
        return [str(exe)]
    return None


def import_fitz():
    """兼容老版 PyMuPDF 的 fitz 导入方式。"""
    try:
        import pymupdf as fitz  # type: ignore
    except Exception:
        import fitz  # type: ignore
    return fitz


def run_command(cmd: Sequence[Any], timeout: int = 3600) -> subprocess.CompletedProcess:
    shown = " ".join(str(x) for x in cmd)
    LOG.info("运行命令: %s", shown)
    p = subprocess.run(
        [str(x) for x in cmd],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )
    if p.returncode != 0:
        raise RuntimeError(
            "命令失败 (%s): %s\nSTDOUT:\n%s\nSTDERR:\n%s"
            % (p.returncode, shown, (p.stdout or "")[-3000:], (p.stderr or "")[-3000:])
        )
    if p.stdout:
        LOG.debug("STDOUT: %s", p.stdout[-2000:])
    if p.stderr:
        LOG.debug("STDERR: %s", p.stderr[-2000:])
    return p


def read_text_auto(path: Path) -> str:
    for enc in ("utf-8-sig", "utf-8", "gb18030", "big5"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def normalize_for_compare(text: str) -> str:
    """只保留中英数字等字符，用于原文/输出文本一致性回比。"""
    text = unicodedata.normalize("NFKC", text or "")
    text = CONTROL_RE.sub("", text)
    return "".join(KEEP_FOR_COMPARE_RE.findall(text))


def bigram_coverage(src: str, out: str) -> float:
    if not src:
        return 1.0
    if len(src) == 1:
        return 1.0 if src in out else 0.0
    sc = Counter(src[i : i + 2] for i in range(len(src) - 1))
    oc = Counter(out[i : i + 2] for i in range(len(out) - 1))
    matched = sum(min(n, oc.get(g, 0)) for g, n in sc.items())
    total = sum(sc.values())
    return matched / float(max(1, total))


def clean_cell(value: Any) -> str:
    s = "" if value is None else str(value)
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    s = s.replace("|", "\\|").strip()
    return s


def rows_to_markdown(rows: Iterable[Iterable[Any]]) -> str:
    cleaned: List[List[str]] = []
    for row in rows:
        r = [clean_cell(c) for c in row]
        if any(x for x in r):
            cleaned.append(r)
    if not cleaned:
        return ""
    max_cols = max(len(r) for r in cleaned)
    cleaned = [r + [""] * (max_cols - len(r)) for r in cleaned]
    header = cleaned[0]
    body = cleaned[1:]
    lines = [
        "| " + " | ".join(header) + " |",
        "|" + "|".join([" --- "] * max_cols) + "|",
    ]
    for row in body:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def fix_image_paths(md_text: str, base_dir: Path, assets_dir: Path) -> str:
    """把 Markdown 图片链接统一成相对 base_dir 的路径，并处理绝对路径。"""
    if not md_text:
        return md_text

    text = md_text.replace(str(assets_dir), assets_dir.name)
    text = text.replace(str(assets_dir).replace("\\", "/"), assets_dir.name)

    def quote_local(path_text: str) -> str:
        cleaned = urllib.parse.unquote(path_text.replace(os.sep, "/"))
        return urllib.parse.quote(cleaned, safe="/:._-~")

    def repl(m: re.Match) -> str:
        alt, raw = m.group(1), m.group(2).strip()
        if raw.startswith(("http://", "https://", "data:")):
            return m.group(0)
        raw = raw.strip("<>").strip()
        if raw.lower().startswith("file:///"):
            raw = raw[8:]
        try:
            p = Path(urllib.parse.unquote(raw).replace("/", os.sep))
            if p.is_absolute():
                rel = os.path.relpath(str(p), str(base_dir))
                return "![%s](%s)" % (alt, quote_local(rel))
        except Exception:
            pass
        return "![%s](%s)" % (alt, quote_local(raw))

    return IMAGE_MD_RE.sub(repl, text)


# ---------------------------------------------------------------------------
# 结果记录
# ---------------------------------------------------------------------------

@dataclass
class ConversionRecord:
    source: str
    output: str = ""
    ext: str = ""
    method: str = ""
    kind: str = ""
    status: str = "OK"
    source_chars: int = 0
    output_chars: int = 0
    coverage: Optional[float] = None
    suspicious: int = 0
    min_conf: Optional[float] = None
    elapsed: float = 0.0
    note: str = ""


# ---------------------------------------------------------------------------
# 转换器
# ---------------------------------------------------------------------------

class Converter:
    def __init__(self, source: Path, output: Path, qa_dir: Path, args: argparse.Namespace):
        self.source = source
        self.output = output
        self.qa_dir = qa_dir
        self.args = args
        self.pandoc = which("pandoc")
        self.soffice = which("soffice")
        self.records: List[ConversionRecord] = []

    # ---- 主流程 -----------------------------------------------------------

    def collect_files(self) -> List[Path]:
        exts = set(SUPPORTED_EXTENSIONS)
        if self.args.extensions:
            exts = {
                (e.strip() if e.strip().startswith(".") else "." + e.strip())
                for e in self.args.extensions.split(",")
                if e.strip()
            }
        files: List[Path] = []
        for p in sorted(self.source.rglob("*")):
            if not p.is_file():
                continue
            if p.suffix.lower() not in exts:
                continue
            if self.output == p or self.output in p.parents:
                continue
            if self.qa_dir == p or self.qa_dir in p.parents:
                continue
            files.append(p)
        return files

    def convert_all(self) -> List[ConversionRecord]:
        files = self.collect_files()
        total = len(files)
        LOG.info("源目录: %s", self.source)
        LOG.info("输出目录: %s", self.output)
        LOG.info("QA 目录: %s", self.qa_dir)
        LOG.info("待转换文件: %d", total)
        for idx, src in enumerate(files, 1):
            if self.args.limit and idx > self.args.limit:
                break
            rel = src.relative_to(self.source)
            LOG.info("[%d/%d] %s", idx, total, rel)
            rec = self.convert_one(src)
            self.records.append(rec)
            LOG.info("  -> %s | %s | %s", rec.status, rec.method, rec.note)
        return self.records

    def convert_one(self, src: Path) -> ConversionRecord:
        rel = src.relative_to(self.source)
        out_dir = self.output / rel.parent
        out_md = out_dir / (src.stem + ".md")
        assets_dir = out_dir / (src.stem + "_assets")
        try:
            out_rel = str(out_md.relative_to(self.output)).replace("\\", "/")
        except ValueError:
            out_rel = str(out_md)
        rec = ConversionRecord(
            source=str(rel).replace("\\", "/"),
            output=out_rel,
            ext=src.suffix.lower(),
        )
        t0 = time.time()
        try:
            if (
                self.args.resume
                and not self.args.overwrite
                and out_md.exists()
                and out_md.stat().st_size > 0
            ):
                rec.status = "SKIP_EXISTING"
                rec.method = "skip"
                rec.note = "输出已存在；加 --overwrite 可重新转换"
                rec.elapsed = time.time() - t0
                return rec

            out_dir.mkdir(parents=True, exist_ok=True)
            suffix = src.suffix.lower()
            if suffix == ".pdf":
                method, kind, extra = self.convert_pdf(src, out_md, assets_dir)
            elif suffix == ".docx":
                method, kind, extra = self.convert_docx(src, out_md, assets_dir)
            elif suffix == ".doc":
                method, kind, extra = self.convert_doc(src, out_md, assets_dir)
            elif suffix in (".txt", ".md"):
                method, kind, extra = self.convert_text(src, out_md, assets_dir)
            elif suffix in (".html", ".htm"):
                method, kind, extra = self.convert_html(src, out_md, assets_dir)
            else:
                raise RuntimeError("不支持的文件类型: %s" % suffix)

            rec.method = method
            rec.kind = kind
            rec.note = str(extra.get("note", "")).strip()
            self.verify(src, out_md, method, kind, extra, rec)
        except Exception as e:
            LOG.exception("转换失败: %s", src)
            rec.status = "FAILED"
            rec.note = "转换失败: %s" % e
            try:
                if out_md.exists():
                    out_md.unlink()
            except OSError:
                pass
        rec.elapsed = time.time() - t0
        return rec

    # ---- PDF --------------------------------------------------------------

    def analyze_pdf(self, path: Path) -> Tuple[bool, float, int, int, int]:
        """判断 PDF 是否具有可用的文本层，返回 native, avg_chars, total_chars, pages, text_pages。"""
        fitz = import_fitz()
        doc = fitz.open(str(path))
        try:
            if doc.needs_pass:
                raise RuntimeError("PDF 已加密，无法转换: %s" % path)
            page_count = doc.page_count
            total = 0
            text_pages = 0
            repl = 0
            for page in doc:
                try:
                    text = page.get_text("text", sort=True) or ""
                except Exception:
                    text = ""
                t = text.strip()
                total += len(t)
                if len(t) >= 30:
                    text_pages += 1
                repl += t.count("\ufffd")
            avg = total / float(max(1, page_count))
            native = (
                page_count > 0
                and avg >= 50
                and text_pages >= max(1, int(page_count * 0.6))
                and repl <= max(2, int(total * 0.002))
            )
            return native, avg, total, page_count, text_pages
        finally:
            doc.close()

    def convert_pdf(self, src: Path, out_md: Path, assets_dir: Path):
        native, avg, total, pages, text_pages = self.analyze_pdf(src)
        LOG.info(
            "PDF 判定: native=%s, avg_chars/page=%.1f, text_pages=%d/%d",
            native,
            avg,
            text_pages,
            pages,
        )
        engine = self.args.pdf_engine
        if engine == "auto":
            # 2026 版 MinerU 支持版面/表格/公式识别，会把公式输出为 $...$ / $$...$$。
            # 因此 auto 模式现在优先使用 MinerU，而不是优先 pymupdf4llm。
            mineru = self._mineru_cmd()
            if mineru:
                try:
                    return self.convert_pdf_mineru(src, out_md, assets_dir, native=native)
                except Exception as e:
                    LOG.warning("MinerU 解析失败，尝试降级: %s", e)
                    if native:
                        return self.convert_pdf_native(src, out_md, assets_dir)
                    if self.args.allow_basic_ocr:
                        return self.convert_pdf_rapidocr(src, out_md, assets_dir)
                    raise
            if native:
                LOG.warning("未检测到 MinerU，数字版 PDF 使用 pymupdf4llm；公式可能无法识别。")
                return self.convert_pdf_native(src, out_md, assets_dir)
            if self.args.allow_basic_ocr:
                return self.convert_pdf_rapidocr(src, out_md, assets_dir)
            raise RuntimeError(
                "该 PDF 判定为扫描件，未找到 MinerU。为避免识别乱码/错字，"
                "程序不会自动输出；请先运行 install_mineru.bat，"
                "或在明确接受人工复核成本时加 --allow-basic-ocr。"
            )
        if engine == "pymupdf":
            if not native:
                raise RuntimeError("该 PDF 无文本层，pymupdf 不能准确提取；请改用 --pdf-engine mineru")
            return self.convert_pdf_native(src, out_md, assets_dir)
        if engine == "mineru":
            if not self._mineru_cmd():
                raise RuntimeError("未找到 MinerU，无法使用 --pdf-engine mineru")
            return self.convert_pdf_mineru(src, out_md, assets_dir, native=native)
        if engine == "rapidocr":
            return self.convert_pdf_rapidocr(src, out_md, assets_dir)
        raise RuntimeError("未知 PDF 引擎: %s" % engine)

    def convert_pdf_native(self, src: Path, out_md: Path, assets_dir: Path):
        try:
            return self.convert_pdf_native_pymupdf4llm(src, out_md, assets_dir)
        except Exception as e:
            LOG.warning("pymupdf4llm 不可用或输出异常，降级为 PyMuPDF 原文提取: %s", e)
            method, kind, extra = self.convert_pdf_native_raw(src, out_md, assets_dir)
            extra["note"] = (
                str(extra.get("note", ""))
                + " 未使用结构增强；已保留可检索的原文和表格，版式可能简化。"
            ).strip()
            return method, kind, extra

    def convert_pdf_native_pymupdf4llm(self, src: Path, out_md: Path, assets_dir: Path):
        import pymupdf4llm  # type: ignore

        assets_dir.mkdir(parents=True, exist_ok=True)
        fitz = import_fitz()
        doc = fitz.open(str(src))
        md_text: Optional[str] = None
        last_err: Optional[Exception] = None
        attempts = [
            dict(write_images=True, image_path=str(assets_dir), table_strategy="lines_strict"),
            dict(write_images=True, image_path=str(assets_dir)),
            dict(write_images=True),
            dict(),
        ]
        try:
            for kwargs in attempts:
                try:
                    md_text = pymupdf4llm.to_markdown(doc, **kwargs)
                    break
                except Exception as e:
                    last_err = e
                    continue
            if md_text is None:
                raise RuntimeError("pymupdf4llm 调用失败: %s" % last_err)
            if isinstance(md_text, list):
                parts = []
                for item in md_text:
                    if isinstance(item, dict):
                        parts.append(str(item.get("text") or item.get("md") or ""))
                    else:
                        parts.append(str(item))
                md_text = "\n\n".join(parts)
            md_text = str(md_text)
            if not normalize_for_compare(md_text):
                raise RuntimeError("pymupdf4llm 输出为空")
            md_text = fix_image_paths(md_text, out_md.parent, assets_dir)
            out_md.write_text(md_text, encoding="utf-8")

            # 回比：pymupdf4llm 输出必须保留原文中的绝大部分字符。
            raw = self.extract_pdf_text(src)
            n1 = normalize_for_compare(raw)
            n2 = normalize_for_compare(md_text)
            if len(n1) >= 20:
                cov = bigram_coverage(n1, n2)
                if cov < 0.85:
                    raise RuntimeError("pymupdf4llm 输出与原始文本层覆盖度异常: %.2f%%" % (cov * 100))
            return "pymupdf4llm", "native", {}
        finally:
            doc.close()

    def convert_pdf_native_raw(self, src: Path, out_md: Path, assets_dir: Path):
        """不依赖 pymupdf4llm 的保底方案：按块抽取文本和表格。"""
        fitz = import_fitz()
        assets_dir.mkdir(parents=True, exist_ok=True)
        doc = fitz.open(str(src))
        try:
            parts: List[str] = []
            for page_no, page in enumerate(doc, 1):
                parts.append("\n<!-- page %d -->\n" % page_no)
                tables: List[Tuple[Any, str]] = []
                try:
                    finder = page.find_tables()
                    for table in getattr(finder, "tables", []) or []:
                        rows = table.extract()
                        if rows and any(any(str(c or "").strip() for c in row) for row in rows):
                            tables.append((table.bbox, rows_to_markdown(rows)))
                except Exception:
                    tables = []

                def intersects(block_bbox: Any) -> bool:
                    try:
                        b = fitz.Rect(block_bbox)
                    except Exception:
                        return False
                    for tb, _ in tables:
                        try:
                            r = fitz.Rect(tb)
                        except Exception:
                            continue
                        if not (b.x1 <= r.x0 or b.x0 >= r.x1 or b.y1 <= r.y0 or b.y0 >= r.y1):
                            return True
                    return False

                elements: List[Tuple[float, float, str, str]] = []
                try:
                    blocks = page.get_text("blocks", sort=True)
                except Exception:
                    blocks = page.get_text("blocks")
                for b in blocks or []:
                    try:
                        x0, y0, x1, y1, text, _block_no, block_type = b[:7]
                    except Exception:
                        continue
                    if block_type != 0:
                        continue
                    text = CONTROL_RE.sub("", str(text or "")).strip()
                    if not text:
                        continue
                    if intersects((x0, y0, x1, y1)):
                        continue
                    elements.append((float(y0), float(x0), "text", text))
                for tb, table_md in tables:
                    try:
                        r = fitz.Rect(tb)
                        elements.append((float(r.y0), float(r.x0), "table", table_md))
                    except Exception:
                        elements.append((0.0, 0.0, "table", table_md))
                elements.sort(key=lambda e: (round(e[0] / 6.0), e[1]))
                for _y, _x, kind, content in elements:
                    if kind == "text":
                        parts.append(content)
                    else:
                        parts.append(content)
                        parts.append("")
            md_text = "\n\n".join(p for p in parts if p and p.strip())
            if not normalize_for_compare(md_text):
                raise RuntimeError("PyMuPDF 未能提取到有效文本")
            out_md.write_text(fix_image_paths(md_text, out_md.parent, assets_dir), encoding="utf-8")
            return "pymupdf-raw", "native", {"note": "使用 PyMuPDF 原文/表格提取"}
        finally:
            doc.close()

    def _mineru_cmd(self) -> Optional[List[str]]:
        if self.args.mineru_cmd:
            p = which(self.args.mineru_cmd)
            if p:
                return [str(p)]
        module_cmd = find_mineru_module()
        if module_cmd:
            return module_cmd
        return find_mineru()

    def convert_pdf_mineru(
        self, src: Path, out_md: Path, assets_dir: Path, native: bool
    ):
        mineru = self._mineru_cmd()
        if not mineru:
            raise RuntimeError("未找到 MinerU。推荐运行 install_mineru.bat 后重试。")
        assets_dir.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="mineru_") as td:
            tmp_dir = Path(td)
            text = ""
            md_file: Optional[Path] = None

            # 先检测新版 MinerU CLI（Typer：mineru parse ...）。
            # 新版命令是：
            #   mineru server start
            #   mineru parse <文件> --pages all --output <md> --format markdown --wait N --force
            is_new_cli = False
            try:
                run_command(mineru + ["parse", "--help"], timeout=60)
                is_new_cli = True
            except Exception as e:
                LOG.info("未检测到新版 MinerU parse 命令，按旧版 CLI 尝试: %s", e)

            if is_new_cli:
                try:
                    # parse 依赖本地 MinerU server；已启动时该命令会直接返回。
                    run_command(mineru + ["server", "start"], timeout=180)
                except Exception as e:
                    LOG.warning("MinerU server start 返回异常，继续尝试 parse: %s", e)
                try:
                    mode_out = run_command(
                        mineru + ["config", "get", "parse_server.local.mode"], timeout=30
                    ).stdout or ""
                    if "disabled" in mode_out.lower():
                        LOG.warning(
                            "MinerU parse_server.local.mode=disabled；当前可能走远程解析，"
                            "公式/离线场景请执行：mineru config set parse_server.local.mode managed"
                        )
                except Exception:
                    pass

                tmp_md = tmp_dir / (src.stem + ".md")
                cmd_new = [
                    *mineru,
                    "parse",
                    str(src),
                    "--pages",
                    "all",
                    "--format",
                    "markdown",
                    "--output",
                    str(tmp_md),
                    "--wait",
                    str(self.args.mineru_timeout),
                    "--force",
                ]
                if self.args.mineru_tier:
                    cmd_new.extend(["--tier", self.args.mineru_tier])
                new_cli_error: Optional[Exception] = None
                max_attempts = max(3, min(60, int(self.args.mineru_timeout // 60) or 3))
                for attempt in range(1, max_attempts + 1):
                    try:
                        run_command(cmd_new, timeout=self.args.mineru_timeout + 180)
                        if tmp_md.exists() and tmp_md.stat().st_size > 0:
                            text = tmp_md.read_text(encoding="utf-8", errors="replace")
                            # 新版 CLI 会把图片导出到输出 md 附近的 images/ 目录。
                            for d in tmp_md.parent.rglob("*"):
                                if d.is_dir() and d.name.lower() == "images":
                                    for img in d.rglob("*"):
                                        if img.is_file():
                                            try:
                                                rel = img.relative_to(d)
                                            except ValueError:
                                                rel = Path(img.name)
                                            dest = assets_dir / rel
                                            dest.parent.mkdir(parents=True, exist_ok=True)
                                            shutil.copy2(str(img), str(dest))
                        if text.strip():
                            break
                    except Exception as e:
                        new_cli_error = e
                        msg = str(e).lower()
                        retryable = any(
                            key in msg
                            for key in (
                                "503",
                                "parse_failed",
                                "connecterror",
                                "connection",
                                "quality_tier_unavailable",
                                "service_unavailable",
                                "server",
                            )
                        )
                        if attempt < max_attempts and retryable:
                            LOG.warning(
                                "MinerU parse 服务尚未就绪（%d/%d），15 秒后重试: %s",
                                attempt,
                                max_attempts,
                                e,
                            )
                            time.sleep(15)
                            continue
                        raise RuntimeError("新版 MinerU parse 失败: %s" % e) from e
                    if not text.strip() and attempt < max_attempts:
                        LOG.warning("MinerU parse 未生成输出（%d/%d），15 秒后重试", attempt, max_attempts)
                        time.sleep(15)
                        continue
                if not text.strip():
                    if new_cli_error is not None:
                        raise RuntimeError("新版 MinerU parse 失败: %s" % new_cli_error)
                    raise RuntimeError("新版 MinerU parse 未生成有效 Markdown 文件")

            if not is_new_cli and not text.strip():
                # 旧版 MinerU 2.x：mineru -p <file> -o <dir> -m auto -l ch
                outdir = tmp_dir / "old_cli"
                outdir.mkdir(parents=True, exist_ok=True)
                variants = [
                    mineru + ["-p", str(src), "-o", str(outdir), "-m", "auto", "-l", "ch"],
                    mineru + ["-p", str(src), "-o", str(outdir)],
                ]
                last_err: Optional[Exception] = None
                for cmd in variants:
                    try:
                        run_command(cmd, timeout=self.args.mineru_timeout)
                        last_err = None
                        break
                    except Exception as e:
                        last_err = e
                        LOG.warning("旧版 MinerU 命令失败，尝试下一种参数: %s", e)
                if last_err is not None:
                    raise RuntimeError("MinerU 执行失败: %s" % last_err)

                md_files = [p for p in outdir.rglob("*.md") if p.is_file()]
                if not md_files:
                    raise RuntimeError("MinerU 未生成 Markdown 文件")
                md_file = max(md_files, key=lambda p: p.stat().st_mtime)
                text = md_file.read_text(encoding="utf-8", errors="replace")

                # 旧版 CLI 通常输出 images/ 目录，复制到目标 assets 并保留相对结构。
                for d in md_file.parent.rglob("*"):
                    if d.is_dir() and d.name.lower() == "images":
                        for img in d.rglob("*"):
                            if img.is_file():
                                try:
                                    rel = img.relative_to(d)
                                except ValueError:
                                    rel = Path(img.name)
                                dest = assets_dir / rel
                                dest.parent.mkdir(parents=True, exist_ok=True)
                                shutil.copy2(str(img), str(dest))

            if not text.strip():
                raise RuntimeError("MinerU 未生成有效 Markdown 内容")

            for old in ("](images/", "](./images/", "](images\\", "](.\\images\\"):
                text = text.replace(old, "](%s/" % assets_dir.name)
            text = fix_image_paths(text, out_md.parent, assets_dir)
            out_md.write_text(text, encoding="utf-8")
            return (
                "mineru",
                "native" if native else "ocr",
                {
                    "note": "MinerU 高精度解析"
                    + ("" if native else "（扫描件 OCR，仍需人工核对原页）"),
                },
            )

    def convert_pdf_rapidocr(self, src: Path, out_md: Path, assets_dir: Path):
        try:
            from rapidocr_onnxruntime import RapidOCR  # type: ignore
        except Exception as e:
            raise RuntimeError(
                "未安装 RapidOCR。运行 pip install rapidocr-onnxruntime opencv-python-headless numpy"
            ) from e

        fitz = import_fitz()
        assets_dir.mkdir(parents=True, exist_ok=True)
        engine = RapidOCR()
        doc = fitz.open(str(src))
        parts: List[str] = []
        confs: List[float] = []
        try:
            for page_no, page in enumerate(doc, 1):
                pix = page.get_pixmap(dpi=self.args.dpi, colorspace=fitz.csRGB, alpha=False)
                img_path = assets_dir / ("page_%04d.png" % page_no)
                pix.save(str(img_path))
                try:
                    ret = engine(str(img_path))
                    # RapidOCR 常见返回 (result, elapse)；部分版本可能直接返回 result。
                    if isinstance(ret, tuple) and len(ret) >= 1:
                        result = ret[0]
                    else:
                        result = ret
                except Exception as e:
                    LOG.warning("RapidOCR 第 %d 页失败: %s", page_no, e)
                    result = None
                if not result:
                    parts.append(
                        "\n<!-- page %d OCR 无文字；原页图: %s/%s -->\n"
                        % (page_no, assets_dir.name, img_path.name)
                    )
                    continue
                lines = self.group_rapidocr_result(result)
                page_confs = [x["score"] for x in lines if x.get("score") is not None]
                if page_confs:
                    confs.extend(page_confs)
                avg_conf = sum(page_confs) / float(len(page_confs)) if page_confs else 0.0
                parts.append(
                    "\n<!-- page %d OCR avg_conf=%.4f；原页图: %s/%s -->\n"
                    % (page_no, avg_conf, assets_dir.name, img_path.name)
                )
                parts.append("\n".join(x["text"] for x in lines if x["text"].strip()))
            md_text = "\n\n".join(p for p in parts if p and p.strip())
            if not normalize_for_compare(md_text):
                raise RuntimeError("RapidOCR 未识别出有效文字")
            out_md.write_text(fix_image_paths(md_text, out_md.parent, assets_dir), encoding="utf-8")
            return (
                "rapidocr-basic",
                "ocr",
                {
                    "min_conf": min(confs) if confs else 0.0,
                    "note": "基础 OCR：已保存逐页原图，必须人工复核；扫描件建议改用 MinerU。",
                },
            )
        finally:
            doc.close()

    @staticmethod
    def group_rapidocr_result(result: Any) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = []
        for entry in result or []:
            try:
                box, text, score = entry[0], entry[1], entry[2]
            except Exception:
                continue
            if not box or not str(text).strip():
                continue
            try:
                xs = [float(p[0]) for p in box]
                ys = [float(p[1]) for p in box]
            except Exception:
                continue
            items.append(
                {
                    "x0": min(xs),
                    "y0": min(ys),
                    "x1": max(xs),
                    "y1": max(ys),
                    "text": str(text).strip(),
                    "score": float(score) if score is not None else 0.0,
                }
            )
        items.sort(key=lambda d: (d["y0"], d["x0"]))
        lines: List[Dict[str, Any]] = []
        for it in items:
            cy = (it["y0"] + it["y1"]) / 2.0
            h = max(1.0, it["y1"] - it["y0"])
            best: Optional[Dict[str, Any]] = None
            for line in lines:
                if abs(line["cy"] - cy) <= max(8.0, line["h"] * 0.6):
                    if best is None or abs(line["cy"] - cy) < abs(best["cy"] - cy):
                        best = line
            if best is None:
                lines.append({"cy": cy, "h": h, "items": [it]})
            else:
                best["items"].append(it)
                best["h"] = max(best["h"], h)
                ys = [(x["y0"] + x["y1"]) / 2.0 for x in best["items"]]
                best["cy"] = sum(ys) / float(len(ys))
        out: List[Dict[str, Any]] = []
        for line in lines:
            line["items"].sort(key=lambda d: d["x0"])
            chunks: List[str] = []
            prev: Optional[Dict[str, Any]] = None
            for it in line["items"]:
                if prev is not None:
                    gap = it["x0"] - prev["x1"]
                    h = max(8.0, min(it["y1"] - it["y0"], prev["y1"] - prev["y0"]))
                    if gap > h * 0.8:
                        chunks.append("  ")
                chunks.append(it["text"])
                prev = it
            text = "".join(chunks).strip()
            if not text:
                continue
            score = sum(x["score"] for x in line["items"]) / float(len(line["items"]))
            out.append({"text": text, "score": score, "cy": line["cy"]})
        out.sort(key=lambda d: d["cy"])
        return out

    # ---- Word -------------------------------------------------------------

    def convert_docx(self, src: Path, out_md: Path, assets_dir: Path):
        errors: List[str] = []
        if self.pandoc:
            try:
                return self._convert_docx_pandoc(src, out_md, assets_dir)
            except Exception as e:
                errors.append("Pandoc: %s" % e)
                LOG.warning("Pandoc 转换失败，尝试 Mammoth: %s", e)
        try:
            return self._convert_docx_mammoth(src, out_md, assets_dir)
        except Exception as e:
            errors.append("Mammoth: %s" % e)
            LOG.warning("Mammoth 转换失败，尝试 python-docx: %s", e)
        try:
            return self._convert_docx_python_docx(src, out_md, assets_dir)
        except Exception as e:
            errors.append("python-docx: %s" % e)
        raise RuntimeError("DOCX 转换失败。\n" + "\n".join(errors))

    def _convert_docx_pandoc(self, src: Path, out_md: Path, assets_dir: Path):
        assets_dir.mkdir(parents=True, exist_ok=True)
        cmd = [
            str(self.pandoc),
            str(src),
            "-t",
            "gfm",
            "--wrap=none",
            "--extract-media",
            str(assets_dir),
            "-o",
            str(out_md),
        ]
        run_command(cmd, timeout=1800)
        if not out_md.exists() or out_md.stat().st_size == 0:
            raise RuntimeError("Pandoc 未生成输出文件")
        text = out_md.read_text(encoding="utf-8", errors="replace")
        out_md.write_text(fix_image_paths(text, out_md.parent, assets_dir), encoding="utf-8")
        return "pandoc", "office", {}

    def _convert_docx_mammoth(self, src: Path, out_md: Path, assets_dir: Path):
        import mammoth  # type: ignore
        from markdownify import markdownify as to_markdown  # type: ignore

        assets_dir.mkdir(parents=True, exist_ok=True)
        counter = itertools.count(1)

        def save_image(image):
            with image.open() as f:
                data = f.read()
            ext = mimetypes.guess_extension(image.content_type or "") or ".png"
            if ext.lower() in (".jpe", ".jpeg"):
                ext = ".jpg"
            name = "image%03d%s" % (next(counter), ext)
            (assets_dir / name).write_bytes(data)
            return {"src": "%s/%s" % (assets_dir.name, name)}

        convert_image = mammoth.images.img_element(save_image)
        with src.open("rb") as f:
            result = mammoth.convert_to_html(f, convert_image=convert_image)
        md_text = to_markdown(result.value or "", heading_style="ATX", bullets="-")
        if not normalize_for_compare(md_text):
            raise RuntimeError("Mammoth 输出为空")
        out_md.write_text(fix_image_paths(md_text, out_md.parent, assets_dir), encoding="utf-8")
        return "mammoth+markdownify", "office", {"note": "Pandoc 不可用，使用 Mammoth 回退。"}

    def _convert_docx_python_docx(self, src: Path, out_md: Path, assets_dir: Path):
        from docx import Document  # type: ignore
        from docx.table import Table  # type: ignore
        from docx.text.paragraph import Paragraph  # type: ignore

        doc = Document(str(src))
        parts: List[str] = []
        for child in doc.element.body.iterchildren():
            tag = str(getattr(child, "tag", ""))
            if tag.endswith("}p"):
                p = Paragraph(child, doc)
                text = p.text.strip()
                if not text:
                    continue
                style = getattr(getattr(p, "style", None), "name", "") or ""
                low = style.lower()
                if "heading" in low or "标题" in style:
                    level = 1
                    for n in range(1, 7):
                        if str(n) in style or str(n) in low:
                            level = n
                            break
                    parts.append("#" * level + " " + text)
                elif "list" in low or "列表" in style:
                    parts.append("- " + text)
                else:
                    parts.append(text)
            elif tag.endswith("}tbl"):
                try:
                    table = Table(child, doc)
                    rows = [
                        [cell.text.strip().replace("\n", "<br>") for cell in row.cells]
                        for row in table.rows
                    ]
                    md_table = rows_to_markdown(rows)
                    if md_table:
                        parts.append(md_table)
                except Exception:
                    continue
        md_text = "\n\n".join(parts)
        if not normalize_for_compare(md_text):
            raise RuntimeError("python-docx 未提取到有效文本")
        out_md.write_text(md_text, encoding="utf-8")
        return (
            "python-docx",
            "office",
            {"note": "未安装 Pandoc/Mammoth；仅保留文本、标题和表格，复杂版式可能丢失。"},
        )

    def convert_doc(self, src: Path, out_md: Path, assets_dir: Path):
        with tempfile.TemporaryDirectory(prefix="doc2docx_") as td:
            temp_dir = Path(td)
            docx_path: Optional[Path] = None
            errors: List[str] = []
            if self.soffice:
                try:
                    cmd = [
                        str(self.soffice),
                        "--headless",
                        "--norestore",
                        "--convert-to",
                        "docx",
                        "--outdir",
                        str(temp_dir),
                        str(src),
                    ]
                    run_command(cmd, timeout=1800)
                    candidates = [p for p in temp_dir.glob("*.docx") if p.is_file()]
                    if candidates:
                        docx_path = candidates[0]
                except Exception as e:
                    errors.append("LibreOffice: %s" % e)
            if docx_path is None and has_module("win32com"):
                try:
                    docx_path = self._convert_doc_with_word(src, temp_dir)
                except Exception as e:
                    errors.append("Word COM: %s" % e)
            if docx_path is None:
                raise RuntimeError(
                    "旧版 .doc 转换需要 LibreOffice(soffice) 或已安装 Microsoft Word + pywin32。\n"
                    + "\n".join(errors)
                )
            method, kind, extra = self.convert_docx(docx_path, out_md, assets_dir)
            note = str(extra.get("note", "")).strip()
            extra["note"] = (note + " 由 .doc 转为 .docx 后处理。").strip()
            return ("doc->" + method, kind, extra)

    @staticmethod
    def _convert_doc_with_word(src: Path, temp_dir: Path) -> Path:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore

        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(src.resolve()), ReadOnly=True)
            out = temp_dir / (src.stem + ".docx")
            doc.SaveAs2(str(out), FileFormat=16)  # wdFormatDocumentDefault -> docx
            doc.Close(SaveChanges=0)
            return out
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    # ---- 其他格式 ---------------------------------------------------------

    def convert_text(self, src: Path, out_md: Path, assets_dir: Path):
        text = read_text_auto(src)
        out_md.write_text(text, encoding="utf-8")
        return "copy", "text", {}

    def convert_html(self, src: Path, out_md: Path, assets_dir: Path):
        if self.pandoc:
            try:
                cmd = [
                    str(self.pandoc),
                    str(src),
                    "-t",
                    "gfm",
                    "--wrap=none",
                    "-o",
                    str(out_md),
                ]
                run_command(cmd, timeout=600)
                if out_md.exists() and out_md.stat().st_size > 0:
                    return "pandoc", "office", {}
            except Exception as e:
                LOG.warning("Pandoc 转换 HTML 失败，尝试 BeautifulSoup: %s", e)
        if has_module("bs4") and has_module("markdownify"):
            from bs4 import BeautifulSoup  # type: ignore
            from markdownify import markdownify as to_markdown  # type: ignore

            soup = BeautifulSoup(read_text_auto(src), "html.parser")
            md_text = to_markdown(str(soup), heading_style="ATX")
            out_md.write_text(md_text, encoding="utf-8")
            return "bs4+markdownify", "office", {"note": "HTML 经本地解析转换。"}
        raise RuntimeError("HTML 转换需要 Pandoc 或 beautifulsoup4+markdownify")

    # ---- 回比/QA ----------------------------------------------------------

    def extract_pdf_text(self, path: Path) -> str:
        fitz = import_fitz()
        doc = fitz.open(str(path))
        try:
            parts = []
            for page in doc:
                try:
                    parts.append(page.get_text("text", sort=True) or "")
                except Exception:
                    parts.append("")
            return "\n".join(parts)
        finally:
            doc.close()

    def extract_docx_text(self, path: Path) -> str:
        try:
            from docx import Document  # type: ignore
        except Exception:
            return ""
        try:
            doc = Document(str(path))
        except Exception:
            return ""
        parts: List[str] = []
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text)
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                parts.append(" ".join(x for x in cells if x))
        return "\n".join(parts)

    def verify(
        self,
        src: Path,
        out_md: Path,
        method: str,
        kind: str,
        extra: Dict[str, Any],
        rec: ConversionRecord,
    ) -> None:
        if not out_md.exists():
            raise FileNotFoundError("输出文件不存在: %s" % out_md)
        text = out_md.read_text(encoding="utf-8", errors="replace")
        rec.output_chars = len(text)
        rec.suspicious = len(SUSPICIOUS_RE.findall(text))

        if kind == "ocr":
            rec.status = "REVIEW_OCR"
            rec.note = (rec.note + " 扫描/OCR 输出，必须人工核对原页。").strip()
            if extra.get("min_conf") is not None:
                rec.min_conf = float(extra["min_conf"])
        else:
            raw = ""
            suffix = src.suffix.lower()
            if suffix == ".pdf" and kind == "native":
                raw = self.extract_pdf_text(src)
            elif suffix == ".docx":
                raw = self.extract_docx_text(src)
            elif suffix in (".txt", ".md"):
                raw = read_text_auto(src)
            if raw:
                n1 = normalize_for_compare(raw)
                n2 = normalize_for_compare(text)
                if len(n1) >= 20:
                    rec.source_chars = len(n1)
                    rec.coverage = bigram_coverage(n1, n2)
                    if rec.coverage < 0.97:
                        rec.status = "REVIEW_TEXT_MISMATCH"
                        rec.note = (
                            rec.note
                            + " 原文/输出字符覆盖度 %.2f%%，请核对。"
                            % (rec.coverage * 100)
                        ).strip()
            elif suffix in (".pdf", ".docx") and kind == "native":
                rec.status = "REVIEW_NO_SOURCE_CHECK"
                rec.note = (rec.note + " 无法自动回比原文文本层，请人工抽查。").strip()
            if rec.suspicious > 0:
                rec.status = "REVIEW_SUSPICIOUS_CHARS"
                rec.note = (
                    rec.note + " 发现可疑替换字符/私用区字符 %d 个。" % rec.suspicious
                ).strip()
            if extra.get("force_review"):
                rec.status = "REVIEW_LAYOUT"
                rec.note = (rec.note + " 版式可能变化，建议抽查。").strip()
            if rec.status == "OK":
                rec.status = "OK"

        if not normalize_for_compare(text) and kind != "ocr":
            rec.status = "FAILED_EMPTY_OUTPUT"
            rec.note = (rec.note + " 输出没有可读文本。").strip()


# ---------------------------------------------------------------------------
# 报告
# ---------------------------------------------------------------------------

def write_reports(
    source: Path,
    output: Path,
    qa_dir: Path,
    records: List[ConversionRecord],
    args: argparse.Namespace,
) -> None:
    qa_dir.mkdir(parents=True, exist_ok=True)

    fieldnames = [
        "source",
        "output",
        "ext",
        "method",
        "kind",
        "status",
        "source_chars",
        "output_chars",
        "coverage",
        "suspicious",
        "min_conf",
        "elapsed",
        "note",
    ]
    rows = []
    for r in records:
        d = asdict(r)
        if d.get("coverage") is not None:
            d["coverage"] = "%.4f" % float(d["coverage"])
        if d.get("min_conf") is not None:
            d["min_conf"] = "%.4f" % float(d["min_conf"])
        d["elapsed"] = "%.2f" % float(d.get("elapsed") or 0)
        rows.append(d)

    with (qa_dir / "qa_report.csv").open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in fieldnames})

    review = [r for r in records if r.status.startswith("REVIEW") or r.status.startswith("FAILED")]
    with (qa_dir / "review_queue.csv").open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in review:
            d = asdict(r)
            d["coverage"] = "" if d["coverage"] is None else "%.4f" % float(d["coverage"])
            d["min_conf"] = "" if d["min_conf"] is None else "%.4f" % float(d["min_conf"])
            d["elapsed"] = "%.2f" % float(d.get("elapsed") or 0)
            w.writerow({k: d.get(k, "") for k in fieldnames})

    write_review_index(qa_dir, output, source, records, args)

    with (qa_dir / "environment.txt").open("w", encoding="utf-8") as f:
        f.write("convert_to_md version: %s\n" % VERSION)
        f.write("python: %s\n" % sys.version.replace("\n", " "))
        f.write("executable: %s\n" % sys.executable)
        f.write("source: %s\n" % source)
        f.write("output: %s\n" % output)
        f.write("pandoc: %s\n" % which("pandoc"))
        f.write("soffice: %s\n" % which("soffice"))
        f.write("mineru: %s\n" % (find_mineru() or "not found"))
        f.write("pymupdf4llm: %s\n" % has_module("pymupdf4llm"))
        f.write("rapidocr_onnxruntime: %s\n" % has_module("rapidocr_onnxruntime"))
        f.write("python-docx: %s\n" % has_module("docx"))
        f.write("mammoth: %s\n" % has_module("mammoth"))
        f.write("win32com: %s\n" % has_module("win32com"))
        f.write("args: %s\n" % vars(args))


def write_review_index(
    qa_dir: Path,
    output: Path,
    source: Path,
    records: List[ConversionRecord],
    args: argparse.Namespace,
) -> None:
    rows = []
    for r in records:
        src_path = (source / r.source).resolve()
        try:
            src_uri = src_path.as_uri()
        except Exception:
            src_uri = str(src_path)
        out_path = (output / r.output).resolve() if r.output else None
        try:
            out_uri = out_path.as_uri() if out_path else ""
        except Exception:
            out_uri = str(out_path or "")
        cls = "review" if (r.status.startswith("REVIEW") or r.status.startswith("FAILED")) else ""
        cov = "" if r.coverage is None else "%.2f%%" % (r.coverage * 100)
        note = html.escape(r.note or "")
        method = html.escape(r.method or "")
        rows.append(
            '<tr class="%s"><td>%s</td><td>%s</td><td>%s</td><td>%s</td>'
            '<td>%s</td><td>%s</td><td><a href="%s">原文件</a> | <a href="%s">Markdown</a></td></tr>'
            % (
                cls,
                html.escape(r.source),
                html.escape(r.status),
                method,
                html.escape(r.kind),
                cov,
                note,
                html.escape(src_uri, quote=True),
                html.escape(out_uri, quote=True),
            )
        )
    html_text = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>文档转换复核清单</title>
<style>
body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:24px;color:#222}
table{border-collapse:collapse;width:100%%;font-size:13px}
th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
th{background:#f5f5f5;text-align:left;position:sticky;top:0}
tr.review{background:#fff6f6}
tr.review td:nth-child(2){color:#b00020;font-weight:600}
a{color:#0b57d0;text-decoration:none}
.small{color:#666;font-size:12px}
</style>
</head>
<body>
<h1>文档转换复核清单</h1>
<p class="small">源目录：%s<br>输出目录：%s<br>生成时间：%s</p>
<p>红色行为需要人工复核（OCR、可疑字符、文本回比覆盖度不足或失败）。建议逐条打开“原文件”和“Markdown”对照。</p>
<table>
<thead><tr><th>源文件</th><th>状态</th><th>方法</th><th>类型</th><th>覆盖度</th><th>说明</th><th>链接</th></tr></thead>
<tbody>
%s
</tbody>
</table>
</body>
</html>
""" % (
        html.escape(str(source)),
        html.escape(str(output)),
        time.strftime("%Y-%m-%d %H:%M:%S"),
        "\n".join(rows),
    )
    (qa_dir / "review_index.html").write_text(html_text, encoding="utf-8")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="批量将 Word/PDF 等文档转换为 Markdown，并对结果做可追溯的 QA 标记。"
    )
    p.add_argument("--source", default=DEFAULT_SOURCE, help="源目录")
    p.add_argument("--output", default=None, help="输出目录，默认在源目录旁创建 <源目录名>_md")
    p.add_argument("--report-dir", default=None, help="QA 报告目录，默认 <输出目录>/_qa")
    p.add_argument(
        "--pdf-engine",
        choices=["auto", "pymupdf", "mineru", "rapidocr"],
        default="auto",
        help="PDF 转换引擎。auto: 优先 MinerU（支持公式识别），失败再降级；pymupdf 只做数字文本提取",
    )
    p.add_argument(
        "--allow-basic-ocr",
        action="store_true",
        help="没有 MinerU 时允许用 RapidOCR 兜底；输出会被标记为需人工复核。",
    )
    p.add_argument("--mineru-cmd", default=None, help="MinerU 命令或路径；默认优先 python -m mineru.cli.main")
    p.add_argument(
        "--mineru-tier",
        choices=["flash", "basic", "standard", "advanced"],
        default=None,
        help="MinerU 解析档位；公式识别建议 standard 或 advanced。默认由 MinerU 自行选择。",
    )
    p.add_argument("--mineru-timeout", type=int, default=7200, help="单文件 MinerU 超时秒数")
    p.add_argument("--dpi", type=int, default=300, help="扫描件 OCR 渲染 DPI")
    p.add_argument("--extensions", default=None, help="只转换指定扩展名，如 pdf,docx")
    p.add_argument("--limit", type=int, default=0, help="只处理前 N 个文件，便于试跑")
    p.add_argument("--resume", dest="resume", action="store_true", default=True, help="跳过已存在的输出")
    p.add_argument("--no-resume", dest="resume", action="store_false", help="不跳过已存在输出")
    p.add_argument("--overwrite", action="store_true", help="覆盖已有输出")
    p.add_argument(
        "--strict",
        action="store_true",
        help="只要有 REVIEW_* 或 FAILED 文件，进程退出码非 0，适合脚本化验收。",
    )
    p.add_argument("--dry-run", action="store_true", help="只列出待处理文件，不转换")
    p.add_argument("--list-engines", action="store_true", help="打印当前检测到的转换引擎")
    return p.parse_args(argv)


def print_engines() -> None:
    print("pymupdf4llm :", has_module("pymupdf4llm"))
    print("rapidocr    :", has_module("rapidocr_onnxruntime"))
    print("python-docx :", has_module("docx"))
    print("mammoth     :", has_module("mammoth"))
    print("markdownify :", has_module("markdownify"))
    print("bs4         :", has_module("bs4"))
    print("win32com    :", has_module("win32com"))
    print("pandoc      :", which("pandoc"))
    print("soffice     :", which("soffice"))
    print("mineru      :", find_mineru())


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    source = Path(args.source).resolve()
    if not source.exists() or not source.is_dir():
        print("源目录不存在: %s" % source)
        return 2
    output = Path(args.output).resolve() if args.output else source.parent / (source.name + "_md")
    if output == source:
        print("输出目录不能与源目录相同。")
        return 2
    output.mkdir(parents=True, exist_ok=True)
    qa_dir = Path(args.report_dir).resolve() if args.report_dir else output / "_qa"
    setup_logging(qa_dir / "convert_to_md.log")

    if args.list_engines:
        print_engines()
        return 0

    converter = Converter(source, output, qa_dir, args)
    if args.dry_run:
        files = converter.collect_files()
        print("共 %d 个文件:" % len(files))
        for p in files:
            print("  ", p.relative_to(source))
        return 0

    records = converter.convert_all()
    write_reports(source, output, qa_dir, records, args)

    total = len(records)
    ok = sum(1 for r in records if r.status in ("OK", "SKIP_EXISTING"))
    review = sum(1 for r in records if r.status.startswith("REVIEW"))
    failed = sum(1 for r in records if r.status.startswith("FAILED"))
    LOG.info("完成: 共 %d，正常/跳过 %d，需复核 %d，失败 %d", total, ok, review, failed)
    LOG.info("报告: %s", qa_dir / "qa_report.csv")
    LOG.info("待复核: %s", qa_dir / "review_queue.csv")
    LOG.info("复核页面: %s", qa_dir / "review_index.html")

    if failed:
        return 1
    if args.strict and review:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
