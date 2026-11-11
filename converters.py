"""Belge -> Markdown + JSON dönüştürücü çekirdeği.

Desteklenen: PDF, DOCX, DOC/RTF/ODT (macOS textutil ile), Pages, resimler (OCR), TXT/MD.
"""
from __future__ import annotations

import io
import os
import re
import shutil
import struct
import subprocess
import tempfile
import zipfile
from pathlib import Path

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
TEXTUTIL_EXT = {".doc", ".rtf", ".odt", ".rtfd", ".wordml"}
SUPPORTED = {".pdf", ".docx", ".pages", ".txt", ".md"} | IMAGE_EXT | TEXTUTIL_EXT


class ConversionError(Exception):
    pass


# --------------------------------------------------------------------------
# Ortak yardımcılar
# --------------------------------------------------------------------------
def md_to_blocks(md: str) -> list[dict]:
    """Markdown metnini yapılandırılmış bloklara böler (JSON çıktısı için)."""
    blocks: list[dict] = []
    para: list[str] = []
    lines = md.splitlines()
    i = 0

    def flush():
        if para:
            blocks.append({"type": "paragraph", "text": " ".join(s.strip() for s in para)})
            para.clear()

    while i < len(lines):
        line = lines[i]
        s = line.strip()
        if not s:
            flush()
        elif m := re.match(r"^(#{1,6})\s+(.*)$", s):
            flush()
            blocks.append({"type": "heading", "level": len(m.group(1)), "text": m.group(2).strip()})
        elif s.startswith("```"):
            flush()
            code = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i])
                i += 1
            blocks.append({"type": "code", "text": "\n".join(code)})
        elif s.startswith("|") and s.endswith("|"):
            flush()
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                    rows.append(cells)
                i += 1
            i -= 1
            blocks.append({"type": "table", "rows": rows})
        elif m := re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$", line):
            flush()
            items = []
            ordered = m.group(2)[0].isdigit()
            while i < len(lines) and (mm := re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$", lines[i])):
                items.append({"level": len(mm.group(1)) // 2, "text": mm.group(3).strip()})
                i += 1
            i -= 1
            blocks.append({"type": "list", "ordered": ordered, "items": items})
        elif s.startswith(">"):
            flush()
            blocks.append({"type": "quote", "text": s.lstrip("> ").strip()})
        else:
            para.append(s)
        i += 1
    flush()
    return blocks


def make_result(path: Path, fmt: str, markdown: str, *, title: str | None = None,
                metadata: dict | None = None, pages: list[dict] | None = None,
                warnings: list[str] | None = None) -> dict:
    markdown = markdown.strip() + "\n"
    blocks = md_to_blocks(markdown)
    if not title:
        title = next((b["text"] for b in blocks if b["type"] == "heading"), path.stem)
    return {
        "source": path.name,
        "format": fmt,
        "title": title,
        "metadata": metadata or {},
        "stats": {
            "characters": len(markdown),
            "words": len(markdown.split()),
            "blocks": len(blocks),
            "pages": len(pages) if pages else None,
        },
        "warnings": warnings or [],
        "pages": pages or [],
        "blocks": blocks,
        "markdown": markdown,
    }


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------
def convert_pdf(path: Path) -> dict:
    import pymupdf
    import pymupdf4llm

    warnings: list[str] = []
    doc = pymupdf.open(path)
    if doc.needs_pass:
        raise ConversionError("PDF şifreli.")
    meta = {k: v for k, v in (doc.metadata or {}).items() if v}
    chunks = pymupdf4llm.to_markdown(doc, page_chunks=True, use_ocr=False)
    pages = []
    ocr_pages = ocr_images = 0
    ocr_missing = False
    for i, c in enumerate(chunks):
        text = c["text"].strip()
        page = doc[i]
        has_text = len(re.sub(r"[\s\-]", "", text)) >= 10
        imgs = _page_images(doc, page)
        if imgs and not _has_tesseract():
            ocr_missing = True
        elif not has_text and (imgs or not text):
            # Metin katmanı yok -> taranmış sayfa; sayfanın tamamını OCR'la.
            ocr = _ocr_pil(_render_page(page))
            if ocr:
                text = ocr
                ocr_pages += 1
        elif imgs:
            # Metinli sayfada gömülü resimler -> her birini OCR'la, sayfa sonuna ekle.
            found = [t for t in (_ocr_pil(im) for im in imgs) if t]
            if found:
                ocr_images += len(found)
                text += "\n\n" + "\n\n".join(
                    "> **[Resimden okunan metin]**\n> " + t.replace("\n", "\n> ") for t in found)
        pages.append({"number": i + 1, "markdown": text})
    md = "\n\n".join(p["markdown"] for p in pages if p["markdown"])
    if ocr_pages:
        warnings.append(f"{ocr_pages} sayfada metin katmanı yoktu, sayfalar OCR ile okundu.")
    if ocr_images:
        warnings.append(f"{ocr_images} gömülü resimden OCR ile metin çıkarıldı.")
    if ocr_missing:
        warnings.append("PDF'de resim var ancak Tesseract kurulu değil, resimler okunamadı "
                        "(`brew install tesseract tesseract-lang`).")
    return make_result(path, "pdf", md, title=meta.get("title") or None,
                       metadata=meta, pages=pages, warnings=warnings)


def _page_images(doc, page, min_px: int = 80):
    """Sayfadaki gömülü resimleri PIL görüntüsü olarak döndürür (küçük süsleri atlar)."""
    from PIL import Image
    import pymupdf

    out, seen = [], set()
    for info in page.get_images(full=True):
        xref = info[0]
        if xref in seen:
            continue
        seen.add(xref)
        try:
            pix = pymupdf.Pixmap(doc, xref)
            if pix.width < min_px or pix.height < min_px:
                continue
            if pix.colorspace is None or pix.colorspace.n not in (1, 3) or pix.alpha:
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix) if pix.alpha == 0 and pix.colorspace else \
                    pymupdf.Pixmap(pymupdf.csRGB, pymupdf.Pixmap(pix, 0)) if pix.alpha else pix
            out.append(Image.open(io.BytesIO(pix.tobytes("png"))))
        except Exception:  # noqa: BLE001
            continue
    return out


def _render_page(page):
    from PIL import Image
    return Image.open(io.BytesIO(page.get_pixmap(dpi=200).tobytes("png")))


def _ocr_pil(img) -> str:
    if not _has_tesseract():
        return ""
    import pytesseract
    return pytesseract.image_to_string(img, lang=_ocr_lang()).strip()


# --------------------------------------------------------------------------
# DOCX
# --------------------------------------------------------------------------
def _runs_to_md(paragraph) -> str:
    out = []
    for run in paragraph.runs:
        t = run.text
        if not t:
            continue
        if not t.strip():
            out.append(t)
            continue
        lead = t[: len(t) - len(t.lstrip())]
        trail = t[len(t.rstrip()):]
        core = t.strip()
        if run.bold and run.italic:
            core = f"***{core}***"
        elif run.bold:
            core = f"**{core}**"
        elif run.italic:
            core = f"*{core}*"
        out.append(lead + core + trail)
    text = "".join(out)
    return re.sub(r"\*\*\s*\*\*|\*\*(\s+)\*\*", r"\1", text).strip()


def convert_docx(path: Path) -> dict:
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    d = docx.Document(str(path))
    cp = d.core_properties
    meta = {k: str(v) for k, v in {
        "title": cp.title, "author": cp.author, "created": cp.created,
        "modified": cp.modified, "subject": cp.subject, "keywords": cp.keywords,
    }.items() if v}

    lines: list[str] = []
    for child in d.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            p = Paragraph(child, d)
            text = _runs_to_md(p)
            if not text:
                lines.append("")
                continue
            style = (p.style.name or "") if p.style is not None else ""
            low = style.lower()
            if low == "title":
                lines.append(f"# {text}")
            elif m := re.match(r"(?:heading|başlık)\s*(\d)", low):
                lines.append(f"{'#' * min(int(m.group(1)), 6)} {text}")
            elif "list bullet" in low or "liste madde" in low:
                lvl = re.search(r"(\d)$", low)
                lines.append(f"{'  ' * (int(lvl.group(1)) - 1 if lvl else 0)}- {text}")
            elif "list number" in low or "liste numara" in low:
                lvl = re.search(r"(\d)$", low)
                lines.append(f"{'  ' * (int(lvl.group(1)) - 1 if lvl else 0)}1. {text}")
            elif child.find(".//{*}numPr") is not None:
                lines.append(f"- {text}")
            elif "quote" in low or "alıntı" in low:
                lines.append(f"> {text}")
            else:
                lines.append(text)
            lines.append("")
        elif tag == "tbl":
            t = Table(child, d)
            rows = []
            for r in t.rows:
                cells, seen = [], set()
                for c in r.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    cells.append(c.text.strip().replace("\n", " ").replace("|", "\\|"))
                rows.append(cells)
            if rows:
                w = max(len(r) for r in rows)
                rows = [r + [""] * (w - len(r)) for r in rows]
                lines.append("| " + " | ".join(rows[0]) + " |")
                lines.append("|" + "---|" * w)
                for r in rows[1:]:
                    lines.append("| " + " | ".join(r) + " |")
                lines.append("")
    md = re.sub(r"\n{3,}", "\n\n", "\n".join(lines))
    # Ardışık liste maddeleri arasındaki boş satırı kaldır (sıkı liste)
    md = re.sub(r"(^[ ]*(?:- |1\. ).*)\n\n(?=[ ]*(?:- |1\. ))", r"\1\n", md, flags=re.M)
    warnings = []
    n_img = len(d.inline_shapes)
    if n_img:
        warnings.append(f"Belgede {n_img} resim var; Markdown çıktısına dahil edilmedi.")
    return make_result(path, "docx", md, title=meta.get("title"), metadata=meta, warnings=warnings)


# --------------------------------------------------------------------------
# DOC / RTF / ODT  (macOS textutil -> docx)
# --------------------------------------------------------------------------
def convert_via_textutil(path: Path) -> dict:
    tu = shutil.which("textutil")
    if not tu:
        raise ConversionError(f"{path.suffix} için macOS 'textutil' gerekir (veya dosyayı DOCX olarak kaydedin).")
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / (path.stem + ".docx")
        r = subprocess.run([tu, "-convert", "docx", "-output", str(out), str(path)],
                           capture_output=True, text=True)
        if r.returncode != 0 or not out.exists():
            raise ConversionError(f"textutil başarısız: {r.stderr.strip()}")
        res = convert_docx(out)
    res["source"] = path.name
    res["format"] = path.suffix.lstrip(".").lower()
    return res


# --------------------------------------------------------------------------
# Pages (.pages) — IWA içinden metin çıkarma
# --------------------------------------------------------------------------
def _snappy_decompress(data: bytes) -> bytes:
    """Ham (framing'siz) Snappy açıcı."""
    pos = 0
    length = shift = 0
    while True:
        b = data[pos]; pos += 1
        length |= (b & 0x7F) << shift
        if not b & 0x80:
            break
        shift += 7
    out = bytearray()
    while pos < len(data):
        tag = data[pos]; pos += 1
        kind = tag & 3
        if kind == 0:
            n = tag >> 2
            if n >= 60:
                nb = n - 59
                n = int.from_bytes(data[pos:pos + nb], "little"); pos += nb
            n += 1
            out += data[pos:pos + n]; pos += n
        else:
            if kind == 1:
                n = ((tag >> 2) & 7) + 4
                off = ((tag >> 5) << 8) | data[pos]; pos += 1
            elif kind == 2:
                n = (tag >> 2) + 1
                off = int.from_bytes(data[pos:pos + 2], "little"); pos += 2
            else:
                n = (tag >> 2) + 1
                off = int.from_bytes(data[pos:pos + 4], "little"); pos += 4
            if off == 0 or off > len(out):
                raise ConversionError("Bozuk Snappy verisi.")
            for _ in range(n):
                out.append(out[-off])
    return bytes(out)


def _iwa_decompress(raw: bytes) -> bytes:
    out, pos = bytearray(), 0
    while pos < len(raw):
        if raw[pos] != 0:
            raise ConversionError("Beklenmeyen IWA başlığı.")
        n = int.from_bytes(raw[pos + 1:pos + 4], "little")
        pos += 4
        out += _snappy_decompress(raw[pos:pos + n])
        pos += n
    return bytes(out)


def _varint(buf: bytes, pos: int) -> tuple[int, int]:
    v = shift = 0
    while True:
        b = buf[pos]; pos += 1
        v |= (b & 0x7F) << shift
        if not b & 0x80:
            return v, pos
        shift += 7


def _pb_fields(buf: bytes):
    """Protobuf alanlarını (alan_no, wire_type, değer) olarak üretir."""
    pos = 0
    while pos < len(buf):
        key, pos = _varint(buf, pos)
        fn, wt = key >> 3, key & 7
        if wt == 0:
            v, pos = _varint(buf, pos)
        elif wt == 1:
            v = buf[pos:pos + 8]; pos += 8
        elif wt == 2:
            n, pos = _varint(buf, pos)
            v = buf[pos:pos + n]; pos += n
        elif wt == 5:
            v = buf[pos:pos + 4]; pos += 4
        else:
            return
        yield fn, wt, v


def _iwa_texts(buf: bytes) -> list[str]:
    """IWA akışındaki TSWP.StorageArchive (tip 2001) metinlerini toplar."""
    texts, pos = [], 0
    while pos < len(buf):
        n, pos = _varint(buf, pos)
        info = buf[pos:pos + n]; pos += n
        lengths = []
        for fn, wt, v in _pb_fields(info):
            if fn == 2 and wt == 2:  # message_info
                mtype = mlen = None
                for f2, w2, v2 in _pb_fields(v):
                    if f2 == 1 and w2 == 0:
                        mtype = v2
                    elif f2 == 3 and w2 == 0:
                        mlen = v2
                lengths.append((mtype, mlen or 0))
        for mtype, mlen in lengths:
            payload = buf[pos:pos + mlen]; pos += mlen
            if mtype == 2001:
                for f, w, v in _pb_fields(payload):
                    if f == 3 and w == 2:
                        try:
                            texts.append(v.decode("utf-8"))
                        except UnicodeDecodeError:
                            pass
    return texts


def convert_pages(path: Path) -> dict:
    if not zipfile.is_zipfile(path):
        raise ConversionError("Bu bir Pages paketi değil (zip değil). Klasör paketi ise önce zip'leyin.")
    warnings: list[str] = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        iwa = [n for n in names if n.endswith(".iwa") and "Index/Document" in n] or \
              [n for n in names if n.endswith(".iwa") and n.startswith("Index/") and "Tables" not in n]
        texts: list[str] = []
        for n in iwa:
            try:
                texts += _iwa_texts(_iwa_decompress(z.read(n)))
            except Exception as e:  # noqa: BLE001
                warnings.append(f"{n} okunamadı: {e}")
        preview = next((n for n in names if re.search(r"(preview|thumbnail)\.(jpg|png)$", n, re.I)), None)
        preview_bytes = z.read(preview) if preview else None

    paragraphs: list[str] = []
    for t in texts:
        for line in re.split(r"[\n  ]", t.replace("￼", "")):
            line = line.strip()
            if line:
                paragraphs.append(line)
    md = "\n\n".join(paragraphs)

    if not md and preview_bytes and _has_tesseract():
        from PIL import Image
        import pytesseract
        md = pytesseract.image_to_string(Image.open(io.BytesIO(preview_bytes)), lang=_ocr_lang()).strip()
        warnings.append("Metin çıkarılamadı; yalnızca ilk sayfa önizlemesine OCR uygulandı.")
    if not md:
        raise ConversionError("Pages dosyasından metin çıkarılamadı. Pages'te Dosya > Dışa Aktar > "
                              "Word/PDF yapıp onu dönüştürmeniz önerilir.")
    warnings.append("Pages desteği deneyseldir: başlık/biçim bilgisi ve tablolar korunmaz, yalnızca metin çıkarılır.")
    return make_result(path, "pages", md, warnings=warnings)


# --------------------------------------------------------------------------
# Resim (OCR)
# --------------------------------------------------------------------------
def _has_tesseract() -> bool:
    return shutil.which("tesseract") is not None


def _ocr_lang() -> str:
    pref = os.environ.get("DOCCONV_OCR_LANG", "tur+eng")
    try:
        import pytesseract
        have = set(pytesseract.get_languages())
        use = [l for l in pref.split("+") if l in have]
        return "+".join(use) or "eng"
    except Exception:  # noqa: BLE001
        return "eng"


def convert_image(path: Path) -> dict:
    from PIL import Image

    img = Image.open(path)
    meta = {"width": img.width, "height": img.height, "mode": img.mode, "format": img.format}
    if not _has_tesseract():
        raise ConversionError("Resimden metin çıkarmak için Tesseract OCR gerekir: "
                              "`brew install tesseract tesseract-lang`")
    import pytesseract
    text = pytesseract.image_to_string(img, lang=_ocr_lang()).strip()
    warnings = [] if text else ["Resimde okunabilir metin bulunamadı."]
    return make_result(path, "image", text, metadata=meta, warnings=warnings)


# --------------------------------------------------------------------------
# Düz metin
# --------------------------------------------------------------------------
def convert_text(path: Path) -> dict:
    for enc in ("utf-8-sig", "cp1254", "latin-1"):
        try:
            return make_result(path, path.suffix.lstrip("."), path.read_text(encoding=enc))
        except UnicodeDecodeError:
            continue
    raise ConversionError("Metin kodlaması anlaşılamadı.")


# --------------------------------------------------------------------------
def convert_file(path: str | Path) -> dict:
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".pdf":
        return convert_pdf(path)
    if ext == ".docx":
        return convert_docx(path)
    if ext in TEXTUTIL_EXT:
        return convert_via_textutil(path)
    if ext == ".pages":
        return convert_pages(path)
    if ext in IMAGE_EXT:
        return convert_image(path)
    if ext in {".txt", ".md"}:
        return convert_text(path)
    raise ConversionError(f"Desteklenmeyen dosya türü: {ext or '(uzantı yok)'}")
