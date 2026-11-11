# DocConv

PDF, Word, Pages, resim vb. dosyaları **Markdown** ve **JSON**'a çevirir. Her şey yerelde çalışır.

## Kurulum

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
brew install tesseract tesseract-lang   # resim / taranmış PDF / PDF içi resim OCR'ı için (Türkçe dahil)
```

## Kullanım

```bash
.venv/bin/python docconv.py web                       # http://127.0.0.1:5050 (sürükle-bırak arayüz)
.venv/bin/python docconv.py convert rapor.pdf not.docx -o cikti
.venv/bin/python docconv.py convert klasor/ --format md
```

## Biçim desteği

| Biçim | Yöntem | Not |
|---|---|---|
| PDF | pymupdf4llm + OCR | Başlık/tablo/liste çıkarımı; metinsiz (taranmış) sayfalar ve sayfalardaki gömülü resimler Tesseract ile okunur |
| DOCX | python-docx | Başlıklar, listeler, kalın/italik, tablolar |
| DOC, RTF, ODT | macOS `textutil` → DOCX | Yalnızca macOS; RTF/DOC'ta stil yoksa başlıklar düz metin kalır |
| Pages | IWA içinden metin | **Deneysel**: yalnızca düz metin; biçim ve tablolar yok. En iyisi Pages'ten Word/PDF dışa aktarmak |
| PNG/JPG/TIFF… | Tesseract OCR | `DOCCONV_OCR_LANG=tur` (varsayılan; `tur+eng` karışımı Türkçe harflerde daha kötü sonuç veriyor) |
| TXT, MD | doğrudan | |

## JSON şeması

`source, format, title, metadata, stats, warnings, pages[{number, markdown}], blocks[{type: heading|paragraph|list|table|code|quote, ...}], markdown`
