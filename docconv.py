#!/usr/bin/env python3
"""docconv — PDF, Word, Pages, resim vb. dosyaları Markdown ve JSON'a çevirir.

Kullanım:
  python docconv.py web                      # tarayıcı arayüzü (http://127.0.0.1:5050)
  python docconv.py convert a.pdf b.docx     # CLI; çıktı ./out klasörüne
  python docconv.py convert klasor/ -o cikti --format md
"""
from __future__ import annotations

import argparse
import io
import json
import sys
import tempfile
import zipfile
from pathlib import Path

from converters import SUPPORTED, ConversionError, convert_file


def collect(paths: list[str]) -> list[Path]:
    files: list[Path] = []
    for p in map(Path, paths):
        if p.is_dir():
            files += sorted(f for f in p.rglob("*") if f.suffix.lower() in SUPPORTED and not f.name.startswith("."))
        else:
            files.append(p)
    return files


def cmd_convert(args) -> int:
    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    files = collect(args.files)
    if not files:
        print("Dönüştürülecek dosya bulunamadı.", file=sys.stderr)
        return 1
    failed = 0
    used: set[str] = set()
    for f in files:
        try:
            res = convert_file(f)
        except (ConversionError, Exception) as e:  # noqa: BLE001
            failed += 1
            print(f"✗ {f.name}: {e}", file=sys.stderr)
            continue
        stem = f.stem
        if stem in used:  # aynı isimli farklı uzantılı dosyalar birbirini ezmesin
            stem = f"{f.stem}_{f.suffix.lstrip('.').lower()}"
        used.add(stem)
        if args.format in ("md", "both"):
            (out / f"{stem}.md").write_text(res["markdown"], encoding="utf-8")
        if args.format in ("json", "both"):
            (out / f"{stem}.json").write_text(
                json.dumps(res, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        print(f"✓ {f.name} → {out}/{stem}  ({res['stats']['words']} kelime)")
        for w in res["warnings"]:
            print(f"  ! {w}")
    return 1 if failed == len(files) else 0


# ---------------------------------------------------------------- web
PAGE = """<!doctype html><html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>DocConv</title>
<style>
:root{--bg:#f6f6f4;--fg:#1c1c1a;--mut:#6b6b66;--card:#fff;--line:#e2e2dd;--acc:#2f5bea;--err:#c0392b}
@media(prefers-color-scheme:dark){:root{--bg:#161615;--fg:#ececea;--mut:#9a9a94;--card:#1f1f1e;--line:#33332f;--acc:#7b9cff;--err:#ff7b6b}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 -apple-system,system-ui,sans-serif}
main{max-width:960px;margin:0 auto;padding:32px 16px}
h1{margin:0 0 4px;font-size:24px}p.sub{margin:0 0 20px;color:var(--mut)}
#drop{border:2px dashed var(--line);border-radius:14px;padding:40px 16px;text-align:center;background:var(--card);cursor:pointer;transition:.15s}
#drop.over{border-color:var(--acc);background:color-mix(in srgb,var(--acc) 8%,var(--card))}
#drop b{color:var(--acc)}
.res{background:var(--card);border:1px solid var(--line);border-radius:12px;margin-top:16px;overflow:hidden}
.hd{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:12px 14px;border-bottom:1px solid var(--line)}
.hd .n{font-weight:600;flex:1;min-width:140px;word-break:break-all}.hd small{color:var(--mut)}
button,.tab{font:inherit;border:1px solid var(--line);background:var(--bg);color:var(--fg);border-radius:8px;padding:5px 11px;cursor:pointer}
button:hover{border-color:var(--acc)}.tab.on{background:var(--acc);color:#fff;border-color:var(--acc)}
pre{margin:0;padding:14px;max-height:420px;overflow:auto;white-space:pre-wrap;word-break:break-word;font:13px/1.5 ui-monospace,Menlo,monospace}
.warn{padding:8px 14px;color:var(--mut);font-size:13px;border-bottom:1px solid var(--line)}
.err{color:var(--err);padding:12px 14px}
.bar{display:flex;gap:8px;margin-top:16px;flex-wrap:wrap}
</style></head><body><main>
<h1>DocConv</h1><p class="sub">PDF, Word, Pages ve resim dosyalarını Markdown + JSON'a çevir. Dosyalar bu bilgisayardan çıkmaz.</p>
<div id="drop"><b>Dosyaları sürükle</b> veya tıklayıp seç<br><small style="color:var(--mut)">.pdf .docx .doc .rtf .odt .pages .png .jpg .tiff .txt</small>
<input id="f" type="file" multiple hidden></div>
<div class="bar" id="bar" hidden><button id="zip">Tümünü ZIP indir</button></div>
<div id="out"></div></main>
<script>
const $=s=>document.querySelector(s),drop=$('#drop'),inp=$('#f');let results=[];
drop.onclick=()=>inp.click();inp.onchange=()=>go(inp.files);
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',ev=>go(ev.dataTransfer.files));
const dl=(name,text,type)=>{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type}));a.download=name;a.click()};
const stem=n=>n.replace(/\\.[^.]+$/,'');
async function go(files){
  for(const file of files){
    const box=document.createElement('div');box.className='res';box.innerHTML='<div class="hd"><span class="n"></span><small>Dönüştürülüyor…</small></div>';
    box.querySelector('.n').textContent=file.name;$('#out').prepend(box);
    const fd=new FormData();fd.append('file',file);
    try{
      const r=await fetch('/convert',{method:'POST',body:fd});const j=await r.json();
      if(!r.ok)throw new Error(j.error||'Hata');
      results.push(j);render(box,j);
    }catch(e){box.innerHTML='<div class="hd"><span class="n"></span></div><div class="err"></div>';
      box.querySelector('.n').textContent=file.name;box.querySelector('.err').textContent='✗ '+e.message}
  }
  $('#bar').hidden=!results.length;inp.value='';
}
function render(box,j){
  const md=j.markdown,js=JSON.stringify(j,null,2);
  box.innerHTML=`<div class="hd"><span class="n"></span><small>${j.stats.words} kelime${j.stats.pages?' · '+j.stats.pages+' sayfa':''}</small>
  <button class="tab on" data-t="md">Markdown</button><button class="tab" data-t="json">JSON</button>
  <button data-a="copy">Kopyala</button><button data-a="dlmd">.md</button><button data-a="dljson">.json</button></div>
  ${j.warnings.map(w=>'<div class="warn"></div>').join('')}<pre></pre>`;
  box.querySelector('.n').textContent=j.source;
  box.querySelectorAll('.warn').forEach((e,i)=>e.textContent='! '+j.warnings[i]);
  const pre=box.querySelector('pre');let cur='md';pre.textContent=md;
  box.onclick=e=>{const t=e.target.dataset;
    if(t.t){cur=t.t;pre.textContent=cur=='md'?md:js;box.querySelectorAll('.tab').forEach(b=>b.classList.toggle('on',b.dataset.t==cur))}
    if(t.a=='copy')navigator.clipboard.writeText(cur=='md'?md:js);
    if(t.a=='dlmd')dl(stem(j.source)+'.md',md,'text/markdown');
    if(t.a=='dljson')dl(stem(j.source)+'.json',js,'application/json')}
}
$('#zip').onclick=async()=>{const r=await fetch('/zip',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(results)});
  const b=await r.blob();const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='docconv.zip';a.click()};
</script></body></html>"""


def cmd_web(args) -> int:
    from flask import Flask, jsonify, request, send_file

    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

    @app.get("/")
    def index():
        return PAGE

    @app.post("/convert")
    def convert():
        f = request.files.get("file")
        if not f or not f.filename:
            return jsonify(error="Dosya yok."), 400
        name = Path(f.filename).name
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / name
            f.save(p)
            try:
                return jsonify(convert_file(p))
            except ConversionError as e:
                return jsonify(error=str(e)), 422
            except Exception as e:  # noqa: BLE001
                return jsonify(error=f"Beklenmeyen hata: {e}"), 500

    @app.post("/zip")
    def make_zip():
        buf = io.BytesIO()
        seen: dict[str, int] = {}
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for r in request.get_json():
                stem = Path(r["source"]).stem
                seen[stem] = seen.get(stem, 0) + 1
                if seen[stem] > 1:
                    stem += f"_{seen[stem]}"
                z.writestr(f"{stem}.md", r["markdown"])
                z.writestr(f"{stem}.json", json.dumps(r, ensure_ascii=False, indent=2))
        buf.seek(0)
        return send_file(buf, mimetype="application/zip", download_name="docconv.zip")

    print(f"→ http://{args.host}:{args.port}")
    app.run(host=args.host, port=args.port)
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(prog="docconv", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("convert", help="dosya/klasörleri dönüştür")
    c.add_argument("files", nargs="+")
    c.add_argument("-o", "--output", default="out")
    c.add_argument("--format", choices=["md", "json", "both"], default="both")
    c.set_defaults(fn=cmd_convert)
    w = sub.add_parser("web", help="yerel web arayüzü")
    w.add_argument("--host", default="127.0.0.1")
    w.add_argument("--port", type=int, default=5050)
    w.set_defaults(fn=cmd_web)
    args = ap.parse_args()
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
