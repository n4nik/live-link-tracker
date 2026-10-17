#!/usr/bin/env python3
"""Live Link Tracker: paste vendor live links -> download Excel.

Run:  python3 live_tracker.py      then open http://localhost:8765
Columns: SNO | PROFILE NAME | PROFILE LINK | FOLLOWERS | LIVE LINK | LIVE DATE
"""
import os, io, re, json, base64, urllib.request, urllib.parse
from datetime import datetime, timezone
from concurrent.futures import ThreadPoolExecutor
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
                    "(KHTML, like Gecko) Version/17.0 Safari/605.1.15",
      "Accept-Language": "en-US,en;q=0.9"}
LINK_RE = re.compile(r"instagram\.com/(?:([A-Za-z0-9._]+)/)?(reel|reels|p|tv)/([A-Za-z0-9_-]+)")


def parse_link(raw):
    """Return (username|None, shortcode, clean_link) or None."""
    m = LINK_RE.search(raw.strip())
    if not m:
        return None
    user, kind, code = m.groups()
    if user in ("reel", "reels", "p", "tv"):  # no username in url
        user = None
    kind = "reel" if kind == "reels" else kind
    return user, code, f"https://www.instagram.com/{kind}/{code}/"


def date_from_shortcode(code):
    """Instagram media id embeds the post time (ms since 2011-08-24)."""
    try:
        n = 0
        for ch in code[:11]:
            n = n * 64 + ALPHABET.index(ch)
        ts = (n >> 23) + 1314220021721
        return datetime.fromtimestamp(ts / 1000, tz=timezone.utc).astimezone().strftime("%d-%m-%Y")
    except ValueError:
        return ""


BOT_UAS = ["facebookexternalhit/1.1 (+http://www.facebook.com/externalhit_uatext.php)",
           "Twitterbot/1.0",
           "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)"]


def fetch(url, ua=None):
    try:
        req = urllib.request.Request(url, headers=dict(UA, **({"User-Agent": ua} if ua else {})))
        return urllib.request.urlopen(req, timeout=15).read().decode("utf-8", "ignore")
    except Exception:
        return ""


def username_from_page(clean_link):
    html = fetch(clean_link)
    for pat in (r'"owner":\{[^}]*?"username":"([A-Za-z0-9._]+)"',
                r'instagram\.com/([A-Za-z0-9._]+)/(?:reel|p)/',
                r'\(@([A-Za-z0-9._]+)\)'):
        m = re.search(pat, html)
        if m and m.group(1) not in ("reel", "p"):
            return m.group(1)
    return None


def to_number(s):
    s = s.replace(",", "").strip()
    mult = {"K": 1e3, "M": 1e6, "B": 1e9}.get(s[-1:].upper(), 1)
    return int(float(s[:-1] if mult != 1 else s) * mult)


def followers_of(username):
    for ua in (None, *BOT_UAS):  # Instagram serves bots when it blocks browsers
        html = fetch(f"https://www.instagram.com/{username}/", ua)
        m = re.search(r'content="([\d.,]+[KMB]?) Followers', html, re.I)
        if m:
            return to_number(m.group(1))
    return None


def build_rows(text, followers=True):
    def one(p):
        user, code, clean = p
        user = user or username_from_page(clean)
        return {
            "name": user or "",
            "plink": f"https://www.instagram.com/{user}/" if user else "",
            "followers": followers_of(user) if (user and followers) else None,
            "live": clean,
            "date": date_from_shortcode(code),
        }
    parsed = [parse_link(l) for l in re.split(r"[\s,]+", text) if l]
    with ThreadPoolExecutor(4) as ex:
        rows = list(ex.map(one, [p for p in parsed if p]))
    counts = {}
    for r in rows:
        counts[r["live"]] = counts.get(r["live"], 0) + 1
    for r in rows:
        r["dup"] = counts[r["live"]] > 1
    return rows


def make_xlsx(rows):
    wb = Workbook(); ws = wb.active; ws.title = "Live Links"
    ws.append(["SNO", "PROFILE NAME", "PROFILE LINK", "FOLLOWERS", "LIVE LINK", "LIVE DATE"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(horizontal="center")
    gap = PatternFill("solid", fgColor="FFF2CC")  # yellow = fill manually
    for i, r in enumerate(rows, 1):
        ws.append([i, r["name"], r["plink"], r["followers"], r["live"], r["date"]])
        for col in (2, 3, 4):
            if ws.cell(i + 1, col).value in (None, ""):
                ws.cell(i + 1, col).fill = gap
        if r.get("dup"):  # duplicate live link -> whole row red
            for col in range(1, 7):
                c = ws.cell(i + 1, col)
                c.fill = PatternFill("solid", fgColor="FF9999")
                c.font = Font(color="9C0006", bold=True)
        if r["followers"]:
            ws.cell(i + 1, 4).number_format = "#,##0"
    for col, w in zip("ABCDEF", (6, 26, 40, 14, 48, 14)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def fill_plan(xlsx_bytes, text):
    """Paste live link + date next to each page in an uploaded plan."""
    wb = load_workbook(io.BytesIO(xlsx_bytes)); ws = wb.active
    heads = {str(c.value).strip().upper(): c.column for c in ws[1] if c.value}
    name_col = next((v for k, v in heads.items() if "NAME" in k), None)
    link_col = next((v for k, v in heads.items() if "LINK" in k), None)
    by_user = {}
    for r in range(2, ws.max_row + 1):
        keys = []
        if name_col and ws.cell(r, name_col).value:
            keys.append(str(ws.cell(r, name_col).value))
        if link_col and ws.cell(r, link_col).value:
            m = re.search(r"instagram\.com/([A-Za-z0-9._]+)", str(ws.cell(r, link_col).value))
            if m: keys.append(m.group(1))
        for k in keys:
            by_user.setdefault(k.strip().lower().lstrip("@"), r)
    orows = ws.max_row
    lc, dc = ws.max_column + 1, ws.max_column + 2
    hdr = ws.cell(1, 1)
    for col, t in ((lc, "LIVE LINK"), (dc, "LIVE DATE")):
        c = ws.cell(1, col, t); c.font = hdr.font.copy(); c.fill = hdr.fill.copy(); c.alignment = hdr.alignment.copy()
    ws.column_dimensions[ws.cell(1, lc).column_letter].width = 48
    ws.column_dimensions[ws.cell(1, dc).column_letter].width = 14
    red = PatternFill("solid", fgColor="FF9999"); redf = Font(color="9C0006", bold=True)
    rows = build_rows(text, followers=False)
    seen, extra = {}, []
    for x in rows:
        r = by_user.get(x["name"].lower())
        if x["live"] in seen:  # duplicate live link -> red
            for rr in {seen[x["live"]], r} - {None}:
                for col in (lc, dc):
                    ws.cell(rr, col).fill = red; ws.cell(rr, col).font = redf
            extra.append((x["name"], x["live"], x["date"], "DUPLICATE LINK"))
            continue
        if r is None:
            extra.append((x["name"], x["live"], x["date"], "Page not found in plan" if x["name"] else "Could not read username"))
        elif ws.cell(r, lc).value:
            ws.cell(r, lc).fill = red; ws.cell(r, lc).font = redf
            extra.append((x["name"], x["live"], x["date"], "Page already has a live link (2nd link)"))
        else:
            ws.cell(r, lc, x["live"]); ws.cell(r, dc, x["date"])
            seen[x["live"]] = r
    if extra:
        e = wb.create_sheet("Not pasted")
        e.append(["PAGE NAME", "LIVE LINK", "LIVE DATE", "REASON"])
        for row in extra: e.append(row)
        for col, w in zip("ABCD", (26, 48, 14, 40)): e.column_dimensions[col].width = w
        for c in e[1]: c.font = Font(bold=True)
    buf = io.BytesIO(); wb.save(buf)
    table = [[str(ws.cell(r, name_col).value or "") if name_col else "", ws.cell(r, lc).value or "",
              ws.cell(r, dc).value or ""] for r in range(2, orows + 1)]
    return buf.getvalue(), len(rows) - len(extra), len(extra), table, [list(e) for e in extra]


PAGE = b"""<!doctype html><meta charset=utf-8><title>Live Link Tracker</title>
<meta name=viewport content="width=device-width,initial-scale=1">
<style>
body{font:16px system-ui;max-width:1000px;margin:30px auto;padding:0 16px}
button{padding:9px 16px;font-size:15px;cursor:pointer}
table{border-collapse:collapse;width:100%;font-size:13px;margin-top:10px}
th{background:#1F3864;color:#fff;padding:6px;text-align:left;position:sticky;top:0}
td{border:1px solid #ccc;padding:5px;max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
tr.dup td{background:#FF9999;color:#9C0006;font-weight:bold}
td.gap{background:#FFF2CC}
#wrap{max-height:520px;overflow:auto}
</style>
<body>
<h2>Live Link Tracker</h2>
<p>Paste the vendor's live links (one per line). Anything after <code>?</code> is removed. Duplicates are marked red.</p>
<textarea id=links rows=9 style="width:100%;font:14px monospace" placeholder="https://www.instagram.com/reel/XXXX/?igsh=..."></textarea>
<p><b>Option A &ndash; New tracker</b><br><button onclick="go('tracker')">Make tracker table</button></p>
<p><b>Option B &ndash; Paste into an existing plan</b><br>
<input type=file id=plan accept=".xlsx"> <button onclick="go('plan')">Fill plan with live links</button></p>
<p id=msg style="color:#444"></p>
<div id=out style="display:none">
 <button onclick="dl()">Download Excel</button>
 <button id=cp1 onclick="cp(0,true)">Copy table</button>
 <button id=cp2 onclick="cp(1,false)" style="display:none">Copy LIVE LINK + DATE only</button>
 <span id=cmsg style="color:green"></span>
 <p id=note style="font-size:13px;color:#666"></p>
 <div id=wrap></div><div id=extra></div>
</div>
<script>
let MODE='',XLSX='',HEAD=[];
const esc=s=>String(s??'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
async function go(mode){
  const msg=document.getElementById('msg'); const links=document.getElementById('links').value;
  if(!links.trim()){msg.textContent='Paste some links first.';return}
  let file='';
  if(mode=='plan'){
    const f=document.getElementById('plan').files[0];
    if(!f){msg.textContent='Choose the plan .xlsx first.';return}
    file=await new Promise(r=>{const fr=new FileReader();fr.onload=()=>r(fr.result.split(',')[1]);fr.readAsDataURL(f)});
  }
  msg.textContent='Working... (looking up pages, can take a minute)';
  const res=await fetch('/'+mode,{method:'POST',body:JSON.stringify({links,file})});
  if(!res.ok){msg.textContent='Error: '+await res.text();return}
  const d=await res.json(); MODE=mode; XLSX=d.xlsx||'';
  let h='';
  if(mode=='tracker'){
    HEAD=['SNO','PROFILE NAME','PROFILE LINK','FOLLOWERS','LIVE LINK','LIVE DATE'];
    h='<tr>'+HEAD.map(x=>'<th>'+x+'</th>').join('')+'</tr>'+d.rows.map((r,i)=>
     '<tr'+(r.dup?' class=dup':'')+'><td>'+(i+1)+'</td><td'+(r.name?'':' class=gap')+'>'+esc(r.name)+'</td><td'+(r.plink?'':' class=gap')+'>'+esc(r.plink)+
     '</td><td contenteditable'+(r.followers?'':' class=gap')+'>'+esc(r.followers)+'</td><td>'+esc(r.live)+'</td><td>'+esc(r.date)+'</td></tr>').join('');
    document.getElementById('cp2').style.display='none';
    document.getElementById('note').textContent='Yellow cells are blank: click one and type to fill it (e.g. followers). Copy table copies everything you see.';
    msg.textContent='Done: '+d.rows.length+' links.';
  } else {
    HEAD=['PAGE NAME','LIVE LINK','LIVE DATE'];
    h='<tr>'+HEAD.map(x=>'<th>'+x+'</th>').join('')+'</tr>'+d.table.map(r=>
     '<tr><td>'+esc(r[0])+'</td><td'+(r[1]?'':' class=gap')+'>'+esc(r[1])+'</td><td'+(r[2]?'':' class=gap')+'>'+esc(r[2])+'</td></tr>').join('');
    document.getElementById('cp2').style.display='';
    document.getElementById('note').textContent='Rows are in the same order as your plan. "Copy LIVE LINK + DATE only" gives two columns you can paste right next to your plan in Excel.';
    msg.textContent='Done: '+d.summary;
    document.getElementById('extra').innerHTML=d.extra.length?'<h4>Not pasted</h4><table><tr><th>PAGE</th><th>LINK</th><th>DATE</th><th>REASON</th></tr>'+d.extra.map(e=>'<tr><td>'+e.map(esc).join('</td><td>')+'</td></tr>').join('')+'</table>':'';
  }
  document.getElementById('wrap').innerHTML='<table>'+h+'</table>';
  document.getElementById('out').style.display='';
}
function rowsText(from,withHead){
  const rows=[...document.querySelectorAll('#wrap tr')].slice(withHead?0:1);
  return rows.map(r=>[...r.cells].slice(from).map(c=>c.innerText.replace(/[\\t\\r\\n]+/g,' ').trim()).join(String.fromCharCode(9))).join(String.fromCharCode(10));
}
async function cp(from,withHead){
  const t=rowsText(from,withHead||(from>0&&false));
  try{await navigator.clipboard.writeText(t)}catch(e){const a=document.createElement('textarea');a.value=t;document.body.appendChild(a);a.select();document.execCommand('copy');a.remove()}
  document.getElementById('cmsg').textContent=' Copied! Now paste in Excel (Ctrl/Cmd+V).';
}
async function dl(){
  let b64=XLSX;
  if(MODE=='tracker'){
    const rows=[...document.querySelectorAll('#wrap tr')].slice(1).map(r=>({name:r.cells[1].innerText,plink:r.cells[2].innerText,followers:r.cells[3].innerText.replace(/,/g,''),live:r.cells[4].innerText,date:r.cells[5].innerText,dup:r.className=='dup'}));
    const res=await fetch('/export',{method:'POST',body:JSON.stringify({rows})}); b64=(await res.json()).xlsx;
  }
  const bin=atob(b64),u=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([u]));a.download=MODE=='plan'?'plan_with_live_links.xlsx':'live_links.xlsx';a.click();
}
</script></body>"""


class H(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200); self.send_header("Content-Type", "text/html"); self.end_headers()
        self.wfile.write(PAGE)

    def do_POST(self):
        try:
            n = int(self.headers["Content-Length"])
            if n > 8_000_000: raise ValueError("File too large (max ~5 MB)")
            d = json.loads(self.rfile.read(n))
            if len(re.findall(r"instagram\.com", d.get("links", ""))) > 300: raise ValueError("Max 300 links per run")
            b64 = lambda b: base64.b64encode(b).decode()
            if self.path == "/plan":
                data, ok, bad, table, extra = fill_plan(base64.b64decode(d["file"]), d["links"])
                out = {"xlsx": b64(data), "table": table, "extra": extra,
                       "summary": f"{ok} pasted into plan, {bad} not pasted."}
            elif self.path == "/export":
                rows = d["rows"]
                for r in rows:
                    f = str(r.get("followers") or "").strip()
                    r["followers"] = int(float(f)) if re.fullmatch(r"[\d.]+", f) else None
                out = {"xlsx": b64(make_xlsx(rows))}
            else:
                rows = build_rows(d["links"])
                out = {"rows": rows, "xlsx": b64(make_xlsx(rows))}
            body = json.dumps(out).encode()
        except Exception as ex:
            self.send_response(500); self.end_headers(); self.wfile.write(str(ex).encode()); return
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers(); self.wfile.write(body)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    host = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
    print(f"Open http://localhost:{port}")
    ThreadingHTTPServer((host, port), H).serve_forever()
