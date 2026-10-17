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
:root{--bg:#141414;--card:#1f1f1f;--ink:#f5f5f1;--mute:#a3a3a3;--line:#333;--brand:#e50914;--brand2:#b20710;--ok:#46d369;--head:#000;--zebra:#262626}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 "Netflix Sans","Helvetica Neue",Helvetica,Arial,Inter,sans-serif}
.hero{background:radial-gradient(ellipse at 20% 0,rgba(229,9,20,.55),transparent 55%),linear-gradient(180deg,#000,#141414);color:#fff;padding:34px 16px 84px;border-bottom:1px solid #222}
.wrap{max-width:1040px;margin:0 auto}
.brand{font-weight:900;letter-spacing:3px;color:var(--brand);font-size:14px;text-transform:uppercase;margin-bottom:26px}
.hero h1{margin:0 0 10px;font-size:48px;line-height:1.05;font-weight:900;letter-spacing:-1.5px}
.hero h1 b{color:var(--brand)}
.hero p{margin:0;color:#d2d2d2;max-width:600px;font-size:17px}
.pill{display:inline-block;border:1px solid #555;border-radius:4px;padding:2px 10px;font-size:12px;margin:18px 0 0;color:#ddd}
main{max-width:1040px;margin:-56px auto 40px;padding:0 16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:22px;box-shadow:0 14px 40px rgba(0,0,0,.6);margin-bottom:18px}
.step{display:flex;align-items:center;gap:10px;font-weight:800;font-size:18px;margin-bottom:12px}
.step i{display:inline-grid;place-items:center;width:26px;height:26px;border-radius:4px;background:var(--brand);color:#fff;font-style:normal;font-size:13px;font-weight:900}
textarea{width:100%;min-height:150px;border:1px solid #444;border-radius:4px;background:#333;color:#fff;padding:12px;font:13px/1.5 ui-monospace,Menlo,monospace;resize:vertical}
textarea:focus{outline:none;border-color:#fff;box-shadow:0 0 0 2px rgba(229,9,20,.6)}
.hint{color:var(--mute);font-size:13px;margin:8px 0 0}
.opts{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:720px){.opts{grid-template-columns:1fr}.hero h1{font-size:34px}}
.opt{border:1px solid var(--line);border-radius:8px;padding:16px;background:#181818;transition:.2s}
.opt:hover{transform:scale(1.02);border-color:#666;box-shadow:0 10px 30px rgba(0,0,0,.7)}
.opt h3{margin:0 0 4px;font-size:17px}.opt p{margin:0 0 12px;color:var(--mute);font-size:13px}
button{border:0;border-radius:4px;padding:11px 20px;font:700 15px inherit;cursor:pointer;transition:.15s}
.pri{background:var(--brand);color:#fff}
.pri:hover{background:#f6121d;transform:scale(1.03)}
.sec{background:rgba(109,109,110,.7);color:#fff}.sec:hover{background:rgba(109,109,110,.45)}
.file{display:block;border:2px dashed #555;border-radius:6px;padding:12px;text-align:center;color:var(--mute);cursor:pointer;margin-bottom:12px;font-size:13px}
.file:hover{border-color:var(--brand);color:#fff}.file input{display:none}
#msg{font-weight:700;color:#fff;min-height:22px;margin:0 0 6px}
.bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:8px}
#cmsg{color:var(--ok);font-weight:700;font-size:13px}
.tw{max-height:540px;overflow:auto;border:1px solid var(--line);border-radius:6px}
table{border-collapse:collapse;width:100%;font-size:13px}
th{background:var(--head);color:#fff;padding:10px;text-align:left;position:sticky;top:0;font-size:11px;letter-spacing:1px;border-bottom:2px solid var(--brand)}
td{border-top:1px solid var(--line);padding:8px 10px;max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
tr:nth-child(even) td{background:var(--zebra)}
tr.dup td{background:#e50914!important;color:#fff;font-weight:700}
td.gap{background:#5a4a00!important;color:#ffe27a}
td[contenteditable]:focus{outline:2px solid #fff}
footer{text-align:center;color:var(--mute);font-size:12px;padding:10px 16px 34px;border-top:1px solid #222}
footer b{color:#fff}
</style>
<div class=hero><div class=wrap><div class=brand>Nikhilesh Bhatia</div>
<h1>Live Link <b>Tracker</b></h1>
<p>Paste the vendor's live links and get a clean tracker, or drop them straight into your plan. Duplicates are flagged in red.</p><span class=pill>Instagram campaign tool &middot; Made by Nikhilesh Bhatia</span></div></div>
<main>
<div class=card>
 <div class=step><i>1</i>Paste live links</div>
 <textarea id=links placeholder="https://www.instagram.com/reel/XXXXXXXX/?igsh=...&#10;https://www.instagram.com/p/YYYYYYYY/"></textarea>
 <p class=hint>One link per line. Extra characters after <b>?</b> are removed automatically.</p>
</div>
<div class=card>
 <div class=step><i>2</i>Choose what to make</div>
 <div class=opts>
  <div class=opt><h3>New tracker</h3><p>SNO, profile, followers, live link and live date in one table.</p>
   <button class="pri" onclick="go('tracker')">Make tracker table</button></div>
  <div class=opt><h3>Fill my plan</h3><p>Upload your plan Excel and the live links are pasted next to each page.</p>
   <label class=file><input type=file id=plan accept=".xlsx"><span id=fn>Click to choose plan .xlsx</span></label>
   <button class="pri" onclick="go('plan')">Fill plan with live links</button></div>
 </div>
</div>
<p id=msg></p>
<div class=card id=out style="display:none">
 <div class=step><i>3</i>Your result</div>
 <div class=bar>
  <button class="pri" onclick="dl()">Download Excel</button>
  <button class="sec" id=cp1 onclick="cp(0,true)">Copy table</button>
  <button class="sec" id=cp2 onclick="cp(1,false)" style="display:none">Copy LIVE LINK + DATE only</button>
  <span id=cmsg></span>
 </div>
 <p class=hint id=note></p>
 <div class=tw id=wrap></div><div id=extra></div>
</div>
</main>
<footer>Red = duplicate link &middot; Yellow = missing, click to type<br><br>Tool made by <b>Nikhilesh Bhatia</b><br>&copy; 2026 Nikhilesh Bhatia. All rights reserved. Copyrighted &mdash; unauthorised copying or resale is prohibited.</footer>
<script>
let MODE='',XLSX='',HEAD=[];
document.getElementById('plan').onchange=e=>{document.getElementById('fn').textContent=e.target.files[0]?e.target.files[0].name:'Click to choose plan .xlsx'};
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
</script>"""


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
