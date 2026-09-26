#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 文件检索管理系统 — Web版
# 功能：正文全文检索（不止文件名）· 多文件夹设置 · 多关键词并行(AND/OR) · 新增一键索引 · 全部重建索引
# 启动: python3 workdoc_web.py [端口，默认8765]
# 浏览器打开: http://localhost:8765
import os, sys, json, sqlite3, html, subprocess, argparse, re
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote_plus, urlparse, parse_qs

DB=os.path.join(os.path.dirname(os.path.abspath(__file__)),"workdoc_index.db")
try:
    sys.path.insert(0,os.path.join(os.path.dirname(os.path.abspath(__file__)),"wpslibs"))
except: pass
DEFAULT_FOLDER="/Users/hh/Documents/工作文档"

def extract_text(path):
    ext=path.lower().rsplit(".",1)[-1]
    try:
        if ext=="pdf":
            import pypdf
            return "\n".join((p.extract_text() or "") for p in pypdf.PdfReader(path).pages)
        if ext=="docx":
            import docx
            d=docx.Document(path); parts=[p.text for p in d.paragraphs]
            for t in d.tables:
                for r in t.rows: parts += [c.text for c in r.cells]
            return "\n".join(parts)
        if ext=="pptx":
            import pptx
            prs=pptx.Presentation(path); parts=[]
            for s in prs.slides:
                for sh in s.shapes:
                    if hasattr(sh,"text"): parts.append(sh.text)
            return "\n".join(parts)
        if ext=="xlsx":
            import openpyxl
            wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
            return "\n".join(str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None)
        if ext=="xls":
            import xlrd
            return "\n".join(str(c.value) for bk in xlrd.open_workbook(path,on_demand=True) for ws in bk.sheets() for r in ws.all_rows() for c in r if c.value not in (None,""))
        if ext=="doc":
            out=subprocess.run(["textutil","-convert","txt","-output","-",path],capture_output=True)
            return out.stdout.decode("utf-8","replace")
        if ext in ("csv","txt","wps","rtf","html","json","xml","md","et"):
            with open(path,"rb") as f: return f.read().decode("utf-8","replace")
        if ext=="ppt":
            out=subprocess.run(["textutil","-convert","txt","-output","-",path],capture_output=True)
            return out.stdout.decode("utf-8","replace")
    except Exception as e:
        return f"<ERR {e}>"
    return ""

def _sanitize(s):
    return s.encode("utf-8","ignore").decode("utf-8","ignore") if s else s

def _conn():
    return sqlite3.connect(DB, timeout=30)

def get_folders():
    conn=_conn()
    conn.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT)")
    v=conn.execute("SELECT value FROM settings WHERE key='folders'").fetchone()
    conn.close()
    folders=json.loads(v[0]) if v else []
    return folders if folders else [DEFAULT_FOLDER]

def set_folders(folders):
    conn=_conn()
    conn.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT)")
    conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('folders',?)",(json.dumps(folders,ensure_ascii=False),))
    conn.commit(); conn.close()

def _ensure_fts(conn):
    conn.execute("CREATE TABLE IF NOT EXISTS docs(path TEXT,name TEXT,ext TEXT,text TEXT)")
    conn.execute("CREATE VIRTUAL TABLE IF NOT EXISTS docs_fts USING fts5(text, content='docs', content_rowid='rowid', tokenize='trigram')")
    conn.commit()
    if conn.execute("SELECT COUNT(*) FROM docs_fts").fetchone()[0]==0:
        for r in conn.execute("SELECT rowid,text FROM docs"):
            conn.execute("INSERT INTO docs_fts(rowid,text) VALUES(?,?)",(r[0],_sanitize(r[1])))
        conn.commit()

def index():
    conn=_conn()
    conn.execute("DROP TABLE IF EXISTS docs_fts"); conn.commit()
    conn.execute("DROP TABLE IF EXISTS docs"); conn.commit()
    conn.execute("CREATE TABLE docs(path TEXT,name TEXT,ext TEXT,text TEXT)")
    conn.execute("CREATE VIRTUAL TABLE docs_fts USING fts5(text, content='docs', content_rowid='rowid', tokenize='trigram')")
    n=0
    for root in get_folders():
        for cur,dirs,files in os.walk(root):
            dirs[:]=[d for d in dirs if d != "_重复待清理"]
            for f in files:
                p=os.path.join(cur,f)
                txt=_sanitize(extract_text(p))
                c2=conn.execute("INSERT INTO docs VALUES(?,?,?,?)",(p,os.path.basename(f),os.path.splitext(f)[1].lower(),txt))
                conn.execute("INSERT INTO docs_fts(rowid,text) VALUES(?,?)",(c2.lastrowid,txt))
                n+=1
                if n%100==0: conn.commit()
    conn.commit(); conn.close()
    return n

def index_new():
    conn=_conn(); _ensure_fts(conn)
    known=set(r[0] for r in conn.execute("SELECT path FROM docs"))
    added=0
    for root in get_folders():
        for cur,dirs,files in os.walk(root):
            dirs[:]=[d for d in dirs if d != "_重复待清理"]
            for f in files:
                p=os.path.join(cur,f)
                if p in known: continue
                txt=_sanitize(extract_text(p))
                c2=conn.execute("INSERT INTO docs VALUES(?,?,?,?)",(p,os.path.basename(f),os.path.splitext(f)[1].lower(),txt))
                conn.execute("INSERT INTO docs_fts(rowid,text) VALUES(?,?)",(c2.lastrowid,txt))
                added+=1
    conn.commit(); conn.close()
    return added

def search(q, mode="and", top=100):
    ks=[k for k in re.split(r"\s+",q.strip()) if k]
    if not ks: return 0,[]
    conn=_conn(); _ensure_fts(conn)
    like=["%"+k+"%" for k in ks]
    if len(ks)==1 and len(ks[0])>=3:
        fts=[r[0] for r in conn.execute("SELECT rowid FROM docs_fts WHERE docs_fts MATCH ?",(ks[0],))]
        nm=[r[0] for r in conn.execute("SELECT rowid FROM docs WHERE name LIKE ?",(like[0],))]
        cand=list(set(fts+nm))
        if cand:
            ph=",".join("?" for _ in cand)
            cond="(text LIKE ? OR name LIKE ?)"
            cnt=conn.execute(f"SELECT COUNT(*) FROM docs WHERE rowid IN ({ph}) AND {cond}",(cand+like+like)).fetchone()[0]
            rows=conn.execute(f"SELECT path,name,ext,text FROM docs WHERE rowid IN ({ph}) AND {cond} LIMIT ?",(cand+like+like+[top])).fetchall()
            conn.close(); return cnt,rows
    clause="(text LIKE ? OR name LIKE ?)"
    sep=" AND " if mode=="and" else " OR "
    cond=sep.join(clause for _ in ks)
    params=sum([[lk,lk] for lk in like],[])
    cnt=conn.execute(f"SELECT COUNT(*) FROM docs WHERE {cond}",params).fetchone()[0]
    rows=conn.execute(f"SELECT path,name,ext,text FROM docs WHERE {cond} LIMIT ?",params+[top]).fetchall()
    conn.close()
    return cnt,rows

# ============ 页面样式 ============
# 整体：图书馆/藏书风格。软件名居中，宋体45号。
HTML_TOP="""<!doctype html><html><head><meta charset=utf-8><title>文件检索管理系统</title>
<style>
*{box-sizing:border-box}
body{font-family:"宋体",SimSun,serif,-apple-system,BlinkMacSystemFont;background:linear-gradient(160deg,#3a2c1a,#5a4526,#7a6548);min-height:100vh;margin:0;padding:0}
.wrap{max-width:960px;margin:auto;padding:28px 22px 60px}
.header{text-align:center;background:linear-gradient(135deg,#3a2c1a,#5a4526,#7a6548);border-radius:18px;padding:30px 20px 26px;border:1px solid #c9a97a;box-shadow:0 6px 18px rgba(0,0,0,.35)}
.brand{font-family:"宋体",SimSun,serif;font-size:45px;font-weight:bold;color:#f7e9d0;text-align:center;letter-spacing:2px;text-shadow:0 3px 6px rgba(0,0,0,.4)}
.sub{font-size:13px;color:#d9c6a0;margin-top:6px;text-align:center}
.topbar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;background:#fbf3e3;border:1px solid #d8c29a;border-radius:16px;padding:14px 20px;box-shadow:0 4px 12px rgba(0,0,0,.15);margin-top:16px}
.topbar a.home{text-decoration:none;color:#5a3c1f;font-weight:600;border:1px solid #c9a97a;padding:8px 14px;border-radius:10px;background:#f7e9d0;transition:.15s}
.topbar a.home:hover{background:#5a3c1f;color:#f7e9d0}
.topbar .brand2{font-size:18px;font-weight:700;color:#5a3c1f;letter-spacing:1px}
.topbar .sub{font-size:12px;color:#9a7b5c}
h2{margin:22px 0 10px;font-size:26px;color:#f3e6c8;letter-spacing:.5px;text-shadow:0 2px 4px rgba(0,0,0,.35)}
.searchcard{background:#fbf3e3;border-radius:16px;padding:16px 20px;border:1px solid #d8c29a;box-shadow:0 4px 12px rgba(0,0,0,.15);margin-top:18px}
.searchbar{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
input{flex:1 1 300px;padding:13px 16px;font-size:15px;border:2px solid #c9a97a;border-radius:12px;outline:none;background:#fbf3e3;box-shadow:0 2px 6px rgba(0,0,0,.1)}
input:focus{border-color:#8b5f3f;box-shadow:0 0 0 3px rgba(139,95,63,.25)}
.searchbtns{display:flex;gap:10px;align-items:center;margin-top:12px}
select{padding:11px 12px;border-radius:11px;border:2px solid #c9a97a;background:#fbf3e3;font-size:14px;color:#5a3c1f}
button{padding:13px 22px;border-radius:11px;background:#8b5f3f;color:#fbf3e3;border:0;cursor:pointer;font-size:15px;box-shadow:0 3px 8px rgba(0,0,0,.25);transition:.15s}
button:hover{background:#5a3c1f;transform:translateY(-1px)}
button.sec{background:#b08857} button.sec:hover{background:#8b5f3f}
.btnrow{display:flex;gap:10px;flex-wrap:wrap;margin:16px 0 8px}
.cnt{color:#e6d3ae;font-size:14px;margin:10px 0 12px}
table{width:100%;border-collapse:collapse;margin-top:8px;background:#fbf3e3;border-radius:14px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,.15)}
td{vertical-align:top;padding:14px 16px;border-bottom:1px solid #e2d2b0;font-size:14px;color:#4a3a1f}
tr:hover{background:#f3e3c0}
.name{font-weight:600;color:#3a2c1a;font-size:16px}
.path{color:#8b7359;font-size:12px;word-break:break-all}
.snp{color:#5a4a2e;font-size:13px;line-height:1.55;border-left:3px solid #c9a97a;padding-left:9px;margin-top:6px}
.tag{display:inline-block;background:#e8d6b8;color:#5a3c1f;font-size:11px;padding:3px 10px;border-radius:8px;margin-bottom:8px}
.shelf{display:flex;flex-wrap:wrap;gap:12px;margin-top:18px}
.shelfcard{flex:1 1 220px;background:#fbf3e3;border-radius:14px;padding:16px;border:1px solid #d8c29a;box-shadow:0 3px 8px rgba(0,0,0,.15)}
.shelfcard b{color:#5a3c1f}
.foldadd{background:#fbf3e3;border-radius:14px;padding:16px;border:1px solid #d8c29a;margin-top:14px;box-shadow:0 3px 8px rgba(0,0,0,.15)}
.foldadd input{width:100%}
.guide{border:2px solid #c9a97a;border-radius:16px;background:#fbf3e3;padding:18px 22px;margin-top:24px;box-shadow:0 4px 12px rgba(0,0,0,.15)}
.guide b{color:#5a3c1f;font-size:16px;display:block;margin-bottom:6px}
.guide ul{list-style:none;padding:0;margin:8px 0 0;color:#5a4a2e;font-size:14px;line-height:1.7}
.guide li{margin:4px 0}
.foot{color:#d9c6a0;font-size:13px;margin-top:20px;text-align:center;border-top:1px solid #9a7b5c;padding-top:14px}
</style></head><body><div class="wrap">
<div class="header"><div class="brand">📚 文件检索管理系统</div><div class="sub">正文全文检索 · 不止文件名 · 多关键词并行</div></div>"""

GUIDE_HTML=f"""
<div class="guide"><b>📖 使用说明</b><ul>
<li><b>检索</b>：输入关键词（可空格分隔多个关键词并行），全部命中选 AND，任一命中选 OR；支持文件名与正文全文。</li>
<li><b>新增文件一键索引</b>：把新文件放进已设置的文件来源文件夹，点此按钮只补新文件，几秒钟完成。</li>
<li><b>全部文件重建索引</b>：需要全量刷新（更换文件夹、清理重复等）时点此按钮，耗时较长（约20分钟），期间检索会短暂暂缓。</li>
<li><b>设置文件来源</b>：添加/移除要索引的文件夹完整路径；可同时设置多个文件夹。</li>
<li><b>使用注意</b>：`_重复待清理` 文件夹自动跳过；索引保存在本程序目录下的 `workdoc_index.db`；本服务面向 macOS（旧 .doc/.ppt 依赖 textutil）。</li>
</ul></div>"""

def snip(text,q):
    i=text.find(q)
    if i<0: return ""
    return text[max(0,i-60):i+120].replace("\n"," ")

def _qs(u):
    return parse_qs(urlparse(unquote_plus(u)).query) if "?" in u else {}

class H(BaseHTTPRequestHandler):
    def log_message(self,*a): pass
    def do_GET(self):
        u=self.path
        qs=_qs(u)
        if u.startswith("/search"):
            q=unquote_plus(qs.get("q",[""])[0])[:80]
            mode=qs.get("mode",["and"])[0] if qs.get("mode") else "and"
            if mode not in ("and","or"): mode="and"
            cnt,rows=search(q,mode)
            body=HTML_TOP.replace("__Q__", html.escape(q))
            body+='<div class="topbar"><a class="home" href="/">← 返回首页</a><span class="brand2">🔍 检索结果</span></div>'
            body+=f'<div class="cnt">命中 <b>{cnt}</b> 份（只显示前100） · 匹配方式：{"全部命中" if mode=="and" else "任一命中"}</div><table>'
            for path,name,ext,text in rows[:100]:
                body+=f'<tr><td><span class="tag">{html.escape(ext)}</span><div class="name">{html.escape(name)}</div><div class="path">{html.escape(path)}</div><div class="snp">{html.escape(snip(text,q))}</div></td></tr>'
            body+="</table><div class='foot'>文件检索管理系统 · 多关键词（空格分隔），全部命中选 AND，任一命中选 OR</div></body></html>"
            self.send_response(200); self.send_header("Content-Type","text/html; charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/settings"):
            folders=get_folders()
            body=HTML_TOP.replace("__Q__","")
            body+='<div class="topbar"><a class="home" href="/">← 返回首页</a><span class="brand2">📁 设置文件来源</span></div>'
            body+='<div class="shelf">'+''.join(f'<div class="shelfcard">📁 <b>{html.escape(f)}</b><br><a href="/remove_folder?path={html.escape(f)}" style="color:#8b5f3f">移除</a></div>' for f in folders)+'</div>'
            body+='<div class="foldadd"><form method=get action="/add_folder"><input name="path" placeholder="输入要加入的文件夹完整路径，如 /Users/xxx/文档"><button>添加文件夹</button></form></div>'
            body+=GUIDE_HTML
            body+='<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a> · 添加后点“新增文件一键索引”补索引新文件</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/add_folder"):
            p=unquote_plus(qs.get("path",[""])[0]).strip()
            body=HTML_TOP.replace("__Q__","")
            if p and os.path.isdir(p):
                fs=get_folders()
                if p not in fs: set_folders(fs+[p])
                body+=f'<div class="cnt">已添加文件夹：<b>{html.escape(p)}</b>。点“新增文件一键索引”补索引新文件。</div>'
            else:
                body+=f'<div class="cnt">❌ 路径不可用或不存在：<b>{html.escape(p)}</b>（请填完整路径）</div>'
            body+='<div class="topbar"><a class="home" href="/">← 返回首页</a><span class="brand2">➕ 添加文件夹</span></div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">管理文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/remove_folder"):
            p=unquote_plus(qs.get("path",[""])[0]).strip()
            fs=[f for f in get_folders() if f!=p]
            set_folders(fs)
            body=HTML_TOP.replace("__Q__","")
            body+=f'<div class="cnt">已移除文件夹：<b>{html.escape(p)}</b>（索引中已收录的文件仍可搜到，下次重建时剔除）。</div>'
            body+='<div class="topbar"><a class="home" href="/">← 返回首页</a><span class="brand2">➖ 移除文件夹</span></div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">管理文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/index_new"):
            n=index_new()
            body=HTML_TOP.replace("__Q__","")
            body+=f'<div class="cnt">新增文件一键索引完成：新增 <b>{n}</b> 份</div>'
            body+='<div class="topbar"><a class="home" href="/">← 返回首页</a><span class="brand2">📥 新增文件一键索引</span></div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">设置文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/reindex"):
            n=index()
            body=HTML_TOP.replace("__Q__","")
            body+=f'<div class="cnt">全部文件重建索引完成：<b>{n}</b> 份（含全部文件来源）</div>'
            body+='<div class="topbar"><a class="home" href="/">← 返回首页</a><span class="brand2">🔁 全部文件重建索引</span></div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">设置文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        else:
            n=get_folders()
            body=HTML_TOP.replace("__Q__","")
            body+=f'<div class="searchcard"><form method=get action="/search"><div class="searchbar"><input name=q placeholder="输入关键词，空格分隔多个关键词并行检索" value=""></div><div class="searchbtns"><select name="mode"><option value="and" selected>全部命中(AND)</option><option value="or">任一命中(OR)</option></select><button>检索</button></div></form></div>'
            body+=f'<div class="btnrow"><form method=get action="/index_new"><button>📥 新增文件一键索引</button></form><form method=get action="/reindex"><button class="sec">🔁 全部文件重建索引</button></form><a href="/settings" style="text-decoration:none"><button class="sec">📁 设置文件来源</button></a></div>'
            body+=f'<div class="cnt">当前文件来源文件夹：{" · ".join(html.escape(f) for f in n)}</div>'
            body+=GUIDE_HTML
            body+='<div class="foot">文件检索管理系统 · 新增文件放进已选文件夹后点“新增文件一键索引”只补新文件；需全量刷新才点“全部文件重建索引”</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())

if __name__=="__main__":
    ap=argparse.ArgumentParser()
    ap.add_argument("port",nargs="?",type=int,default=8765)
    a=ap.parse_args()
    print("启动中 → 浏览器打开: http://localhost:%d  (Ctrl+C 停止)"%a.port)
    ThreadingHTTPServer(("127.0.0.1",a.port),H).serve_forever()