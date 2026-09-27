#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 文件检索管理系统 — Web版（MediaWiki 风格的界面与文档管理参考）
# 功能：正文全文检索（不止文件名）· 多文件夹 · 多关键词并行(AND/OR) · 分类筛选 · 统计信息 · 最近新增
# 启动: python3 workdoc_web.py [端口，默认8765]  浏览器打开 http://localhost:8765
import os, sys, json, sqlite3, html, subprocess, argparse, re
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote_plus, urlparse, parse_qs

DB=os.path.join(os.path.dirname(os.path.abspath(__file__)),"workdoc_index.db")
try: sys.path.insert(0,os.path.join(os.path.dirname(os.path.abspath(__file__)),"wpslibs"))
except: pass
DEFAULT_FOLDER="/Users/hh/Documents/工作文档"
TYPE_GROUPS={"all":[], "pdf":["pdf"], "word":["doc","docx"], "excel":["xls","xlsx","et"], "ppt":["ppt","pptx"], "text":["csv","txt","wps","rtf","html","json","xml","md","et"]}

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

def search(q, mode="and", ty="all", top=100):
    ks=[k for k in re.split(r"\s+",q.strip()) if k]
    conn=_conn(); _ensure_fts(conn)
    if not ks:
        conn.close(); return 0,[]
    like=["%"+k+"%" for k in ks]
    tyfilt=""
    if ty and ty in TYPE_GROUPS and TYPE_GROUPS[ty]:
        tyfilt=f" AND ext IN ({','.join('?' for _ in TYPE_GROUPS[ty])})"
    if len(ks)==1 and len(ks[0])>=3:
        fts=[r[0] for r in conn.execute("SELECT rowid FROM docs_fts WHERE docs_fts MATCH ?",(ks[0],))]
        nm=[r[0] for r in conn.execute("SELECT rowid FROM docs WHERE name LIKE ?",(like[0],))]
        cand=list(set(fts+nm))
        if cand:
            ph=",".join("?" for _ in cand)
            cond="(text LIKE ? OR name LIKE ?)"+tyfilt
            params=cand+like+like
            if tyfilt: params+=TYPE_GROUPS[ty]
            cnt=conn.execute(f"SELECT COUNT(*) FROM docs WHERE rowid IN ({ph}) AND {cond}",params).fetchone()[0]
            rows=conn.execute(f"SELECT path,name,ext,text FROM docs WHERE rowid IN ({ph}) AND {cond} LIMIT ?",params+[top]).fetchall()
            conn.close(); return cnt,rows
    clause="(text LIKE ? OR name LIKE ?)"
    sep=" AND " if mode=="and" else " OR "
    cond=sep.join(clause for _ in ks)+tyfilt
    params=sum([[lk,lk] for lk in like],[])
    if tyfilt: params+=TYPE_GROUPS[ty]
    cnt=conn.execute(f"SELECT COUNT(*) FROM docs WHERE {cond}",params).fetchone()[0]
    rows=conn.execute(f"SELECT path,name,ext,text FROM docs WHERE {cond} LIMIT ?",params+[top]).fetchall()
    conn.close()
    return cnt,rows

def stats():
    conn=_conn(); _ensure_fts(conn)
    tot=conn.execute("SELECT COUNT(*) FROM docs").fetchone()[0]
    byext=conn.execute("SELECT ext,COUNT(*) c FROM docs GROUP BY ext ORDER BY c DESC").fetchall()
    rows=conn.execute("SELECT path FROM docs").fetchall()
    conn.close()
    folders={}
    for (p,) in rows:
        d=os.path.dirname(p) or p
        folders[d]=folders.get(d,0)+1
    fld=sorted(folders.items(), key=lambda x:-x[1])[:20]
    return tot,byext,fld

def recent(n=30):
    conn=_conn(); _ensure_fts(conn)
    rows=conn.execute("SELECT path,name,ext FROM docs ORDER BY rowid DESC LIMIT ?",(n,)).fetchall()
    conn.close()
    return rows

# ============ 页面样式（参考 MediaWiki：左侧导航栏 + 顶部标签页 + 分类/统计/最近新增 + 页面卡片式结果） ============
HTML_TOP="""<!doctype html><html><head><meta charset=utf-8><title>文件检索管理系统</title>
<style>
*{box-sizing:border-box}
body{font-family:"宋体",SimSun,serif,-apple-system,BlinkMacSystemFont;background:linear-gradient(160deg,#3a2c1a,#5a4526,#7a6548);min-height:100vh;margin:0;padding:0}
.wrap{max-width:1100px;margin:auto;padding:26px 22px 60px;display:flex;flex-wrap:wrap;gap:20px}
.sidebar{flex:0 0 210px;background:#3a2c1a;border-radius:18px;padding:16px;border:1px solid #c9a97a;position:sticky;top:12px;align-self:flex-start}
.sidebar .brand{font-family:"宋体",SimSun,serif;font-size:20px;font-weight:bold;color:#f7e9d0;text-align:center;margin-bottom:14px;letter-spacing:1px}
.sidebar a{display:block;color:#d9c6a0;padding:9px 12px;border-radius:10px;text-decoration:none;transition:.15s;font-size:14px}
.sidebar a:hover{background:#5a3c1f;color:#f7e9d0}
.sidebar a.active{background:#8b5f3f;color:#fff;font-weight:600}
.main{flex:1 1 auto;min-width:0}
.header{text-align:center;background:linear-gradient(135deg,#3a2c1a,#5a4526,#7a6548);border-radius:18px;padding:28px 20px 24px;border:1px solid #c9a97a;box-shadow:0 6px 18px rgba(0,0,0,.35)}
.brand{font-family:"宋体",SimSun,serif;font-size:45px;font-weight:bold;color:#f7e9d0;text-align:center;letter-spacing:2px;text-shadow:0 3px 6px rgba(0,0,0,.4)}
.sub{font-size:13px;color:#d9c6a0;margin-top:6px;text-align:center}
.tabbar{display:flex;gap:4px;flex-wrap:wrap;margin-top:14px;background:#fbf3e3;border-radius:14px;border:1px solid #d8c29a;padding:8px 12px;box-shadow:0 3px 8px rgba(0,0,0,.15)}
.tabbar a{flex:1 1 auto;text-align:center;padding:9px 14px;border-radius:10px;text-decoration:none;color:#5a3c1f;font-size:14px}
.tabbar a:hover{background:#f3e3c0}
.tabbar a.active{background:#8b5f3f;color:#fff;font-weight:600}
.backrow{display:flex;align-items:center;margin-top:12px;background:#fbf3e3;border-radius:14px;border:1px solid #d8c29a;padding:10px 16px;box-shadow:0 3px 8px rgba(0,0,0,.15)}
.backrow a.home{text-decoration:none;color:#5a3c1f;font-weight:600;border:1px solid #c9a97a;padding:7px 12px;border-radius:10px;background:#f7e9d0;transition:.15s}
.backrow a.home:hover{background:#5a3c1f;color:#f7e9d0}
.backrow .brand2{font-size:16px;font-weight:700;color:#5a3c1f}
h2{margin:20px 0 10px;font-size:24px;color:#f3e6c8;letter-spacing:.5px;text-shadow:0 2px 4px rgba(0,0,0,.35)}
.searchcard{background:#fbf3e3;border-radius:16px;padding:16px 20px;border:1px solid #d8c29a;box-shadow:0 4px 12px rgba(0,0,0,.15);margin-top:16px}
.searchbar{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
input{flex:1 1 300px;padding:13px 16px;font-size:15px;border:2px solid #c9a97a;border-radius:12px;outline:none;background:#fbf3e3;box-shadow:0 2px 6px rgba(0,0,0,.1)}
input:focus{border-color:#8b5f3f;box-shadow:0 0 0 3px rgba(139,95,63,.25)}
.searchbtns{display:flex;gap:10px;align-items:center;margin-top:12px;flex-wrap:wrap}
select{padding:11px 12px;border-radius:11px;border:2px solid #c9a97a;background:#fbf3e3;font-size:14px;color:#5a3c1f}
button{padding:13px 22px;border-radius:11px;background:#8b5f3f;color:#fbf3e3;border:0;cursor:pointer;font-size:15px;box-shadow:0 3px 8px rgba(0,0,0,.25);transition:.15s}
button:hover{background:#5a3c1f;transform:translateY(-1px)}
button.sec{background:#b08857} button.sec:hover{background:#8b5f3f}
.btnrow{display:flex;gap:10px;flex-wrap:wrap;margin:16px 0 8px}
.cnt{color:#e6d3ae;font-size:14px;margin:10px 0 12px}
table{width:100%;border-collapse:collapse;margin-top:8px;background:#fbf3e3;border-radius:14px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,.15)}
td{vertical-align:top;padding:14px 16px;border-bottom:1px solid #e2d2b0;font-size:14px;color:#4a3a1f}
tr:hover{background:#f3e3c0}
.pagecard{background:#fbf3e3;border-radius:14px;padding:16px;border:1px solid #d8c29a;box-shadow:0 4px 12px rgba(0,0,0,.15);margin-top:12px}
.pagecard .name{font-weight:600;color:#3a2c1a;font-size:18px;border-bottom:1px solid #e2d2b0;padding-bottom:6px;margin-bottom:6px}
.tag{display:inline-block;background:#e8d6b8;color:#5a3c1f;font-size:11px;padding:3px 10px;border-radius:8px;margin-right:6px}
.path{color:#8b7359;font-size:12px;word-break:break-all}
.snp{color:#5a4a2e;font-size:13px;line-height:1.55;border-left:3px solid #c9a97a;padding-left:9px;margin-top:6px}
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
<div class="sidebar"><div class="brand">📚 文件检索管理系统</div><a href="/" {A1}>🔍 检索</a><a href="/settings" {A2}>📁 文件来源</a><a href="/index_new" {A3}>📥 新增一键索引</a><a href="/reindex" {A4}>🔁 全部重建索引</a><a href="/stats" {A5}>📊 统计信息</a><a href="/recent" {A6}>🕘 最近新增</a><a href="/" {A7}>📖 使用说明</a></div>
<div class="main"><div class="header"><div class="brand">📚 文件检索管理系统</div><div class="sub">正文全文检索 · 不止文件名 · 多关键词并行 · 分类筛选</div></div>"""

GUIDE_HTML=f"""
<div class="guide"><b>📖 使用说明</b><ul>
<li><b>检索</b>：输入关键词（空格分隔多个关键词并行），全部命中选 AND，任一命中选 OR；可加“分类”筛选（PDF/Word/Excel/PPT/文本）。</li>
<li><b>新增文件一键索引</b>：把新文件放进已设置的文件来源文件夹，点此按钮只补新文件，几秒钟完成。</li>
<li><b>全部文件重建索引</b>：需要全量刷新（更换文件夹、清理重复等）时点此按钮，耗时较长（约20分钟），期间检索会短暂暂缓。</li>
<li><b>设置文件来源</b>：添加/移除要索引的文件夹完整路径；可同时设置多个文件夹。</li>
<li><b>统计信息</b>：查看按文件类型、按所在文件夹的分布；<b>最近新增</b>查看最近收录的文档。</li>
<li><b>使用注意</b>：_重复待清理 文件夹自动跳过；索引保存在本程序目录下的 workdoc_index.db；本服务面向 macOS（旧 .doc/.ppt 依赖 textutil）。</li>
</ul></div>"""

def snip(text,q):
    i=text.find(q)
    if i<0: return ""
    return text[max(0,i-60):i+120].replace("\n"," ")

def _qs(u):
    return parse_qs(urlparse(unquote_plus(u)).query) if "?" in u else {}

def _active(route):
    return 'class="active"' if route else ""

def _tabbar(active):
    return f'<div class="tabbar"><a href="/" {"class=active" if active=="home" else ""}>首页</a><a href="/search" {"class=active" if active=="search" else ""}>检索</a><a href="/settings" {"class=active" if active=="settings" else ""}>文件来源</a><a href="/stats" {"class=active" if active=="stats" else ""}>统计</a><a href="/recent" {"class=active" if active=="recent" else ""}>最近新增</a></div>'

class H(BaseHTTPRequestHandler):
    def log_message(self,*a): pass
    def do_GET(self):
        u=self.path
        qs=_qs(u)
        if u.startswith("/search"):
            q=unquote_plus(qs.get("q",[""])[0])[:80]
            mode=qs.get("mode",["and"])[0] if qs.get("mode") else "and"
            if mode not in ("and","or"): mode="and"
            ty=qs.get("t",["all"])[0] if qs.get("t") else "all"
            if ty not in TYPE_GROUPS: ty="all"
            cnt,rows=search(q,mode,ty)
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("search")
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">🔍 检索结果</span></div>'
            body+=f'<div class="cnt">命中 <b>{cnt}</b> 份（只显示前100） · 匹配方式：{"全部命中" if mode=="and" else "任一命中"} · 分类：{html.escape(ty)}</div>'
            body+=''.join(f'<div class="pagecard"><span class="tag">{html.escape(ext)}</span><div class="name">{html.escape(name)}</div><div class="path">{html.escape(path)}</div><div class="snp">{html.escape(snip(text,q))}</div></div>' for path,name,ext,text in rows[:100])
            body+='<div class="foot">文件检索管理系统 · 多关键词（空格分隔），全部命中选 AND，任一命中选 OR</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html; charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/stats"):
            tot,byext,fld=stats()
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("stats")
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">📊 统计信息</span></div>'
            body+=f'<div class="cnt">收录文档总数：<b>{tot}</b> 份</div><div class="searchcard"><b>按文件类型</b><table><tr><td>类型</td><td>数量</td></tr>'+''.join(f'<tr><td>{html.escape(e)}</td><td>{c}</td></tr>' for e,c in byext)+'</table></div>'
            body+='<div class="searchcard"><b>按所在文件夹（前20）</b><table><tr><td>文件夹</td><td>数量</td></tr>'+''.join(f'<tr><td>{html.escape(d)}</td><td>{c}</td></tr>' for d,c in fld)+'</table></div>'
            body+='<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/recent"):
            rows=recent(30)
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("recent")
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">🕘 最近新增</span></div>'
            body+=f'<div class="cnt">最近收录的 <b>{len(rows)}</b> 份（按收录先后）</div>'
            body+=''.join(f'<div class="pagecard"><span class="tag">{html.escape(ext)}</span><div class="name">{html.escape(name)}</div><div class="path">{html.escape(path)}</div></div>' for path,name,ext in rows)
            body+='<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/settings"):
            folders=get_folders()
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("settings")
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">📁 设置文件来源</span></div>'
            body+='<div class="shelf">'+''.join(f'<div class="shelfcard">📁 <b>{html.escape(f)}</b><br><a href="/remove_folder?path={html.escape(f)}" style="color:#8b5f3f">移除</a></div>' for f in folders)+'</div>'
            body+='<div class="foldadd"><form method=get action="/add_folder"><input name="path" placeholder="输入要加入的文件夹完整路径，如 /Users/xxx/文档"><button>添加文件夹</button></form></div>'
            body+=GUIDE_HTML
            body+='<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a> · 添加后点“新增文件一键索引”补索引新文件</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/add_folder"):
            p=unquote_plus(qs.get("path",[""])[0]).strip()
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("settings")
            if p and os.path.isdir(p):
                fs=get_folders()
                if p not in fs: set_folders(fs+[p])
                body+=f'<div class="cnt">已添加文件夹：<b>{html.escape(p)}</b>。点“新增文件一键索引”补索引新文件。</div>'
            else:
                body+=f'<div class="cnt">❌ 路径不可用或不存在：<b>{html.escape(p)}</b>（请填完整路径）</div>'
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">➕ 添加文件夹</span></div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">管理文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/remove_folder"):
            p=unquote_plus(qs.get("path",[""])[0]).strip()
            fs=[f for f in get_folders() if f!=p]
            set_folders(fs)
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("settings")
            body+=f'<div class="cnt">已移除文件夹：<b>{html.escape(p)}</b>（索引中已收录的文件仍可搜到，下次重建时剔除）。</div>'
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">➖ 移除文件夹</span></div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">管理文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/index_new"):
            n=index_new()
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("index_new")
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">📥 新增文件一键索引</span></div>'
            body+=f'<div class="cnt">新增文件一键索引完成：新增 <b>{n}</b> 份</div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">设置文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/reindex"):
            n=index()
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("reindex")
            body+='<div class="backrow"><a class="home" href="/">← 返回首页</a><span class="brand2">🔁 全部文件重建索引</span></div>'
            body+=f'<div class="cnt">全部文件重建索引完成：<b>{n}</b> 份（含全部文件来源）</div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">设置文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+=GUIDE_HTML+'<div class="foot"><a href="/" style="color:#d9c6a0">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        else:
            n=get_folders()
            body=HTML_TOP.replace("{A1}",_active("home")).replace("{A2}",_active("settings")).replace("{A3}",_active("index_new")).replace("{A4}",_active("reindex")).replace("{A5}",_active("stats")).replace("{A6}",_active("recent")).replace("{A7}",_active("guide"))
            body+=_tabbar("home")
            body+=f'<div class="searchcard"><form method=get action="/search"><div class="searchbar"><input name=q placeholder="输入关键词，空格分隔多个关键词并行检索" value=""></div><div class="searchbtns"><select name="mode"><option value="and" selected>全部命中(AND)</option><option value="or">任一命中(OR)</option></select><select name="t"><option value="all" selected>分类：全部</option><option value="pdf">PDF</option><option value="word">Word</option><option value="excel">Excel</option><option value="ppt">PPT</option><option value="text">文本</option></select><button>检索</button></div></form></div>'
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