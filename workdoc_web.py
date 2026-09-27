#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 文件检索管理系统 — Web版（百度式简洁界面 · 一点设计感）
# 功能：正文全文检索（不止文件名）· 多文件夹 · 多关键词并行(AND/OR) · 分类筛选 · 统计信息 · 最近新增
# 启动: python3 workdoc_web.py [端口，默认8765]  浏览器打开 http://localhost:8765
import os, sys, json, sqlite3, html, subprocess, argparse, re
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote_plus, urlparse, parse_qs, quote

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

# ============ 页面样式（百度式简洁 · 仅保留品牌行 · 使用说明为独立按钮页） ============
HTML_TOP="""<!doctype html><html><head><meta charset=utf-8><title>文件检索管理系统</title>
<style>
*{box-sizing:border-box}
body{font-family:"宋体",SimSun,serif,-apple-system,BlinkMacSystemFont;background:linear-gradient(180deg,#f8f4ee,#efe7d8);min-height:100vh;margin:0;padding:0;color:#3a2c1a}
.col{max-width:900px;margin:0 auto;padding:30px 20px 70px}
.brandrow{display:flex;align-items:center;justify-content:center;padding:26px 0 14px}
.brand{font-family:"宋体",SimSun,serif;font-size:45px;font-weight:bold;color:#3a2c1a;letter-spacing:2px;text-align:center}
.backrow{display:flex;align-items:center;gap:10px;padding:14px 18px;background:#fff;border:1px solid #e2d2b0;border-radius:14px;box-shadow:0 3px 8px rgba(90,60,31,.08);margin-bottom:14px}
.backrow a{text-decoration:none;color:#8b5f3f;font-weight:600;padding:7px 12px;border-radius:10px;background:#f3e3c0}
.backrow a:hover{background:#8b5f3f;color:#fff}
.backrow span{font-size:15px;font-weight:600;color:#3a2c1a}
.searchcard{background:#fff;border-radius:20px;border:1px solid #e2d2b0;box-shadow:0 6px 18px rgba(90,60,31,.1);padding:30px 34px}
.searchbar{display:flex;justify-content:center;align-items:center;gap:12px;flex-wrap:wrap}
.searchbox{flex:1 1 100%;max-width:560px;padding:16px 22px;font-size:17px;border:2px solid #d8c29a;border-radius:16px;outline:none;background:#fff;box-shadow:inset 0 2px 6px rgba(90,60,31,.1);transition:.15s}
.searchbox:focus{border-color:#8b5f3f;box-shadow:0 0 0 4px rgba(139,95,63,.15)}
.searchbtns{display:flex;justify-content:center;align-items:center;gap:12px;flex-wrap:wrap;margin-top:14px}
select{padding:11px 14px;border-radius:12px;border:2px solid #d8c29a;background:#fff;font-size:14px;color:#5a3c1f;box-shadow:0 2px 5px rgba(90,60,31,.1)}
button{padding:13px 30px;border-radius:12px;background:#8b5f3f;color:#fff;border:0;cursor:pointer;font-size:15px;box-shadow:0 4px 10px rgba(90,60,31,.25);transition:.15s;letter-spacing:1px}
button:hover{background:#5a3c1f;transform:translateY(-1px)}
button.sec{background:#fff;color:#5a3c1f;border:2px solid #d8c29a;box-shadow:0 3px 8px rgba(90,60,31,.1)}
button.sec:hover{background:#8b5f3f;color:#fff;border-color:#8b5f3f}
.btnrow{display:flex;justify-content:center;gap:14px;flex-wrap:wrap;margin:22px 0 8px}
.cnt{color:#6a5337;font-size:14px;margin:14px 0;text-align:center}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 4px 12px rgba(90,60,31,.1);margin-top:6px}
td{vertical-align:top;padding:13px 16px;border-bottom:1px solid #eee2c6;font-size:14px;color:#4a3a1f}
tr:hover{background:#f8f2e4}
.pagecard{background:#fff;border-radius:14px;padding:18px 20px;border:1px solid #e2d2b0;box-shadow:0 4px 12px rgba(90,60,31,.08);margin:12px 0}
.pagecard .name{font-weight:600;color:#3a2c1a;font-size:18px;padding-bottom:6px;margin-bottom:6px}
.tag{display:inline-block;background:#f3e3c0;color:#5a3c1f;font-size:11px;padding:3px 10px;border-radius:8px;margin-right:6px}
.path{color:#8b7359;font-size:12px;word-break:break-all}
.snp{color:#5a4a2e;font-size:13px;line-height:1.55;border-left:3px solid #d8c29a;padding-left:9px;margin-top:6px}
.openrow{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;padding-top:8px;border-top:1px solid #eee2c6}
.openrow a{text-decoration:none;font-size:12px;padding:6px 12px;border-radius:8px;border:1px solid #c9a97a;background:#fbf3e3;color:#5a3c1f;transition:.15s}
.openrow a:hover{background:#8b5f3f;color:#fff;border-color:#8b5f3f}
.preview{background:#fff;border-radius:14px;padding:16px;border:1px solid #e2d2b0;box-shadow:0 4px 12px rgba(90,60,31,.08);margin:12px 0}
.preview .pre-title{font-weight:600;color:#3a2c1a;font-size:17px;padding-bottom:6px;margin-bottom:6px;border-bottom:1px solid #eee2c6}
.preview .pre-txt{color:#4a3a1f;font-size:13px;line-height:1.6;white-space:pre-wrap;word-break:break-word}
.preview embed,.preview img{width:100%;border-radius:10px}
.shelf{display:flex;flex-wrap:wrap;gap:12px;justify-content:center;margin:18px 0}
.shelfcard{flex:1 1 240px;background:#fff;border-radius:14px;padding:18px;border:1px solid #e2d2b0;box-shadow:0 3px 8px rgba(90,60,31,.1)}
.shelfcard b{color:#5a3c1f}
.foldadd{background:#fff;border-radius:14px;padding:20px;border:1px solid #e2d2b0;box-shadow:0 3px 8px rgba(90,60,31,.1)}
.foldadd input{width:100%;padding:13px 16px;border:2px solid #d8c29a;border-radius:12px;font-size:15px}
.guide{border:2px solid #d8c29a;border-radius:18px;background:#fff;padding:20px 24px;margin:8px 0 10px;box-shadow:0 5px 14px rgba(90,60,31,.1)}
.guide b{color:#5a3c1f;font-size:16px;display:block;margin-bottom:6px;text-align:center}
.guide ul{list-style:none;padding:0;margin:8px 0 0;color:#5a4a2e;font-size:14px;line-height:1.75}
.guide li{margin:5px 0}
.guide-title{font-family:"宋体",SimSun,serif;font-size:20px;font-weight:bold;color:#5a3c1f;text-align:center;margin-bottom:14px;letter-spacing:1px}
.guide-sec{margin:12px 0}
.guide-sec-title{font-weight:bold;color:#5a3c1f;font-size:15px;text-align:left;letter-spacing:.5px}
.guide-sec-text{color:#5a4a2e;font-size:14px;line-height:1.75;text-align:left;margin-top:4px;padding-left:14px;border-left:3px solid #d8c29a}
.foot{color:#8b7359;font-size:13px;margin-top:26px;text-align:center;border-top:1px solid #e0d0b0;padding-top:16px}
</style></head><body><div class="col">
<div class="brandrow"><div class="brand">文件检索管理系统</div></div>"""

GUIDE_HTML=f"""
<div class="guide"><div class="guide-title">📖 使用说明</div>
<div class="guide-sec"><div class="guide-sec-title">检索</div><div class="guide-sec-text">输入关键词，空格分隔多个关键词并行；全部命中选 AND，任一命中选 OR；可加“分类”筛选（PDF/Word/Excel/PPT/文本）。</div></div>
<div class="guide-sec"><div class="guide-sec-title">新增文件一键索引</div><div class="guide-sec-text">把新文件放进已设置的文件来源文件夹，点此按钮只补新文件，几秒钟完成。</div></div>
<div class="guide-sec"><div class="guide-sec-title">全部文件重建索引</div><div class="guide-sec-text">需要全量刷新（更换文件夹、清理重复等）时点此按钮，耗时较长（约20分钟），期间检索会短暂暂缓。</div></div>
<div class="guide-sec"><div class="guide-sec-title">设置文件来源</div><div class="guide-sec-text">添加/移除要索引的文件夹完整路径；可同时设置多个文件夹。</div></div>
<div class="guide-sec"><div class="guide-sec-title">统计信息 / 最近新增</div><div class="guide-sec-text">统计信息查看按文件类型、按所在文件夹的分布；最近新增查看最近收录的文档。</div></div>
<div class="guide-sec"><div class="guide-sec-title">使用注意</div><div class="guide-sec-text">_重复待清理 文件夹自动跳过；索引保存在本程序目录下的 workdoc_index.db；本服务面向 macOS（旧 .doc/.ppt 依赖 textutil）。</div></div>
</div>"""

def snip(text,q):
    i=text.find(q)
    if i<0: return ""
    return text[max(0,i-60):i+120].replace("\n"," ")

def _qs(u):
    return parse_qs(urlparse(unquote_plus(u)).query) if "?" in u else {}

def _base():
    return HTML_TOP

def _back(title):
    return f'<div class="backrow"><a href="/">← 返回首页</a><span>{title}</span></div>'

class H(BaseHTTPRequestHandler):
    def log_message(self,*a): pass
    def do_GET(self):
        u=self.path
        qs=_qs(u)
        if u.startswith("/guide"):
            body=_base()+_back("📖 使用说明")
            body+=GUIDE_HTML
            body+='<div class="btnrow"><a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+='<div class="foot">文件检索管理系统 · 使用说明</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/search"):
            q=unquote_plus(qs.get("q",[""])[0])[:80]
            mode=qs.get("mode",["and"])[0] if qs.get("mode") else "and"
            if mode not in ("and","or"): mode="and"
            ty=qs.get("t",["all"])[0] if qs.get("t") else "all"
            if ty not in TYPE_GROUPS: ty="all"
            cnt,rows=search(q,mode,ty)
            body=_base()+_back("🔍 检索结果")
            mode_sel=f'<select name="mode"><option value="and" {"selected" if mode=="and" else ""}>全部命中(AND)</option><option value="or" {"selected" if mode=="or" else ""}>任一命中(OR)</option></select>'
            t_sel='<select name="t"><option value="all" {"selected" if ty=="all" else ""}>分类：全部</option><option value="pdf" {"selected" if ty=="pdf" else ""}>PDF</option><option value="word" {"selected" if ty=="word" else ""}>Word</option><option value="excel" {"selected" if ty=="excel" else ""}>Excel</option><option value="ppt" {"selected" if ty=="ppt" else ""}>PPT</option><option value="text" {"selected" if ty=="text" else ""}>文本</option></select>'
            body+=f'<div class="searchcard"><form method=get action="/search"><div class="searchbar"><input class="searchbox" name=q value="{html.escape(q)}" placeholder="输入关键词，空格分隔多个关键词并行检索"></div><div class="searchbtns">{mode_sel}{t_sel}<button>检索</button></div></form></div>'
            body+=f'<div class="cnt">命中 <b>{cnt}</b> 份（只显示前100） · 匹配方式：{"全部命中" if mode=="and" else "任一命中"} · 分类：{html.escape(ty)}</div>'
            body+=''.join(f'<div class="pagecard"><span class="tag">{html.escape(ext)}</span><div class="name">{html.escape(name)}</div><div class="path">{html.escape(path)}</div><div class="snp">{html.escape(snip(text,q))}</div><div class="openrow"><a href="/open?path={quote(path)}">🔗 打开文件</a><a href="/preview?path={quote(path)}">👁 预览</a></div></div>' for path,name,ext,text in rows[:100])
            body+='<div class="foot">文件检索管理系统 · 多关键词（空格分隔），全部命中选 AND，任一命中选 OR</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html; charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/stats"):
            tot,byext,fld=stats()
            body=_base()+_back("📊 统计信息")
            body+=f'<div class="cnt">收录文档总数：<b>{tot}</b> 份</div><div class="searchcard"><b>按文件类型</b><table><tr><td>类型</td><td>数量</td></tr>'+''.join(f'<tr><td>{html.escape(e)}</td><td>{c}</td></tr>' for e,c in byext)+'</table></div>'
            body+='<div class="searchcard"><b>按所在文件夹（前20）</b><table><tr><td>文件夹</td><td>数量</td></tr>'+''.join(f'<tr><td>{html.escape(d)}</td><td>{c}</td></tr>' for d,c in fld)+'</table></div>'
            body+='<div class="foot">文件检索管理系统 · <a href="/" style="color:#8b7359">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/recent"):
            rows=recent(30)
            body=_base()+_back("🕘 最近新增")
            body+=f'<div class="cnt">最近收录的 <b>{len(rows)}</b> 份（按收录先后）</div>'
            body+=''.join(f'<div class="pagecard"><span class="tag">{html.escape(ext)}</span><div class="name">{html.escape(name)}</div><div class="path">{html.escape(path)}</div><div class="openrow"><a href="/open?path={quote(path)}">🔗 打开文件</a><a href="/preview?path={quote(path)}">👁 预览</a></div></div>' for path,name,ext in rows)
            body+='<div class="foot">文件检索管理系统 · <a href="/" style="color:#8b7359">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/settings"):
            folders=get_folders()
            body=_base()+_back("📁 设置文件来源")
            body+='<div class="shelf">'+''.join(f'<div class="shelfcard">📁 <b>{html.escape(f)}</b><br><a href="/remove_folder?path={html.escape(f)}" style="color:#8b5f3f">移除</a></div>' for f in folders)+'</div>'
            body+='<div class="foldadd"><form method=get action="/add_folder"><input name="path" placeholder="输入要加入的文件夹完整路径，如 /Users/xxx/文档"><div class="btnrow"><button>添加文件夹</button></div></form></div>'
            body+='<div class="foot">文件检索管理系统 · 添加后点“新增文件一键索引”补索引新文件 · <a href="/" style="color:#8b7359">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/add_folder"):
            p=unquote_plus(qs.get("path",[""])[0]).strip()
            body=_base()+_back("➕ 添加文件夹")
            if p and os.path.isdir(p):
                fs=get_folders()
                if p not in fs: set_folders(fs+[p])
                body+=f'<div class="cnt">已添加文件夹：<b>{html.escape(p)}</b>。点“新增文件一键索引”补索引新文件。</div>'
            else:
                body+=f'<div class="cnt">❌ 路径不可用或不存在：<b>{html.escape(p)}</b>（请填完整路径）</div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">管理文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+='<div class="foot">文件检索管理系统 · <a href="/" style="color:#8b7359">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/remove_folder"):
            p=unquote_plus(qs.get("path",[""])[0]).strip()
            fs=[f for f in get_folders() if f!=p]
            set_folders(fs)
            body=_base()+_back("➖ 移除文件夹")
            body+=f'<div class="cnt">已移除文件夹：<b>{html.escape(p)}</b>（索引中已收录的文件仍可搜到，下次重建时剔除）。</div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">管理文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+='<div class="foot">文件检索管理系统 · <a href="/" style="color:#8b7359">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/index_new"):
            n=index_new()
            body=_base()+_back("📥 新增文件一键索引")
            body+=f'<div class="cnt">新增文件一键索引完成：新增 <b>{n}</b> 份</div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">设置文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+='<div class="foot">文件检索管理系统 · <a href="/" style="color:#8b7359">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/reindex"):
            n=index()
            body=_base()+_back("🔁 全部文件重建索引")
            body+=f'<div class="cnt">全部文件重建索引完成：<b>{n}</b> 份（含全部文件来源）</div>'
            body+='<div class="btnrow"><a href="/settings" style="text-decoration:none"><button class="sec">设置文件来源</button></a> <a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+='<div class="foot">文件检索管理系统 · <a href="/" style="color:#8b7359">← 返回首页</a></div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/open"):
            p=unquote_plus(qs.get("path",[""])[0])
            msg=""
            if os.path.isfile(p):
                subprocess.Popen(["open",p])
                msg=f'已在系统默认应用中打开：<b>{html.escape(p)}</b>'
            else:
                msg=f'❌ 文件路径不存在或不可访问：<b>{html.escape(p)}</b>'
            body=_base()+_back("🔗 打开文件")
            body+=f'<div class="cnt">{msg}</div>'
            body+='<div class="btnrow"><a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+='<div class="foot">文件检索管理系统 · 打开文件</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/preview"):
            p=unquote_plus(qs.get("path",[""])[0])
            conn=_conn(); _ensure_fts(conn)
            row=conn.execute("SELECT name,ext,text FROM docs WHERE path=?",(p,)).fetchone()
            conn.close()
            body=_base()+_back("👁 文件预览")
            if row:
                name,ext,text=row
                body+=f'<div class="preview"><div class="pre-title">{html.escape(name)} <span class="tag">{html.escape(ext)}</span></div>'
                body+=f'<div class="pre-txt">{html.escape(text or "(无可预览文本，请点“打开文件”查看)")}</div>'
                body+=f'<div class="openrow"><a href="/open?path={quote(p)}">🔗 打开文件（默认应用）</a></div></div>'
                body+=f'<div class="path" style="text-align:center;color:#8b7359">{html.escape(p)}</div>'
            else:
                body+=f'<div class="cnt">❌ 未在索引中找到该文件：<b>{html.escape(p)}</b></div>'
            body+='<div class="btnrow"><a href="/" style="text-decoration:none"><button class="sec">返回首页</button></a></div>'
            body+='<div class="foot">文件检索管理系统 · 文件预览</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())
        elif u.startswith("/file"):
            p=unquote_plus(qs.get("path",[""])[0])
            if not os.path.isfile(p):
                self.send_response(404); self.send_header("Content-Type","text/plain"); self.end_headers(); self.wfile.write(b"not found"); return
            ext=os.path.splitext(p)[1].lower()
            mime={"pdf":"application/pdf","png":"image/png","jpg":"image/jpeg","jpeg":"image/jpeg","gif":"image/gif","webp":"image/webp","bmp":"image/bmp","svg":"image/svg+xml","txt":"text/plain","csv":"text/plain","md":"text/plain","json":"text/plain","xml":"text/plain","wps":"text/plain","rtf":"text/plain","html":"text/html"}
            ct=mime.get(ext,"application/octet-stream")
            with open(p,"rb") as f: data=f.read()
            self.send_response(200); self.send_header("Content-Type",ct); self.send_header("Content-Disposition","inline"); self.end_headers(); self.wfile.write(data)
        else:
            n=get_folders()
            body=_base()
            body+=f'<div class="searchcard"><form method=get action="/search"><div class="searchbar"><input class="searchbox" name=q placeholder="输入关键词，空格分隔多个关键词并行检索" value=""></div><div class="searchbtns"><select name="mode"><option value="and" selected>全部命中(AND)</option><option value="or">任一命中(OR)</option></select><select name="t"><option value="all" selected>分类：全部</option><option value="pdf">PDF</option><option value="word">Word</option><option value="excel">Excel</option><option value="ppt">PPT</option><option value="text">文本</option></select><button>检索</button></div></form></div>'
            body+=f'<div class="btnrow"><form method=get action="/index_new"><button class="sec">📥 新增文件一键索引</button></form><form method=get action="/reindex"><button class="sec">🔁 全部文件重建索引</button></form><a href="/settings" style="text-decoration:none"><button class="sec">📁 设置文件来源</button></a><a href="/guide" style="text-decoration:none"><button class="sec">📖 使用说明</button></a></div>'
            body+=f'<div class="cnt">当前文件来源文件夹：{" · ".join(html.escape(f) for f in n)}</div>'
            body+='<div class="foot">文件检索管理系统 · 新增文件放进已选文件夹后点“新增文件一键索引”只补新文件；需全量刷新才点“全部文件重建索引”</div></body></html>'
            self.send_response(200); self.send_header("Content-Type","text/html;charset=utf-8"); self.end_headers(); self.wfile.write(body.encode())

if __name__=="__main__":
    ap=argparse.ArgumentParser()
    ap.add_argument("port",nargs="?",type=int,default=8765)
    a=ap.parse_args()
    print("启动中 → 浏览器打开: http://localhost:%d  (Ctrl+C 停止)"%a.port)
    ThreadingHTTPServer(("127.0.0.1",a.port),H).serve_forever()