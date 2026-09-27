# 文件检索管理系统

正文全文检索系统（不止文件名），基于 Python 内置 HTTP 服务 + SQLite FTS5 全文索引。可复制给其他人使用，面向 macOS。

## 功能
- **正文全文检索**：搜索词不只看文件名，也检索文件正文内容。
- **多关键词并行**：空格分隔多个关键词，可选“全部命中(AND)”或“任一命中(OR)”。
- **分类筛选**：检索时可按 PDF / Word / Excel / PPT / 文本 分类过滤（参考 MediaWiki 的分类功能）。
- **设置文件来源**：用户可添加/移除要索引的文件夹完整路径，可同时设置多个。
- **新增文件一键索引**：只扫描尚未进索引的新文件，几秒钟完成（不触发全量重建）。
- **全部文件重建索引**：需要全量刷新（更换文件夹、清理重复）时使用，耗时约20分钟。
- **统计信息**：按文件类型、按所在文件夹查看收录分布（MediaWiki 式统计页）。
- **最近新增**：按收录先后查看最近进入索引的文档（MediaWiki 式最近更改页）。

## 一键启动
```bash
cd 文件检索管理系统
./start.sh
```
浏览器打开：**http://localhost:8765**

停止：`./stop.sh`

## 开机启动（macOS）
把 `launchd/FileRetrieval.plist` 复制到 `~/Library/LaunchAgents/`，然后：
```bash
launchctl load ~/Library/LaunchAgents/FileRetrieval.plist
```
服务将在登录后自动运行在 http://localhost:8765。

## 使用说明（详见网页最下方“使用说明”框）
1. **检索**：输入关键词（空格分隔多个关键词），选 AND=全部命中 / OR=任一命中。
2. **新增文件**：把新文件放进已设置的文件来源文件夹，点“新增文件一键索引”只补新文件。
3. **全部重建**：更换文件夹或清理重复后，点“全部文件重建索引”（耗时约20分钟，期间检索短暂暂缓）。
4. **设置文件来源**：添加/移除文件夹完整路径，可同时设置多个。

## 文件结构
```
文件检索管理系统/
├── workdoc_web.py      # 主程序
├── start.sh            # 启动脚本
├── stop.sh             # 停止脚本
├── workdoc_index.db    # 索引数据库（首次启动自动创建）
├── wpslibs/            # 解析库（pdf/docx/pptx/xlsx/xls/doc 等）
└── launchd/            # 开机启动配置（macOS）
```

## 依赖说明
- 系统自带 `python3`（macOS 自带 3.9）。
- 老格式 `.doc` / `.ppt` 依赖 macOS 自带 `textutil`。
- 首次使用：在“设置文件来源”中添加你自己的文件夹路径，然后点“新增文件一键索引”（或“全部文件重建索引”）。

## 安全提示
- 服务只监听 `127.0.0.1`（本机），不对外开放，请勿随意改动端口。
- 索引数据库保存在程序目录下，需要备份时复制 `workdoc_index.db` 即可。

## GitHub 上传
本目录即为可上传的仓库；`git init` 后提交即可（README 已备好，.gitignore 已忽略重复/临时文件）。

## 部署说明（如何部署、访问地址）
拿到本程序（克隆自 GitHub 或直接复制目录）后，在 macOS 上按以下步骤部署：

### 1. 获取代码
```bash
git clone <仓库地址> 文件检索管理系统
cd 文件检索管理系统
```

### 2. 一键启动
```bash
./start.sh
```
浏览器打开访问地址：**http://localhost:8765**
启动后即可在首页检索、设置文件来源、点“新增文件一键索引”等。

### 3. 开机启动（登录后自动运行）
把 `launchd/FileRetrieval.plist` 复制到 `~/Library/LaunchAgents/`：
```bash
cp launchd/FileRetrieval.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/FileRetrieval.plist
```
重启或登录后，服务自动运行在 http://localhost:8765。
（若想取消开机启动：`launchctl unload ~/Library/LaunchAgents/FileRetrieval.plist`）

### 4. 首次使用流程
1. 打开 http://localhost:8765；
2. 点“📁 设置文件来源”，添加你自己的文档文件夹完整路径；
3. 点“📥 新增文件一键索引”补索引（或点“🔁 全部文件重建索引”全量）；
4. 回到首页输入关键词检索。

### 5. 停止服务
```bash
./stop.sh
```

### 6. 备份
索引数据保存在 `workdoc_index.db`，备份/迁移时复制该文件即可（wpslibs/ 目录为解析库，一并带上）。

### 7. 环境依赖
- 系统自带 `python3`（macOS 自带 3.9）；
- 老格式 `.doc` / `.ppt` 依赖 macOS 自带 `textutil`；
- 首次运行会自动创建 `workdoc_index.db`（首次请先“设置文件来源”再索引）。

### 8. 安全提示
- 服务只监听 `127.0.0.1`（本机），不对外开放；
- 端口默认 8765，可在 `./start.sh 端口号` 自定义（如 `./start.sh 9000`，访问 http://localhost:9000）。