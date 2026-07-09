# AICU 评论查询 & 数据分析工具

基于 [aicu.cc](https://www.aicu.cc) 的 B站用户数据爬虫，支持爬取评论、弹幕、粉丝牌并输出为 Excel 多表文件，附带用户画像分析。

## 文件说明

| 文件 | 用途 |
|------|------|
| `aicu_gui.py` | **图形化主程序**（推荐使用）。tkinter 界面，输入 UID 一键爬取，支持输出目录选择、数据类型勾选、B站登录获取IP属地 |
| `aicu_crawler.py` | **命令行爬虫**。支持 `--type reply\|videodm\|livedm\|medal\|all` 选择数据类型，自动翻页+视频标题缓存 |
| `build_excel.py` | Excel 生成器（通用版）。从 JSON/CSV 生成多表 Excel，含关键信息推断表 |
| `build_excel_93403197.py` | UID 93403197 专用分析脚本 |
| `build_excel_120134752.py` | UID 120134752 专用分析脚本 |
| `build_excel_3546630485707666.py` | UID 3546630485707666 专用分析脚本 |
| `requirements.txt` | Python 依赖清单 (`requests`, `openpyxl`, `curl_cffi`) |
| `分析表格以及锐评/` | **分析模板文件夹**。包含评论数据分析所需的 Skill、提示词和代码模板 |

### `分析表格以及锐评/` 文件夹

| 文件 | 用途 |
|------|------|
| `SKILL_个人画像分析_锐评版.md` | **Skill 定义文件**。触发热词（"分析评论"/"锐评"/"画像"）后自动加载，指导 AI 从 B站评论中推断用户个人信息并以毒舌风格写报告 |
| `提示词模板_个人画像分析_锐评版.md` | **提示词模板**。包含分析的维度框架（性别/年龄/地点/兴趣/性格/消费等）和锐评文风指南 |
| `代码模板_analyze_profile.py` | **分析代码模板**。自动化生成含证据表格的 Excel，将推断与评论行号关联，每条推断可溯源到原始评论 |

## 快速开始

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. 启动 GUI（推荐）

**方式一：终端命令**（最稳妥）
在项目目录 `E:\aicu评论-ds-wb\` 下，打开终端（在文件夹地址栏输入 `cmd` 回车），执行：

```bash
python aicu_gui.py
```

**方式二：双击 bat 批处理**（最方便）
项目根目录已有一个 `启动GUI.bat`，直接双击即可。如果没有，新建一个文本文件，写入以下内容后改名为 `启动GUI.bat`：

```bat
@echo off
cd /d "%~dp0"
call venv\Scripts\python.exe aicu_gui.py
pause
```

**方式三：如果不小心用 VS Code 打开了 `.py` 文件**
右键 `aicu_gui.py` → 选择「打开方式」→ 勾选「Python」并设为默认。或者在 VS Code 中打开后按 `Ctrl+F5` 直接运行。

### 3. 命令行方式
```bash
# 爬取全部数据
python aicu_crawler.py 1703916229 --output-dir output/

# 只爬评论
python aicu_crawler.py 1703916229 --type reply

# 只爬粉丝牌
python aicu_crawler.py 1703916229 --type medal
```

## Excel 输出

生成的 `aicu_data_{uid}.xlsx` 包含多个工作表：

| Sheet | 内容 | 数据来源 |
|-------|------|----------|
| 评论数据 | 视频评论（含标题+跳转链接+IP属地） | `api.aicu.cc/api/v3/search/getreply` + B站IP API(需登录) |
| 视频弹幕 | 在视频中发送的弹幕（含出现时间点） | `api.aicu.cc/api/v3/search/getvideodm` |
| 直播弹幕 | 在直播间发送的弹幕（含主播/房间信息） | `api.aicu.cc/api/v3/search/getlivedm` |
| 粉丝牌 | 持有的粉丝牌列表（含等级+主播空间链接） | `api.aicu.cc/api/v3/user/getmedal` |

## B站登录（获取IP属地）与安全说明

### 登录方式

**方式一：自动读取浏览器 Cookie**（推荐）
1. 点击「🔑 登录B站（浏览器）」→ 在浏览器中登录 B站
2. 登录完成后，点击「📋 自动读取Cookie」→ 程序自动从 Chrome/Edge/Firefox 读取 `SESSDATA`
3. Cookie 自动保存到 `secrets/sessdata.txt`，下次启动自动加载

**方式二：手动粘贴**
1. 浏览器登录 B站 后，按 `F12` → `Application` → `Cookies` → 复制 `SESSDATA` 的值
2. 粘贴到输入框，点击「验证」
3. 验证通过后自动保存到 `secrets/sessdata.txt`

**方式三：手动管理**
- 点击「🔒 管理密钥」→ 直接打开 `secrets/` 文件夹，可手动编辑或删除 `sessdata.txt`

### 安全说明

⚠ **Cookie 安全须知**：
- `SESSDATA` 相当于你的 B站登录凭证，**拥有它可以操作你的账号**
- 该 Cookie **仅保存在你本机的 `secrets/sessdata.txt` 文件中**
- `secrets/` 文件夹已在 `.gitignore` 中排除，**不会被上传到 GitHub**
- 本软件不会将你的 Cookie 发送给除 B站官方 API 以外的任何第三方
- 如果不放心，点击「🔒 管理密钥」可随时删除已保存的凭据
- Cookie 有过期时间（通常几个月），过期后需重新登录

## 数据来源说明

- **评论/弹幕/粉丝牌**：来自 [aicu.cc](https://www.aicu.cc)，数据非实时，有更新延迟
- **视频标题**：来自 B站官方 API (`api.bilibili.com/x/web-interface/view`)
- **评论跳转链接**：`https://www.bilibili.com/video/av{oid}#reply{rpid}`
- **IP 属地**：需登录 B站 后通过 `api.bilibili.com/x/v2/reply/detail` 获取
- **用户动态**：aicu.cc 暂不支持动态内容查询

## 依赖

- Python 3.10+
- `requests` — HTTP 请求
- `openpyxl` — Excel 读写
- `browser-cookie3`（可选）— 自动读取浏览器 B站 Cookie，免手动粘贴
- `pycryptodome`（可选）— 解密 Chrome/Edge Cookie 数据库

---

## 项目位置

本项目的 GitHub 仓库地址：
```
https://github.com/cloud9-sqzsq/aicu--convert-to-excel
```

克隆到本地：
```bash
git clone https://github.com/cloud9-sqzsq/aicu--convert-to-excel.git
cd aicu--convert-to-excel
pip install -r requirements.txt
python aicu_gui.py
```

本地路径（如已克隆）：
```
E:\aicu评论-ds-wb\
```

## AI 使用指南

### 如果你是 AI 助手，如何找到并使用本项目：

1. **找到项目**：项目位于 `E:\aicu评论-ds-wb\`，如果没有这个目录，执行上面的 git clone 命令
2. **确认环境**：项目使用 venv 虚拟环境，Python 3.10+，依赖见 `requirements.txt`
3. **主入口**：`aicu_gui.py` 是图形界面，`aicu_crawler.py` 是命令行版本
4. **爬取流程**：
   - 用户提供 B站 UID（纯数字）
   - 调用 `aicu_crawler.py <uid>` 或启动 GUI 手动输入
   - 输出 Excel 保存在 `output/` 目录（默认）
5. **分析流程**：
   - 爬取完成后，运行对应的 `build_excel_{uid}.py` 生成含分析报告的多表 Excel
   - 文件包含：评论数据表 + 关键信息推断表（含证据行号）
6. **常用命令**：
   ```bash
   # 爬取
   python aicu_crawler.py <uid> --output-dir output/<uid>/ --type reply
   
   # 生成分析Excel（需先有对应的 build_excel_{uid}.py）
   python build_excel.py  # 通用版，编辑 UID 变量后运行
   ```
