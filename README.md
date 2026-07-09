# AICU 评论查询 & 数据分析工具

基于 [aicu.cc](https://www.aicu.cc) 的 B站用户数据爬虫，支持爬取评论、弹幕、粉丝牌并输出为 Excel 多表文件，附带用户画像分析。

## 文件说明

| 文件 | 用途 |
|------|------|
| `aicu_gui.py` | **图形化主程序**（推荐使用）。提供 tkinter 界面，输入 UID 即可一键爬取，支持选择输出目录和数据类型勾选 |
| `aicu_crawler.py` | **命令行爬虫**。支持 `--type reply\|videodm\|livedm\|medal\|all` 选择数据类型，自动翻页+视频标题缓存 |
| `build_excel.py` | Excel 生成器（通用版）。从已有 JSON/CSV 生成多表 Excel，含关键信息推断表 |
| `build_excel_93403197.py` | UID 93403197 专用分析脚本 |
| `build_excel_120134752.py` | UID 120134752 专用分析脚本 |
| `build_excel_3546630485707666.py` | UID 3546630485707666 专用分析脚本 |
| `requirements.txt` | Python 依赖清单 (`requests`, `openpyxl`, `curl_cffi`) |

## 快速开始

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. 启动 GUI（推荐）
```bash
python aicu_gui.py
```
输入 B站 UID → 选择输出目录 → 点 "开始爬取"。

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
| 评论数据 | 视频评论（含标题+跳转链接+IP属地占位） | `api.aicu.cc/api/v3/search/getreply` |
| 视频弹幕 | 在视频中发送的弹幕（含出现时间点） | `api.aicu.cc/api/v3/search/getvideodm` |
| 直播弹幕 | 在直播间发送的弹幕（含主播/房间信息） | `api.aicu.cc/api/v3/search/getlivedm` |
| 粉丝牌 | 持有的粉丝牌列表（含等级+主播空间链接） | `api.aicu.cc/api/v3/user/getmedal` |

## 数据来源说明

- **评论/弹幕/粉丝牌**：来自 [aicu.cc](https://www.aicu.cc)，数据非实时，有更新延迟
- **视频标题**：来自 B站官方 API (`api.bilibili.com/x/web-interface/view`)
- **评论跳转链接**：`https://www.bilibili.com/video/av{oid}#reply{rpid}`
- **IP 属地**：B站评论 IP 属地需要登录态 API，aicu.cc 暂不提供此字段
- **用户动态**：aicu.cc 暂不支持动态内容查询

## 依赖

- Python 3.10+
- `requests` — HTTP 请求
- `openpyxl` — Excel 读写
- `curl_cffi`（可选）— 绕过 Cloudflare 封锁，本地运行推荐安装
