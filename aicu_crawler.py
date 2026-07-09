#!/usr/bin/env python3
"""
AICU.cc B站用户数据爬虫（完整版）
================================
根据 B站 UID 查询用户的所有历史数据，包括：
  1. 评论 (getreply)      - 视频评论区发言
  2. 粉丝牌 (getmedal)    - 持有的直播粉丝牌
  3. 视频弹幕 (getvideodm) - 视频中发送的弹幕
  4. 直播弹幕 (getlivedm)  - 直播间发送的弹幕

同时调用 B站 API 获取视频标题。

API 端点:
  - api.aicu.cc/api/v3/search/getreply
  - api.aicu.cc/api/v3/user/getmedal
  - api.aicu.cc/api/v3/search/getvideodm
  - api.aicu.cc/api/v3/search/getlivedm
  - api.bilibili.com/x/web-interface/view

用法:
    python aicu_crawler.py <uid>              # 查询全部 4 类数据
    python aicu_crawler.py <uid> --type reply # 仅查评论
    python aicu_crawler.py <uid> --type medal # 仅查粉丝牌
    python aicu_crawler.py <uid> --type videodm  # 仅查视频弹幕
    python aicu_crawler.py <uid> --type livedm   # 仅查直播弹幕
"""

import requests
try:
    from curl_cffi import requests as cf_requests
    HAS_CURL_CFFI = True
except ImportError:
    HAS_CURL_CFFI = False
import csv
import json
import time
import sys
import os
import argparse
from datetime import datetime, timezone, timedelta

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
AICU_BASE = "https://api.aicu.cc/api/v3"
BILIBILI_VIDEO_INFO_API = "https://api.bilibili.com/x/web-interface/view"
BILIBILI_VIDEO_URL = "https://www.bilibili.com/video/av{oid}"

DEFAULT_PAGE_SIZE = 500
REQUEST_DELAY = 0.6
BILI_DELAY = 0.3
REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
RETRY_BACKOFF = 2

BEIJING_TZ = timezone(timedelta(hours=8))

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0.0.0 Safari/537.36"
)

TITLE_CACHE_NAME = "aicu_video_titles_cache.json"

# 数据类型配置
DATA_TYPES = {
    "reply": {
        "name": "评论",
        "endpoint": "/search/getreply",
        "params": lambda uid, pn, ps: {"uid": uid, "pn": pn, "ps": ps, "mode": 0},
        "list_key": "replies",
        "total_key": ("data", "cursor", "all_count"),
        "is_end_key": ("data", "cursor", "is_end"),
        "csv_prefix": "aicu_comments",
        "sheet_name": "评论数据",
    },
    "videodm": {
        "name": "视频弹幕",
        "endpoint": "/search/getvideodm",
        "params": lambda uid, pn, ps: {"uid": uid, "pn": pn, "ps": ps},
        "list_key": "videodmlist",
        "total_key": ("data", "cursor", "all_count"),
        "is_end_key": ("data", "cursor", "is_end"),
        "csv_prefix": "aicu_videodm",
        "sheet_name": "视频弹幕",
    },
    "livedm": {
        "name": "直播弹幕",
        "endpoint": "/search/getlivedm",
        "params": lambda uid, pn, ps: {"uid": uid, "pn": pn, "ps": ps},
        "list_key": None,  # 特殊结构：分组弹幕
        "total_key": ("data", "cursor", "all_count"),
        "is_end_key": ("data", "cursor", "is_end"),
        "csv_prefix": "aicu_livedm",
        "sheet_name": "直播弹幕",
    },
    "medal": {
        "name": "粉丝牌",
        "endpoint": "/user/getmedal",
        "params": lambda uid, pn, ps: {"uid": uid},
        "list_key": None,
        "csv_prefix": "aicu_medals",
        "sheet_name": "粉丝牌",
    },
}


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def ts_to_beijing(ts) -> str:
    if not ts: return ""
    if ts > 1e12: ts = ts // 1000  # 毫秒转秒
    return datetime.fromtimestamp(ts, tz=BEIJING_TZ).strftime("%Y-%m-%d %H:%M:%S")


def oid_to_url(oid):
    return BILIBILI_VIDEO_URL.format(oid=oid) if oid else ""


def comment_jump_url(oid, rpid):
    return f"https://www.bilibili.com/video/av{oid}#reply{rpid}" if oid and rpid else ""


def rank_label(rank: int) -> str:
    return "一级评论" if rank == 1 else "回复"


def millis_to_time_str(ms):
    """将毫秒进度转为时:分:秒"""
    if not ms: return ""
    s = int(ms) // 1000
    h, m, s = s // 3600, (s % 3600) // 60, s % 60
    return f"{h}:{m:02d}:{s:02d}"


# ---------------------------------------------------------------------------
# 通用 API 请求
# ---------------------------------------------------------------------------
def api_get(endpoint: str, params: dict) -> dict:
    """请求 aicu.cc API。优先使用 curl_cffi 模拟浏览器指纹绕过 Cloudflare 检测。"""
    url = f"{AICU_BASE}{endpoint}"
    headers = {"User-Agent": UA, "Referer": "https://www.aicu.cc/"}

    for attempt in range(MAX_RETRIES):
        try:
            if HAS_CURL_CFFI:
                resp = cf_requests.get(url, params=params, headers=headers,
                                       timeout=REQUEST_TIMEOUT, impersonate="chrome124")
            else:
                resp = requests.get(url, params=params, headers=headers,
                                    timeout=REQUEST_TIMEOUT)
            if resp.status_code in (403, 468):
                print(f"[WARN] aicu.cc 拦截 (HTTP {resp.status_code})，等待重试...")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_BACKOFF ** (attempt + 1))
                    continue
            resp.raise_for_status()
            data = resp.json()
            if data.get("code") != 0:
                print(f"[WARN] API code={data.get('code')}: {data.get('message','')}")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_BACKOFF ** attempt)
                    continue
            return data
        except Exception as e:
            print(f"[WARN] 请求失败 (尝试 {attempt+1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_BACKOFF ** attempt)
    return {"code": -1}


# ---------------------------------------------------------------------------
# 通用翻页爬取
# ---------------------------------------------------------------------------
def crawl_paginated(uid: str, dtype: str, page_size: int) -> list:
    """通用的翻页爬取：适用于 reply 和 videodm"""
    cfg = DATA_TYPES[dtype]
    print(f"\n[INFO] 开始查询 {cfg['name']}...")
    all_items = []
    page = 1

    first = api_get(cfg["endpoint"], cfg["params"](uid, 1, page_size))
    if first.get("code") != 0:
        print(f"[ERROR] 无法获取{cfg['name']}数据")
        return []

    # 获取总数
    total = first
    for k in cfg["total_key"]:
        total = total.get(k, 0) if isinstance(total, dict) else 0
    if not isinstance(total, int): total = 0

    items = first.get("data", {}).get(cfg["list_key"], [])
    all_items.extend(items)
    total_pages = (total - 1) // page_size + 1 if total > 0 else 1
    print(f"[INFO] 总{cfg['name']}数: {total}  总页数: {total_pages}")

    is_end = first
    for k in cfg["is_end_key"]:
        is_end = is_end.get(k, True) if isinstance(is_end, dict) else True

    time.sleep(REQUEST_DELAY)

    while not is_end and len(all_items) < total:
        page += 1
        data = api_get(cfg["endpoint"], cfg["params"](uid, page, page_size))
        if data.get("code") != 0:
            continue

        is_end = data
        for k in cfg["is_end_key"]:
            is_end = is_end.get(k, True) if isinstance(is_end, dict) else True

        page_items = data.get("data", {}).get(cfg["list_key"], [])
        all_items.extend(page_items)
        print(f"  第 {page}/{total_pages} 页 -> {len(page_items)} 条 | 累计: {len(all_items)}/{total}")
        time.sleep(REQUEST_DELAY)

    print(f"[DONE] {cfg['name']}爬取完成！共 {len(all_items)} 条。")
    return all_items


def crawl_medals(uid: str) -> list:
    """粉丝牌：单次请求"""
    cfg = DATA_TYPES["medal"]
    print(f"\n[INFO] 开始查询 {cfg['name']}...")
    data = api_get(cfg["endpoint"], cfg["params"](uid, 1, 1))
    if data.get("code") != 0:
        print(f"[ERROR] 无法获取粉丝牌数据")
        return []
    medals = data.get("data", {}).get("list", [])
    print(f"[DONE] 粉丝牌获取完成！共 {len(medals)} 个。")
    return medals


def crawl_livedm(uid: str, page_size: int) -> list:
    """直播弹幕：特殊结构，按直播间分组，需要展平为行"""
    cfg = DATA_TYPES["livedm"]
    print(f"\n[INFO] 开始查询 {cfg['name']}...")
    all_danmaku_rows = []
    page = 1

    first = api_get(cfg["endpoint"], cfg["params"](uid, 1, page_size))
    if first.get("code") != 0:
        print(f"[ERROR] 无法获取{cfg['name']}数据")
        return []

    total = first.get("data", {}).get("cursor", {}).get("all_count", 0)
    is_end = first.get("data", {}).get("cursor", {}).get("is_end", True)
    total_pages = (total - 1) // page_size + 1 if total > 0 else 1
    print(f"[INFO] 总直播弹幕组数: {total}  总页数: {total_pages}")

    # 展平第一页
    for group in first.get("data", {}).get("list", []):
        ri = group.get("roominfo", {})
        for dm in group.get("danmu", []):
            all_danmaku_rows.append({
                "room_id": ri.get("roomid", ""),
                "up_name": ri.get("upname", ""),
                "up_uid": ri.get("upuid", ""),
                "room_name": ri.get("roomname", ""),
                "uname": dm.get("uname", ""),
                "text": dm.get("text", ""),
                "ts": dm.get("ts", 0),
            })
    print(f"  第 1/{total_pages} 页 | 累计: {len(all_danmaku_rows)} 条弹幕")

    time.sleep(REQUEST_DELAY)

    while not is_end:
        page += 1
        data = api_get(cfg["endpoint"], cfg["params"](uid, page, page_size))
        if data.get("code") != 0:
            continue
        is_end = data.get("data", {}).get("cursor", {}).get("is_end", True)
        for group in data.get("data", {}).get("list", []):
            ri = group.get("roominfo", {})
            for dm in group.get("danmu", []):
                all_danmaku_rows.append({
                    "room_id": ri.get("roomid", ""),
                    "up_name": ri.get("upname", ""),
                    "up_uid": ri.get("upuid", ""),
                    "room_name": ri.get("roomname", ""),
                    "uname": dm.get("uname", ""),
                    "text": dm.get("text", ""),
                    "ts": dm.get("ts", 0),
                })
        print(f"  第 {page}/{total_pages} 页 | 累计: {len(all_danmaku_rows)} 条弹幕")
        time.sleep(REQUEST_DELAY)

    print(f"[DONE] 直播弹幕爬取完成！共 {len(all_danmaku_rows)} 条。")
    return all_danmaku_rows


# ---------------------------------------------------------------------------
# B站视频标题
# ---------------------------------------------------------------------------
def fetch_video_titles(oids: list, cache_dir: str = ".") -> dict:
    cache_path = os.path.join(cache_dir, TITLE_CACHE_NAME)
    cache = {}
    if os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as f:
            cache = json.load(f)

    uncached = [o for o in oids if str(o) not in cache and o]
    if not uncached:
        print(f"[INFO] 视频标题: 全部命中缓存 ({len(oids)} 个)")
        return {str(k): v for k, v in cache.items()}

    print(f"[INFO] 视频标题: 需查询 {len(uncached)}/{len(oids)} 个")

    for i, oid in enumerate(uncached, 1):
        title = "(获取失败)"
        for attempt in range(MAX_RETRIES):
            try:
                resp = requests.get(BILIBILI_VIDEO_INFO_API, params={"aid": oid},
                                    timeout=REQUEST_TIMEOUT,
                                    headers={"User-Agent": UA, "Referer": "https://www.bilibili.com/"})
                if resp.status_code == 412:
                    time.sleep(3)
                    continue
                data = resp.json()
                if data.get("code") == 0:
                    title = data["data"]["title"]
                elif data.get("code") == -404:
                    title = "(视频不存在或已删除)"
                break
            except:
                if attempt == MAX_RETRIES - 1:
                    title = "(获取失败)"
        cache[str(oid)] = title
        if i % 20 == 0 or i == len(uncached):
            pct = "" if i == len(uncached) else f"  (最近: {title[:40]}...)"
            try:
                print(f"  查询进度: {i}/{len(uncached)}{pct}")
            except UnicodeEncodeError:
                safe_title = title[:40].encode('ascii', errors='replace').decode('ascii')
                print(f"  查询进度: {i}/{len(uncached)}  (最近: {safe_title}...)")
        if i % 10 == 0:
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False, indent=2)
        time.sleep(BILI_DELAY)

    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    print(f"[DONE] 视频标题获取完成！")
    return {str(k): v for k, v in cache.items()}


def collect_oids_from_data(comments, videodm_rows, livedm_rows) -> set:
    """从所有数据类型中收集视频 oid"""
    oids = set()
    for c in comments:
        oid = c.get("dyn", {}).get("oid", "")
        if oid: oids.add(str(oid))
    for dm in videodm_rows:
        oid = dm.get("oid", "")
        if oid: oids.add(str(oid))
    return oids


# ---------------------------------------------------------------------------
# 展平：评论
# ---------------------------------------------------------------------------
def flatten_comments(comments: list, titles: dict = None) -> list:
    rows = []
    for idx, c in enumerate(comments, 1):
        rpid = c.get("rpid", "")
        oid = c.get("dyn", {}).get("oid", "")
        content_type = c.get("dyn", {}).get("type", 1)
        ts = c.get("time", 0)
        parent = c.get("parent", {})

        row = {
            "序号": idx,
            "评论ID": rpid,
            "评论内容": c.get("message", ""),
            "评论时间": ts_to_beijing(ts),
            "层级": rank_label(c.get("rank", 0)),
            "根评论ID": parent.get("rootid", ""),
            "父评论ID": parent.get("parentid", ""),
            "视频ID": oid,
            "视频链接": oid_to_url(oid),
            "视频标题": titles.get(str(oid), "") if titles else "",
            "评论跳转链接": comment_jump_url(oid, rpid),
        }
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 展平：视频弹幕
# ---------------------------------------------------------------------------
def flatten_videodm(dmlist: list, titles: dict = None) -> list:
    rows = []
    for idx, dm in enumerate(dmlist, 1):
        oid = dm.get("oid", "")
        row = {
            "序号": idx,
            "弹幕ID": dm.get("id", ""),
            "弹幕内容": dm.get("content", ""),
            "发送时间": ts_to_beijing(dm.get("ctime", 0)),
            "视频ID": oid,
            "视频链接": oid_to_url(oid),
            "视频标题": titles.get(str(oid), "") if titles else "",
            "出现时间点": millis_to_time_str(dm.get("progress", 0)),
            "进度(毫秒)": dm.get("progress", 0),
        }
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 展平：直播弹幕
# ---------------------------------------------------------------------------
def flatten_livedm(livedm_rows: list) -> list:
    rows = []
    for idx, dm in enumerate(livedm_rows, 1):
        rows.append({
            "序号": idx,
            "直播间ID": dm.get("room_id", ""),
            "主播昵称": dm.get("up_name", ""),
            "主播UID": dm.get("up_uid", ""),
            "直播间标题": dm.get("room_name", ""),
            "弹幕内容": dm.get("text", ""),
            "发送时间": ts_to_beijing(dm.get("ts", 0)),
        })
    return rows


# ---------------------------------------------------------------------------
# 展平：粉丝牌
# ---------------------------------------------------------------------------
def flatten_medals(medals: list) -> list:
    rows = []
    for idx, m in enumerate(medals, 1):
        ruid = m.get("ruid", "")
        rows.append({
            "序号": idx,
            "粉丝牌名称": m.get("name", ""),
            "粉丝牌等级": m.get("level", 0),
            "主播UID": ruid,
            "主播空间": f"https://space.bilibili.com/{ruid}" if ruid else "",
        })
    return rows


# ---------------------------------------------------------------------------
# 保存函数
# ---------------------------------------------------------------------------
def save_csv(rows: list, uid: str, dtype: str, output_dir: str):
    cfg = DATA_TYPES[dtype]
    filename = os.path.join(output_dir, f"{cfg['csv_prefix']}_{uid}.csv")
    if not rows: return
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"[OUTPUT] CSV: {filename}")


def save_json(data, uid: str, dtype: str, output_dir: str):
    cfg = DATA_TYPES[dtype]
    filename = os.path.join(output_dir, f"{cfg['csv_prefix']}_{uid}_raw.json")
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OUTPUT] JSON: {filename}")


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="AICU.cc B站用户数据爬虫 — 评论/弹幕/粉丝牌一键爬取",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python aicu_crawler.py 1703916229                     # 查询全部
  python aicu_crawler.py 1703916229 --type reply        # 仅评论
  python aicu_crawler.py 1703916229 --type medal        # 仅粉丝牌
  python aicu_crawler.py 1703916229 --type videodm      # 仅视频弹幕
  python aicu_crawler.py 1703916229 --type livedm       # 仅直播弹幕
        """
    )
    parser.add_argument("uid", help="B站用户 UID")
    parser.add_argument("--type", "-t", choices=["reply", "medal", "videodm", "livedm", "all"],
                        default="all", help="数据类型 (默认: all)")
    parser.add_argument("--ps", type=int, default=DEFAULT_PAGE_SIZE, help=f"每页条数 (默认: {DEFAULT_PAGE_SIZE})")
    parser.add_argument("--output-dir", "-o", default="output", help="输出目录")
    parser.add_argument("--json", action="store_true", help="同时输出原始 JSON")
    parser.add_argument("--no-titles", action="store_true", help="跳过视频标题查询")

    args = parser.parse_args()
    os.makedirs(args.output_dir, exist_ok=True)

    select_types = ["reply", "medal", "videodm", "livedm"] if args.type == "all" else [args.type]

    all_comments_raw = []
    all_videodm_raw = []
    all_livedm_raw = []
    all_medals_raw = []
    all_comment_rows = []
    all_videodm_rows = []
    all_livedm_rows_flat = []
    all_medal_rows = []
    titles = {}

    # 1) 爬取各类数据
    if "reply" in select_types:
        all_comments_raw = crawl_paginated(args.uid, "reply", args.ps)
        time.sleep(REQUEST_DELAY)

    if "videodm" in select_types:
        all_videodm_raw = crawl_paginated(args.uid, "videodm", args.ps)
        time.sleep(REQUEST_DELAY)

    if "livedm" in select_types:
        all_livedm_raw = crawl_livedm(args.uid, args.ps)
        time.sleep(REQUEST_DELAY)

    if "medal" in select_types:
        all_medals_raw = crawl_medals(args.uid)

    # 2) 获取视频标题
    if not args.no_titles:
        oids = collect_oids_from_data(all_comments_raw, all_videodm_raw, all_livedm_raw)
        if oids:
            print(f"\n[INFO] 去重后共 {len(oids)} 个视频需要查询标题...")
            titles = fetch_video_titles(list(oids), args.output_dir)

    # 3) 展平
    if all_comments_raw:
        all_comment_rows = flatten_comments(all_comments_raw, titles)
    if all_videodm_raw:
        all_videodm_rows = flatten_videodm(all_videodm_raw, titles)
    if all_medals_raw:
        all_medal_rows = flatten_medals(all_medals_raw)
    if all_livedm_raw:
        all_livedm_rows_flat = flatten_livedm(all_livedm_raw)

    # 4) 输出
    print()
    if all_comment_rows:
        save_csv(all_comment_rows, args.uid, "reply", args.output_dir)
    if all_videodm_rows:
        save_csv(all_videodm_rows, args.uid, "videodm", args.output_dir)
    if all_livedm_rows_flat:
        save_csv(all_livedm_rows_flat, args.uid, "livedm", args.output_dir)
    if all_medal_rows:
        save_csv(all_medal_rows, args.uid, "medal", args.output_dir)

    if args.json:
        if all_comments_raw:
            save_json(all_comments_raw, args.uid, "reply", args.output_dir)
        if all_videodm_raw:
            save_json(all_videodm_raw, args.uid, "videodm", args.output_dir)
        if all_livedm_raw:
            save_json(all_livedm_raw, args.uid, "livedm", args.output_dir)
        if all_medals_raw:
            save_json(all_medals_raw, args.uid, "medal", args.output_dir)

    # 打印摘要
    print(f"\n{'='*50}")
    print("数据摘要:")
    if all_comment_rows:
        print(f"  评论:     {len(all_comment_rows)} 条")
    if all_videodm_rows:
        print(f"  视频弹幕: {len(all_videodm_rows)} 条")
    if all_livedm_rows_flat:
        print(f"  直播弹幕: {len(all_livedm_rows_flat)} 条")
    if all_medal_rows:
        print(f"  粉丝牌:   {len(all_medal_rows)} 个")
    print(f"\n[DONE] 全部完成！数据保存在: {args.output_dir}/")


if __name__ == "__main__":
    main()
