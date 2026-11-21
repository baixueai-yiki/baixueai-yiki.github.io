#!/usr/bin/env python3
"""把 Content/Posts 下的文章列成清单，供网页画按钮。

约定（不需要任何配置）：
    Content/Posts/<标签页>/<分组>/<标题>.pdf    ← 推荐，浏览器原样显示
    Content/Posts/<标签页>/<分组>/<标题>.html   ← 兼容早期直接写的网页片段

    文件夹名 = 折叠块的分组
    文件名   = 按钮上的字
    <标题>/ 文件夹 = 这篇自己的图片（html 文章才会用到）

跑完在每个标签页目录下写一份 index.json。

用法：
    python tools/build_posts.py
"""

import io
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
POSTS = os.path.join(ROOT, "Content", "Posts")

# 能直接给浏览器看的后缀
KINDS = {".pdf": "pdf", ".html": "html"}
# 需要他自己另存一下的
NEEDS_EXPORT = {".docx": "pdf", ".doc": "pdf"}


def natural_key(name):
    # 数字按大小排，其余按字面排；统一成 (类型, 值) 免得 int 和 str 比大小报错
    return [(0, int(p)) if p.isdigit() else (1, p.lower()) for p in re.split(r"(\d+)", name)]


# 文件夹名可以写成「1-教程」「2 作品」，前面的数字只用来排序，标题里不显示
ORDER_RE = re.compile(r"^(\d+)\s*[-_.、\s]\s*(.+)$")


def group_title(name):
    match = ORDER_RE.match(name)
    return match.group(2).strip() if match else name


def group_key(name):
    match = ORDER_RE.match(name)
    if match:
        return (0, int(match.group(1)))
    return (1, natural_key(name))


def rel_to_root(path):
    """转成从网站根开始的路径，这样内容页（在根目录）引用也不会错。"""
    return "/" + os.path.relpath(path, ROOT).replace("\\", "/")


def build_tab(tab_dir):
    groups = []

    for entry in sorted(os.listdir(tab_dir), key=group_key):
        path = os.path.join(tab_dir, entry)
        if not os.path.isdir(path) or entry.startswith("."):
            continue

        items = []
        for fname in sorted(os.listdir(path), key=natural_key):
            full = os.path.join(path, fname)
            if not os.path.isfile(full) or fname.startswith("~$") or fname.startswith("."):
                continue

            title, ext = os.path.splitext(fname)
            ext = ext.lower()

            if ext in KINDS:
                items.append({
                    "title": title,
                    "path": rel_to_root(full).lstrip("/"),
                    "kind": KINDS[ext],
                })
            elif ext in NEEDS_EXPORT:
                print("  ! %s 是 %s，浏览器不能直接显示，请先另存为 %s"
                      % (fname, ext, NEEDS_EXPORT[ext]))

        if items:
            groups.append({"title": group_title(entry), "folder": entry, "items": items})

    return groups


def main():
    if not os.path.isdir(POSTS):
        print("没有找到", POSTS)
        return 1

    for tab in sorted(os.listdir(POSTS), key=natural_key):
        tab_dir = os.path.join(POSTS, tab)
        if not os.path.isdir(tab_dir):
            continue

        print("[%s]" % tab)
        groups = build_tab(tab_dir)
        index = {"tab": tab, "groups": groups}

        out = os.path.join(tab_dir, "index.json")
        with io.open(out, "w", encoding="utf-8", newline="\n") as f:
            json.dump(index, f, ensure_ascii=False, indent=2)
            f.write("\n")

        total = sum(len(g["items"]) for g in groups)
        print("  → index.json：%d 组 %d 篇" % (len(groups), total))

    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
