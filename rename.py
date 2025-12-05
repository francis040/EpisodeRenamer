#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
剧集自动重命名 v3.1（支持完整设置持久化）
- 配置保存到软件同目录的 config.json
- 记住所有设置项：剧名、季数、补零位数、偏移量、递归、Dry-run、模板、排序方式、集标题开关、冲突后缀、季文件夹开关、最近文件夹
- 其余逻辑保持 v3.0 行为不变
"""

import os, re, sys, csv, time, json, shutil
import concurrent.futures
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
from pathlib import Path

# optional: tkinterdnd2 for drag & drop
USE_TKINTERDND = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    USE_TKINTERDND = True
except Exception:
    USE_TKINTERDND = False

# ======================================================
# 应用目录 & 配置文件路径：改为软件同目录
# ======================================================
def get_app_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

APP_DIR = get_app_dir()
CONFIG_FILE = os.path.join(APP_DIR, "config.json")

# ========== 配置 ==========
VIDEO_EXTS = {".mp4", ".mkv", ".avi", ".mov", ".flv", ".wmv", ".ts", ".m4v"}
AUTO_PREVIEW_DEBOUNCE_MS = 300
MAX_RECENT = 10
EXECUTOR = concurrent.futures.ThreadPoolExecutor(max_workers=4)

# stop tokens used when extracting episode title/extra info
_STOP_TOKENS = set([
    "2160p","1080p","720p","480p","WEB-DL","WEB","WEBRip","WEB-DLRip","WEB-DL2",
    "BluRay","BRRip","BDRip","HDRip","HDTV","DVDRip",
    "x264","x265","h264","h265","hevc","AVC","AAC","DDP","DDP5.1","DD+","Atmos",
    "HEVC","10bit","8bit","PROPER","REPACK","EXTENDED","UNRATED"
])
_STOP_TOKENS = {t.lower() for t in _STOP_TOKENS}

# ======================================================
# 工具函数
# ======================================================
def natural_key(s):
    parts = re.split(r'(\d+)', s)
    key = []
    for p in parts:
        if p.isdigit():
            key.append(int(p))
        else:
            key.append(p.lower())
    return key

def is_video(filename):
    return os.path.splitext(filename)[1].lower() in VIDEO_EXTS

# ======================================================
# 智能识别 A 方案： parse_episode_info
# ======================================================
def parse_episode_info(filename):
    """
    高级规则解析：
    输入：单个文件名（含扩展）
    输出： (season:int|None, episode:int|None, extra_title:str|None)
    - extra_title 是从文件名中 SxxEyy 之后提取的“集标题”（直到遇到常见技术标签）
    """
    name = os.path.splitext(os.path.basename(filename))[0]
    low = name.lower()

    # helper to extract trailing title tokens after a match position
    def _extract_after_pos(s, pos):
        # split remainder by dot / space / underscore / dash
        rem = s[pos:].strip(" ._-")
        if not rem:
            return None
        tokens = re.split(r'[._\-\s]+', rem)
        out_tokens = []
        for tok in tokens:
            if not tok:
                continue
            if tok.lower() in _STOP_TOKENS:
                break
            # if token looks like resolution/codec (e.g. 2160p or 1080p or 5.1) break
            if re.fullmatch(r'\d{3,4}p', tok.lower()):
                break
            if re.fullmatch(r'\d+\.\d+', tok):
                # audio channel e.g. 5.1
                break
            # skip pure release-group bracketed tokens like [group]
            if re.fullmatch(r'\[.*\]|\(.*\)', tok):
                continue
            out_tokens.append(tok)
        if not out_tokens:
            return None
        # join using '.' to mimic scene naming
        return ".".join(out_tokens)

    # 1) SxxEyy (most reliable)
    m = re.search(r'[sS](\d{1,2})[ ._-]?[eE](\d{1,3})', name)
    if m:
        season = int(m.group(1))
        episode = int(m.group(2))
        extra = _extract_after_pos(name, m.end())
        return season, episode, extra

    # 2) 2x03 style
    m = re.search(r'(\d{1,2})[xX](\d{1,3})', name)
    if m:
        season = int(m.group(1))
        episode = int(m.group(2))
        extra = _extract_after_pos(name, m.end())
        return season, episode, extra

    # 3) "Season 2 Episode 3" / "Season 2 Ep 03"
    m = re.search(r'[sS]eason[ ._-]*?(\d{1,2}).*?[eE]p(?:isode)?[ ._-]*?(\d{1,3})', name, flags=re.IGNORECASE)
    if m:
        season = int(m.group(1)); episode = int(m.group(2))
        extra = _extract_after_pos(name, m.end())
        return season, episode, extra

    # 4) Chinese: "第2季 第3集"
    m = re.search(r'第\s*(\d{1,2})\s*季', name)
    if m:
        season = int(m.group(1))
        m2 = re.search(r'第\s*(\d{1,3})\s*[集话回]', name)
        if m2:
            episode = int(m2.group(1))
            extra = _extract_after_pos(name, m2.end())
            return season, episode, extra
        # no episode info, fallback
        return season, None, None

    # 5) Chinese only "第12集" or "第12话"
    m = re.search(r'第\s*(\d{1,3})\s*[集话回]', name)
    if m:
        episode = int(m.group(1))
        extra = _extract_after_pos(name, m.end())
        return None, episode, extra

    # 6) look for pattern like "Episode 12" English
    m = re.search(r'[eE]pisode[ ._-]*?(\d{1,3})', name)
    if m:
        episode = int(m.group(1))
        extra = _extract_after_pos(name, m.end())
        return None, episode, extra

    # 7) fallback: first standalone number (but this is ambiguous)
    nums = re.findall(r'\b(\d{1,3})\b', name)
    if nums:
        # return first numeric, leave season None
        episode = int(nums[0])
        # attempt to grab remainder after that numeric token
        first_pos = re.search(r'\b' + re.escape(nums[0]) + r'\b', name)
        if first_pos:
            extra = _extract_after_pos(name, first_pos.end())
        else:
            extra = None
        return None, episode, extra

    # nothing found
    return None, None, None

# ======================================================
# 模板格式化
# ======================================================
def format_template(template, title, season, episode, ext, orig):
    try:
        s = 0 if season is None else int(season)
        e = 0 if episode is None else int(episode)
        context = {"title": title, "season": s, "episode": e, "ep": e, "ext": ext.lstrip("."), "orig": orig}
        return template.format(**context)
    except Exception:
        out = template.replace("{title}", title).replace("{ext}", ext.lstrip(".")).replace("{orig}", orig)
        out = out.replace("{season}", str(season if season is not None else 0))
        out = out.replace("{episode}", str(episode if episode is not None else 0)).replace("{ep}", str(episode if episode is not None else 0))
        return out

# ======================================================
# 配置管理：保存所有设置到同目录 config.json
# ======================================================
DEFAULT_CONFIG = {
    "recent_folders": [],
    "title": "",
    "season": "1",
    "pad": 3,
    "offset": 0,
    "recursive": False,
    "move_season_folder": False,
    "template": "{title}.S{season:02}E{episode:03}.{ext}",
    "conflict_suffix": "_dup",
    "dryrun": True,
    "sort_method": "name",  # name / guess / numeric
    "include_episode_title": True,
}

def load_config():
    # 选项一：如果存在就加载；不存在就创建默认再返回
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            # 补全缺失字段，兼容未来扩展
            for k, v in DEFAULT_CONFIG.items():
                data.setdefault(k, v)
            return data
        except Exception:
            # 如果读取出错，保守起见：仍然返回 DEFAULT_CONFIG，并覆盖写入一份新的
            pass

    save_config(DEFAULT_CONFIG)
    return dict(DEFAULT_CONFIG)

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("保存配置失败:", e)

# ======================================================
# 主类
# ======================================================
class SeriesRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("剧集自动重命名（专业智能版 v3.1）")
        self.root.geometry("1080x700")

        # 先加载配置
        self.cfg = load_config()

        # 数据
        self.folders = []
        self.file_list = []
        self.preview_list = []
        self.log_rows = []
        self.configured_tags = set()  # 保留你朋友的补丁：已配置 tag 集合

        # debounce
        self._debounce_after_id = None

        # variables（初始值从配置恢复）
        self.var_recursive = tk.BooleanVar(value=bool(self.cfg.get("recursive", False)))
        self.var_template = tk.StringVar(value=self.cfg.get("template", DEFAULT_CONFIG["template"]))
        self.var_title = tk.StringVar(value=self.cfg.get("title", ""))
        self.var_season = tk.StringVar(value=self.cfg.get("season", "1"))
        self.var_pad = tk.IntVar(value=int(self.cfg.get("pad", 3)))
        self.var_offset = tk.IntVar(value=int(self.cfg.get("offset", 0)))
        self.var_move_season_folder = tk.BooleanVar(value=bool(self.cfg.get("move_season_folder", False)))
        self.var_conflict_suffix = tk.StringVar(value=self.cfg.get("conflict_suffix", "_dup"))
        self.var_dryrun = tk.BooleanVar(value=bool(self.cfg.get("dryrun", True)))
        self.var_sort_method = tk.StringVar(value=self.cfg.get("sort_method", "name"))
        # 新增：是否在新文件名中追加解析到的集标题
        self.var_include_episode_title = tk.BooleanVar(value=bool(self.cfg.get("include_episode_title", True)))

        # UI
        self._build_ui()
        self._load_recent()
        self._bind_auto_preview()

        # dnd
        if USE_TKINTERDND:
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self._on_drop)
            except Exception:
                pass
        else:
            self.status_set("提示：未检测到 tkinterdnd2，拖拽功能需安装（pip install tkinterdnd2）。仍可用“添加文件夹”。")

        # 启动时立即保存一次（保证 config.json 完整）
        self._save_all_config()

    # ----- UI -----
    def _build_ui(self):
        top = ttk.Frame(self.root)
        top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        btn_add = ttk.Button(top, text="添加文件夹", command=self.add_folder)
        btn_add.pack(side=tk.LEFT)
        btn_remove = ttk.Button(top, text="移除选中文件夹", command=self.remove_selected_folder)
        btn_remove.pack(side=tk.LEFT, padx=6)
        btn_scan = ttk.Button(top, text="扫描", command=self.scan_async)
        btn_scan.pack(side=tk.LEFT, padx=6)
        btn_preview = ttk.Button(top, text="强制预览", command=self.preview)
        btn_preview.pack(side=tk.LEFT, padx=6)
        btn_rename = ttk.Button(top, text="重命名 (执行)", command=self.execute_async)
        btn_rename.pack(side=tk.LEFT, padx=6)
        btn_undo = ttk.Button(top, text="撤销（Undo）", command=self.undo_via_log)
        btn_undo.pack(side=tk.LEFT, padx=6)

        left = ttk.Frame(self.root)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=6)

        ttk.Label(left, text="文件夹:").pack(anchor=tk.W)
        self.lb_folders = tk.Listbox(left, width=44, height=8)
        self.lb_folders.pack(fill=tk.Y)
        self.lb_folders.bind("<Double-Button-1>", self._open_folder_at_selection)

        chk_recursive = ttk.Checkbutton(left, text="递归扫描子文件夹", variable=self.var_recursive, command=self._on_setting_change)
        chk_recursive.pack(anchor=tk.W, pady=4)

        settings = ttk.LabelFrame(left, text="设置")
        settings.pack(fill=tk.X, pady=6)

        ttk.Label(settings, text="剧名:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        e_title = ttk.Entry(settings, textvariable=self.var_title, width=28)
        e_title.grid(row=0, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(settings, text="季 (数字):").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2)
        e_season = ttk.Entry(settings, textvariable=self.var_season, width=6)
        e_season.grid(row=1, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(settings, text="集数补零位数:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=2)
        sb_pad = ttk.Spinbox(settings, from_=1, to=6, textvariable=self.var_pad, width=4)
        sb_pad.grid(row=2, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(settings, text="集数偏移 (offset):").grid(row=3, column=0, sticky=tk.W, padx=4, pady=2)
        e_offset = ttk.Entry(settings, textvariable=self.var_offset, width=6)
        e_offset.grid(row=3, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Checkbutton(settings, text="移动到季子文件夹 Sxx", variable=self.var_move_season_folder, command=self._on_setting_change).grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=4, pady=2)

        ttk.Label(settings, text="冲突后缀:").grid(row=5, column=0, sticky=tk.W, padx=4, pady=2)
        e_conflict = ttk.Entry(settings, textvariable=self.var_conflict_suffix, width=10)
        e_conflict.grid(row=5, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(settings, text="命名模板:").grid(row=6, column=0, sticky=tk.W, padx=4, pady=2)
        e_template = ttk.Entry(settings, textvariable=self.var_template, width=34)
        e_template.grid(row=6, column=1, sticky=tk.W, padx=4, pady=2)
        ttk.Label(settings, text="占位符: {title} {season} {episode} {ext} {orig}").grid(row=7, column=0, columnspan=2, sticky=tk.W, padx=4, pady=2)

        ttk.Label(settings, text="排序方式:").grid(row=8, column=0, sticky=tk.W, padx=4, pady=2)
        cb_sort = ttk.Combobox(settings, values=["name", "guess", "numeric"], textvariable=self.var_sort_method, width=12, state="readonly")
        cb_sort.grid(row=8, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Checkbutton(settings, text="仅预览（Dry-run）", variable=self.var_dryrun, command=self._on_setting_change).grid(row=9, column=0, columnspan=2, sticky=tk.W, padx=4, pady=2)

        # 新增开关：是否在新文件名中追加解析到的集标题
        ttk.Checkbutton(settings, text="在新文件名中追加集标题（如有）", variable=self.var_include_episode_title, command=self._on_setting_change).grid(row=10, column=0, columnspan=2, sticky=tk.W, padx=4, pady=2)

        right = ttk.Frame(self.root)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=6)

        preview_frame = ttk.LabelFrame(right, text="预览（原文件名 → 新文件名）")
        preview_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("oldname", "newname")
        self.tree = ttk.Treeview(preview_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("oldname", text="原文件名")
        self.tree.heading("newname", text="新文件名")
        self.tree.column("oldname", width=300, anchor="w")
        self.tree.column("newname", width=500, anchor="w")
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # style/tags
        self.tree.tag_configure("conflict", foreground="red")
        self.tree.tag_configure("unchanged", foreground="#888888")
        self.ext_colors = {}
        self._default_ext_colors()

        bottom = ttk.Frame(self.root)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, padx=8, pady=6)
        self.status_label = ttk.Label(bottom, text="就绪")
        self.status_label.pack(side=tk.LEFT)
        ttk.Button(bottom, text="帮助 / 关于", command=self.show_help).pack(side=tk.RIGHT)

        # store widgets to bind debounce
        self._entry_widgets = [e_title, e_season, sb_pad, e_offset, e_conflict, e_template, cb_sort]

    def _default_ext_colors(self):
        palette = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
        exts = [".mkv", ".mp4", ".avi", ".mov", ".flv", ".wmv"]
        for i, e in enumerate(exts):
            self.ext_colors[e] = palette[i % len(palette)]

    # ========== 配置保存辅助 ==========
    def _save_all_config(self):
        cfg = dict(self.cfg) if hasattr(self, "cfg") else dict(DEFAULT_CONFIG)
        # 最近文件夹（只存存在的）
        recent = [p for p in self.folders if os.path.isdir(p)]
        recent = recent[-MAX_RECENT:]
        cfg["recent_folders"] = recent

        cfg["title"] = self.var_title.get()
        cfg["season"] = self.var_season.get()
        # 防止非法输入导致崩溃
        try:
            cfg["pad"] = int(self.var_pad.get())
        except Exception:
            cfg["pad"] = DEFAULT_CONFIG["pad"]
        try:
            cfg["offset"] = int(self.var_offset.get())
        except Exception:
            cfg["offset"] = DEFAULT_CONFIG["offset"]

        cfg["recursive"] = bool(self.var_recursive.get())
        cfg["move_season_folder"] = bool(self.var_move_season_folder.get())
        cfg["template"] = self.var_template.get()
        cfg["conflict_suffix"] = self.var_conflict_suffix.get()
        cfg["dryrun"] = bool(self.var_dryrun.get())
        cfg["sort_method"] = self.var_sort_method.get()
        cfg["include_episode_title"] = bool(self.var_include_episode_title.get())

        self.cfg = cfg
        save_config(cfg)

    # ========= drag & drop handling =========
    def _on_drop(self, event):
        data = event.data
        paths = self._parse_dnd_paths(data)
        added = False
        for p in paths:
            if os.path.isdir(p) and p not in self.folders:
                self._add_folder_internal(p)
                added = True
        if added:
            self.status_set("检测到拖拽文件夹，已添加并开始扫描")
            self.scan_async()

    @staticmethod
    def _parse_dnd_paths(data):
        res = []
        cur = ""
        inbrace = False
        for ch in data:
            if ch == "{":
                inbrace = True
                cur = ""
            elif ch == "}":
                inbrace = False
                res.append(cur)
                cur = ""
            elif inbrace:
                cur += ch
            elif ch.isspace():
                if cur:
                    res.append(cur)
                    cur = ""
            else:
                cur += ch
        if cur:
            res.append(cur)
        return res

    # ========== recent folders ==========
    def _load_recent(self):
        cfg = getattr(self, "cfg", load_config())
        recent = cfg.get("recent_folders", [])
        for p in recent:
            if os.path.isdir(p) and p not in self.folders:
                self.folders.append(p)
                self.lb_folders.insert(tk.END, p)
        if self.folders:
            self.status_set(f"已加载 {len(self.folders)} 个常用文件夹（来自配置）")
            self.scan_async()

    def _save_recent(self):
        # 现在所有保存统一走 _save_all_config
        self._save_all_config()

    def _add_folder_internal(self, path):
        self.folders.append(path)
        self.lb_folders.insert(tk.END, path)
        self._save_recent()

    # ========== binding auto preview ==========
    def _bind_auto_preview(self):
        for w in self._entry_widgets:
            try:
                w.bind("<KeyRelease>", lambda e: self._debounce_preview())
                if isinstance(w, ttk.Combobox):
                    w.bind("<<ComboboxSelected>>", lambda e: self._on_setting_change())
            except Exception:
                pass

    def _debounce_preview(self):
        # 设置变化时顺便保存配置
        try:
            self._save_all_config()
        except Exception:
            pass
        if self._debounce_after_id:
            self.root.after_cancel(self._debounce_after_id)
        self._debounce_after_id = self.root.after(AUTO_PREVIEW_DEBOUNCE_MS, self.preview)

    def _on_setting_change(self):
        # 保存配置 + 自动预览
        try:
            self._save_all_config()
        except Exception:
            pass
        self._debounce_preview()

    # ========== folder ops ==========
    def add_folder(self):
        path = filedialog.askdirectory()
        if path:
            if path not in self.folders:
                self._add_folder_internal(path)
                self.status_set(f"添加文件夹: {path}")
                self.scan_async()
                self._debounce_preview()

    def remove_selected_folder(self):
        sel = self.lb_folders.curselection()
        if not sel:
            return
        idx = sel[0]
        folder = self.lb_folders.get(idx)
        self.folders.remove(folder)
        self.lb_folders.delete(idx)
        self._save_recent()
        self.status_set(f"移除文件夹: {folder}")
        self.scan_async()
        self._debounce_preview()

    def _open_folder_at_selection(self, event):
        sel = self.lb_folders.curselection()
        if not sel: return
        idx = sel[0]
        folder = self.lb_folders.get(idx)
        if os.path.isdir(folder):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(folder)
                elif sys.platform.startswith("darwin"):
                    os.system(f'open "{folder}"')
                else:
                    os.system(f'xdg-open "{folder}"')
            except Exception:
                pass

    # ========== scanning (线程池) ==========
    def scan_async(self):
        self.status_set("开始扫描（后台线程）...")
        EXECUTOR.submit(self._scan_task)

    def _scan_task(self):
        files = []
        seen = set()
        for folder in list(self.folders):
            for root, dirs, filenames in os.walk(folder) if self.var_recursive.get() else [(folder, [], os.listdir(folder))]:
                for f in filenames:
                    if is_video(f):
                        p = os.path.join(root, f)
                        if p not in seen:
                            seen.add(p)
                            files.append(p)
        files_sorted = sorted(files, key=lambda p: natural_key(os.path.basename(p)))
        self.root.after(0, self._on_scan_done, files_sorted)

    def _on_scan_done(self, files_sorted):
        self.file_list = files_sorted
        self.status_set(f"扫描完成：{len(self.file_list)} 个视频文件")
        self.preview()

    # ========== preview logic ==========
    def preview(self):
        # clear tree
        for it in self.tree.get_children():
            self.tree.delete(it)
        self.preview_list.clear()
        self.configured_tags.clear()  # reset configured tags for new preview
        if not self.file_list:
            self.status_set("没有可预览的文件")
            return

        method = self.var_sort_method.get()
        files = list(self.file_list)
        if method == "name":
            files = sorted(files, key=lambda p: natural_key(os.path.basename(p)))
        elif method == "numeric":
            files = sorted(files, key=lambda p: [int(x) if x.isdigit() else x.lower() for x in re.split(r'(\d+)', os.path.basename(p))])
        elif method == "guess":
            # use our smarter parse for ordering: if many files have parseable episode numbers, sort by them
            def guess_key(p):
                s_auto, e_auto, _ = parse_episode_info(p)
                if e_auto is not None:
                    return (0, e_auto, natural_key(os.path.basename(p)))
                return (1, natural_key(os.path.basename(p)))
            files = sorted(files, key=guess_key)

        # Use parse_episode_info across files to decide numbering strategy
        parsed = []
        for p in files:
            s_auto, e_auto, extra_auto = parse_episode_info(p)
            parsed.append((p, s_auto, e_auto, extra_auto))

        # Determine whether we should trust parsed episode numbers:
        parsed_es = [e for (_, s, e, ex) in parsed if e is not None]
        use_parsed_numbers = False
        if len(parsed_es) >= max(2, len(files)//3):
            use_parsed_numbers = True

        # Also detect "pure numeric filenames" as sequence (e.g., 01.mkv, 02.mkv)
        pure_numeric_files = []
        for p, s, e, ex in parsed:
            name_no_ext = os.path.splitext(os.path.basename(p))[0]
            if re.fullmatch(r'\d{1,3}', name_no_ext):
                pure_numeric_files.append(p)

        pad = int(self.var_pad.get())
        offset = int(self.var_offset.get())
        season_input = self.var_season.get().strip()
        season_val = int(season_input) if season_input.isdigit() else None
        title = self.var_title.get().strip() or "Series"
        include_title = bool(self.var_include_episode_title.get())

        # Build preview list using parse results and strategy:
        if use_parsed_numbers:
            # sort by parsed episode (if available), fallback to natural name
            parsed_sorted = sorted(parsed, key=lambda x: (x[2] if x[2] is not None else 999999, natural_key(os.path.basename(x[0]))))
            for idx, (p, s_auto, e_auto, extra_auto) in enumerate(parsed_sorted):
                # season: user input overrides auto if provided
                season_final = season_val if season_val is not None else s_auto
                # episode:
                ep = (e_auto if e_auto is not None else (idx + 1)) + offset
                if season_final is None:
                    season_final = 1
                ext = os.path.splitext(p)[1]
                base_new = format_template(self.var_template.get(), title, season_final, ep, ext, os.path.basename(p))
                # optionally append episode title if requested and available
                if include_title and extra_auto:
                    # insert extra before extension
                    b, e = os.path.splitext(base_new)
                    safe_extra = re.sub(r'[\\/:*?"<>|]+', '', extra_auto).strip('. ')
                    if safe_extra:
                        base_new = f"{b}.{safe_extra}{e}"
                # padding fix
                base_new = self._apply_padding_to_template(base_new, pad)
                newname = os.path.basename(base_new)
                self.preview_list.append((p, os.path.join(os.path.dirname(p), newname), os.path.basename(p), newname))
        elif pure_numeric_files and len(pure_numeric_files) >= max(2, len(files)//3):
            # treat pure numeric as sequence, sort by numeric value
            files_sorted_num = sorted(pure_numeric_files, key=lambda p: int(os.path.splitext(os.path.basename(p))[0]))
            for idx, p in enumerate(files_sorted_num, start=1):
                season_final = season_val if season_val is not None else 1
                ep = idx + offset
                ext = os.path.splitext(p)[1]
                base_new = format_template(self.var_template.get(), title, season_final, ep, ext, os.path.basename(p))
                base_new = self._apply_padding_to_template(base_new, pad)
                newname = os.path.basename(base_new)
                self.preview_list.append((p, os.path.join(os.path.dirname(p), newname), os.path.basename(p), newname))
            # also include remaining non-pure-numeric files after these (keep order)
            other_files = [p for p, s, e, ex in parsed if p not in set(pure_numeric_files)]
            for p in other_files:
                season_final = season_val if season_val is not None else 1
                # fallback sequential numbering for remaining
                ep = len(files_sorted_num) + 1 + offset
                ext = os.path.splitext(p)[1]
                base_new = format_template(self.var_template.get(), title, season_final, ep, ext, os.path.basename(p))
                base_new = self._apply_padding_to_template(base_new, pad)
                newname = os.path.basename(base_new)
                self.preview_list.append((p, os.path.join(os.path.dirname(p), newname), os.path.basename(p), newname))
        else:
            # fallback sequential assignment by current order
            for idx, (p, s_auto, e_auto, extra_auto) in enumerate(parsed, start=1):
                season_final = season_val if season_val is not None else (s_auto if s_auto is not None else 1)
                ep = (e_auto if e_auto is not None else idx) + offset
                ext = os.path.splitext(p)[1]
                base_new = format_template(self.var_template.get(), title, season_final, ep, ext, os.path.basename(p))
                if include_title and extra_auto:
                    b, e = os.path.splitext(base_new)
                    safe_extra = re.sub(r'[\\/:*?"<>|]+', '', extra_auto).strip('. ')
                    if safe_extra:
                        base_new = f"{b}.{safe_extra}{e}"
                base_new = self._apply_padding_to_template(base_new, pad)
                newname = os.path.basename(base_new)
                self.preview_list.append((p, os.path.join(os.path.dirname(p), newname), os.path.basename(p), newname))

        # detect conflicts (duplicate target names inside same folder)
        conflict_set = set()
        seen_targets = {}
        for oldabs, newabs, oldname, newname in self.preview_list:
            parent = os.path.dirname(oldabs)
            key = (parent, newname.lower())
            if key in seen_targets:
                conflict_set.add(key)
            else:
                seen_targets[key] = oldabs

        # insert into tree with tags; measure widths (font fallback handled)
        try:
            fnt = font.nametofont(self.tree.cget("font"))
        except Exception:
            fnt = font.Font(family="TkDefaultFont", size=10)
        max_old_w = 50
        max_new_w = 100

        for oldabs, newabs, oldname, newname in self.preview_list:
            tag = ""
            parent = os.path.dirname(oldabs)
            key = (parent, newname.lower())
            if key in conflict_set:
                tag = "conflict"
            elif oldname == newname:
                tag = "unchanged"
            else:
                tag = ""

            ext = os.path.splitext(newname)[1].lower()
            ext_tag = f"ext_{ext}"
            if ext not in self.ext_colors:
                self.ext_colors[ext] = "#000000"
            if ext_tag not in self.configured_tags:
                try:
                    self.tree.tag_configure(ext_tag, foreground=self.ext_colors.get(ext, "#000000"))
                    self.configured_tags.add(ext_tag)
                except Exception:
                    pass

            tags = []
            if tag:
                tags.append(tag)
            tags.append(ext_tag)

            self.tree.insert("", tk.END, values=(oldname, newname), tags=tags)
            w_old = fnt.measure(oldname) + 20
            w_new = fnt.measure(newname) + 20
            if w_old > max_old_w: max_old_w = w_old
            if w_new > max_new_w: max_new_w = w_new

        total = max(self.root.winfo_width() - 60, 200)
        col_old = min(max_old_w, int(total * 0.45))
        col_new = min(max_new_w, int(total * 0.55))
        self.tree.column("oldname", width=col_old)
        self.tree.column("newname", width=col_new)

        self.status_set(f"预览已生成: {len(self.preview_list)} 项 (红=冲突，灰=不变)")
        return

    def _apply_padding_to_template(self, newbase, pad):
        m = re.search(r'[eE](\d{1,})', newbase)
        if m:
            num = m.group(1)
            if len(num) < pad:
                newbase = newbase.replace(num, num.zfill(pad), 1)
        return newbase

    # ========== execute (线程) ==========
    def execute_async(self):
        if not self.preview_list:
            messagebox.showwarning("提示", "请先生成预览或扫描文件夹")
            return
        if self.var_dryrun.get():
            messagebox.showinfo("Dry-run", "当前为仅预览模式（Dry-run）。取消勾选后再次执行以实际重命名。")
            return
        if not messagebox.askyesno("确认", "确认执行重命名？此操作会修改文件名并不可撤销（可通过日志回滚）。建议先备份。"):
            return
        self.status_set("正在执行重命名（后台线程）...")
        EXECUTOR.submit(self._execute_task)

    def _execute_task(self):
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        # 日志输出到脚本同目录下的 logs/ 文件夹
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            script_dir = os.getcwd()
        log_dir = os.path.join(script_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"剧集自动重命名_log_{timestamp}.csv")
        log_rows = []
        conflict_suffix = self.var_conflict_suffix.get()
        processed = 0
        for oldabs, newabs, oldname, newname in list(self.preview_list):
            try:
                target_dir = os.path.dirname(newabs)
                if self.var_move_season_folder.get():
                    season_match = re.search(r'[sS](\d{1,2})', newname)
                    season_to_use = int(season_match.group(1)) if season_match else int(self.var_season.get() or 1)
                    season_folder = f"S{int(season_to_use):02}"
                    target_dir = os.path.join(os.path.dirname(oldabs), season_folder)
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir, exist_ok=True)
                    target = os.path.join(target_dir, newname)
                else:
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir, exist_ok=True)
                    target = newabs

                if os.path.exists(target):
                    base, ext = os.path.splitext(target)
                    t = f"{base}{conflict_suffix}{ext}"
                    i = 1
                    while os.path.exists(t):
                        t = f"{base}{conflict_suffix}{i}{ext}"
                        i += 1
                    target = t

                shutil.move(oldabs, target)
                log_rows.append((oldabs, target))
                processed += 1
            except Exception as ex:
                print("执行重命名失败:", oldabs, ex)
                continue

        try:
            with open(log_file, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["old", "new"])
                for r in log_rows:
                    w.writerow(r)
            self.root.after(0, lambda: self._on_execute_done(processed, log_file))
        except Exception as ex:
            self.root.after(0, lambda: messagebox.showwarning("日志写入失败", str(ex)))

    def _on_execute_done(self, processed, log_file):
        messagebox.showinfo("完成", f"重命名完成，已处理 {processed} 项。\n日志：{log_file}")
        self.status_set(f"已重命名 {processed} 项，日志：{log_file}")
        self.scan_async()

    # ========== Undo ==========
    def undo_via_log(self):
        csv_path = filedialog.askopenfilename(title="选择要回滚的日志 CSV 文件", filetypes=[("CSV 文件","*.csv"),("所有文件","*.*")])
        if not csv_path:
            return
        if not messagebox.askyesno("确认回滚", "将依据日志文件尝试将文件移动回原位，可能会覆盖现有文件。确定执行？"):
            return
        self.status_set("开始回滚（后台）...")
        EXECUTOR.submit(self._undo_task, csv_path)

    def _undo_task(self, csv_path):
        rollback_ts = time.strftime("%Y%m%d_%H%M%S")
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            script_dir = os.getcwd()
        log_dir = os.path.join(script_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        rollback_log = os.path.join(log_dir, f"剧集自动重命名_rollback_{rollback_ts}.csv")
        done = 0
        rows_out = []
        try:
            with open(csv_path, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader, None)
                for r in reader:
                    if not r: continue
                    old, new = (r[0], r[1] if len(r) > 1 else None)
                    source = new if new and os.path.exists(new) else (old if old and os.path.exists(old) else None)
                    if source is None:
                        cand_new = os.path.abspath(new) if new else None
                        cand_old = os.path.abspath(old) if old else None
                        if cand_new and os.path.exists(cand_new):
                            source = cand_new
                        elif cand_old and os.path.exists(cand_old):
                            source = cand_old
                    if source is None:
                        continue
                    target = old if os.path.isabs(old) else os.path.abspath(old)
                    tdir = os.path.dirname(target)
                    if not os.path.exists(tdir):
                        os.makedirs(tdir, exist_ok=True)
                    try:
                        shutil.move(source, target)
                        rows_out.append((source, target))
                        done += 1
                    except Exception as ex:
                        print("回滚移动失败", source, target, ex)
                        continue
            with open(rollback_log, "w", newline="", encoding="utf-8") as rf:
                w = csv.writer(rf)
                w.writerow(["moved_from", "moved_to"])
                for rr in rows_out:
                    w.writerow(rr)
            self.root.after(0, lambda: self._on_undo_done(done, rollback_log))
        except Exception as ex:
            self.root.after(0, lambda: messagebox.showwarning("回滚错误", str(ex)))

    def _on_undo_done(self, done, rollback_log):
        messagebox.showinfo("回滚完成", f"回滚完成（已处理 {done} 项）。回滚日志：{rollback_log}")
        self.status_set(f"回滚完成 {done} 项，回滚日志：{rollback_log}")
        self.scan_async()

    # ========== helpers ==========
    def _on_tree_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = sel[0]
        vals = self.tree.item(idx, "values")
        if not vals:
            return
        oldname = vals[0]
        for oldabs, newabs, oldn, newn in self.preview_list:
            if oldn == oldname:
                parent = os.path.dirname(oldabs)
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(parent)
                    elif sys.platform.startswith("darwin"):
                        os.system(f'open "{parent}"')
                    else:
                        os.system(f'xdg-open "{parent}"')
                except Exception:
                    pass
                break

    def status_set(self, text):
        try:
            self.status_label.config(text=text)
        except Exception:
            pass

    def show_help(self):
        txt = (
            "增强版使用说明：\n\n"
            "• 可将文件夹直接拖拽到程序窗口（需安装 tkinterdnd2），或点击“添加文件夹”。\n"
            "• 预览仅显示 原文件名 → 新文件名（红=冲突，灰=不变）。\n"
            "• 常用文件夹会自动保存到配置，程序下次启动会自动加载。\n"
            "• 扫描与执行在后台线程运行，不会阻塞 UI。\n"
            "• 先用 Dry-run 预览，确认无误后取消 Dry-run 并点击“重命名 (执行)”。\n"
            "• 如需回滚，点击“撤销（Undo）”并选择之前生成的日志 CSV。\n\n"
            "提示：若要启用拖拽请安装依赖：pip install tkinterdnd2"
        )
        messagebox.showinfo("帮助 / 关于", txt)

# ========== 启动 ==========
def main():
    if USE_TKINTERDND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = SeriesRenamerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
