# -*- coding: utf-8 -*-
"""对话框（ttkbootstrap 版）：连接设置 + 范围选择（用户过滤器树）。

连接设置：数据库账号密码仅从 S3D 认证 ini 读取（注册表记录路径，可手动
指定），不支持手动输入；保存前先测试连接（后台线程 + queue 回主线程）。
范围选择：过滤器存于 MDB，树型呈现；含系统层次条件的过滤器解析出系统
节点 oid；仅对象类型条件/无条件的过滤器列出但置灰不可选。
"""
import getpass
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk

import db
import excel_out
import mat_core
import snapshots

POLL_MS = 80


def set_app_icon(widget):
    """窗口高清图标：从 64.ico 取 32/48 档直接喂 PhotoImage。

    Tk 的 iconbitmap 加载 ico 内嵌档位会拉伸发糊；iconphoto 多档且带
    default=True 实测不显示——仅「双档 32/48 + False」清晰可靠（32 管
    标题栏/任务栏，48 管 Alt-Tab）。PhotoImage 挂 widget 属性防 GC。
    """
    try:
        from PIL import Image, ImageTk
        ico = Path(__file__).resolve().parent / "64.ico"
        if not ico.exists():
            return
        im = Image.open(ico)
        sizes = im.info.get("sizes") or []
        picks = [s for s in ((32, 32), (48, 48)) if s in sizes]
        frames = [im.ico.getimage(s).convert("RGBA") for s in picks]
        if not frames:
            frames = [im.convert("RGBA")]
        widget._icons = [ImageTk.PhotoImage(f, master=widget) for f in frames]
        widget.iconphoto(False, *widget._icons)
    except Exception:
        pass


# ------------------------------------------------------------ 连接设置 ----

def open_conn_dialog(parent, cfg: dict):
    """模态连接设置；保存成功返回更新后的 cfg，取消返回 None。"""
    dlg = _ConnDialog(parent, cfg)
    parent.wait_window(dlg)
    return dlg.result


class _ConnDialog(ttk.Toplevel):
    """凭据来源优先级：注册表认证 ini → 上次手动选择的 ini（cfg["auth_ini"]）。"""

    def __init__(self, parent, cfg: dict):
        super().__init__(parent)
        self.title("数据库连接设置")
        set_app_icon(self)
        self.transient(parent)
        self.resizable(False, False)
        self.cfg = dict(cfg)
        self.result = None       # 保存成功时为更新后的 cfg
        self.cred = None         # (ini_server, user, password)
        self.cred_source = ""
        self._saving = False     # 保存触发的测试：通过才写盘，失败不保存不关闭
        self._aborted = False    # 对话框已取消：测试线程的迟到结果不再处理
        self._q = queue.Queue()

        registry = db.read_site_connection_from_registry()
        site = registry.get("site") or {}
        schema = registry.get("schema") or {}
        site_db = cfg.get("site_db") or site.get("db", "")

        body = ttk.Frame(self, padding=(16, 12, 16, 12))
        body.pack(fill="both", expand=True)
        # 服务器只能从下拉选择（选项仅来自认证 ini 条目）：
        # 密码只存在于 ini，手输服务器配不到凭据；锁定选择可杜绝服务器与账号错配
        ttk.Label(body, text="服务器:").grid(row=0, column=0, sticky="e", pady=4, padx=(0, 8))
        self.cmb_server = ttk.Combobox(body, state="readonly", width=36)
        self.cmb_server.grid(row=0, column=1, sticky="we", pady=4)
        ttk.Label(body, text="Site 数据库:").grid(row=1, column=0, sticky="e", pady=4, padx=(0, 8))
        self.ed_site_db = ttk.Entry(body, width=38)
        self.ed_site_db.insert(0, site_db)
        self.ed_site_db.grid(row=1, column=1, sticky="we", pady=4)
        ttk.Label(body, text="Schema 库:").grid(row=2, column=0, sticky="e", pady=4, padx=(0, 8))
        self.lbl_schema = ttk.Label(body, text=schema.get("db", "") or "（注册表未读到 Schema 库）",
                                    foreground="#6b7280")
        self.lbl_schema.grid(row=2, column=1, sticky="w", pady=4)
        ttk.Label(body, text="数据库账号:").grid(row=3, column=0, sticky="e", pady=4, padx=(0, 8))
        self.lbl_cred = tk.Label(body, anchor="w", justify="left", wraplength=340)
        self.lbl_cred.grid(row=3, column=1, sticky="we", pady=4)
        body.columnconfigure(1, weight=1)

        note = ttk.Label(body, text="账号密码由 S3D 认证文件（ini）自动提供，无法手动输入。",
                         foreground="#6b7280", wraplength=380, justify="left")
        note.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))

        hb = ttk.Frame(self, padding=(16, 0, 16, 8))
        hb.pack(fill="x")
        self.btn_test = ttk.Button(hb, text="测试连接", width=12, command=self._on_test)
        btn_ini = ttk.Button(hb, text="选择认证 ini 文件...", width=18, command=self._on_pick_ini)
        btn_diag = ttk.Button(hb, text="诊断信息", width=10,
                              command=lambda: messagebox.showinfo(
                                  "连接诊断信息", db.collect_diagnostics(), parent=self))
        for b in (self.btn_test, btn_ini, btn_diag):
            b.configure(bootstyle="secondary-outline")
            b.pack(side="left", padx=(0, 8))

        bb = ttk.Frame(self, padding=(16, 4, 16, 12))
        bb.pack(fill="x")
        self.btn_ok = ttk.Button(bb, text="保存", width=10, bootstyle="primary",
                                 command=self._on_accept)
        btn_cancel = ttk.Button(bb, text="取消", width=10, command=self._cancel)
        btn_cancel.configure(bootstyle="secondary-outline")
        self.btn_ok.pack(side="right")
        btn_cancel.pack(side="right", padx=(0, 8))

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.cmb_server.bind("<<ComboboxSelected>>", lambda _e: self._match_cred())
        self._populate_servers()
        self._match_cred()
        self.after(POLL_MS, self._poll)
        # 窗口可见后再抓取输入焦点（未 viewable 时 grab_set 抛 TclError）
        self.wait_visibility()
        self.grab_set()

    # ---- 凭据 ----
    def _poll(self):
        if not self.winfo_exists():  # 已销毁：不再消费队列/触碰控件
            return
        try:
            while True:
                self._on_test_result(self._q.get_nowait())
        except queue.Empty:
            pass
        self.after(POLL_MS, self._poll)

    def _cancel(self):
        # 取消时测试线程可能仍在跑：置标志丢弃迟到结果
        self._aborted = True
        self.result = None
        self.destroy()

    def _populate_servers(self):
        """服务器下拉选项仅来自认证 ini 条目（注册表 ini + 手动 ini）。

        注册表 Site 服务器不进选项——走到连接设置多半是自动连接已失败，
        注册表环境不可信；仅当它也存在于 ini 条目中时参与默认选中。
        预选优先级：上次配置的服务器 → 注册表 Site 服务器 → 第一项。
        """
        servers = []

        def _add(s):
            s = (s or "").strip()
            if s and s not in servers:
                servers.append(s)

        prov = (self.cfg.get("provider") or "MSSQL").upper()
        for path in (db.read_auth_ini_path_from_registry(prov), self.cfg.get("auth_ini")):
            for e in db.parse_auth_ini(path or ""):
                if e.get("provider", "MSSQL") == prov:
                    _add(e["server"])
        reg = db.read_site_connection_from_registry().get("site") or {}

        self.cmb_server.configure(values=servers)
        for target in ((self.cfg.get("server") or "").strip(),
                       (reg.get("server") or "").strip()):
            idx = self._server_index(target)
            if idx >= 0:
                self.cmb_server.current(idx)
                return
        if servers:
            self.cmb_server.current(0)

    def _server_index(self, target: str) -> int:
        if not target:
            return -1
        for i, s in enumerate(self.cmb_server.cget("values")):
            if s == target or db._server_matches(target, s):
                return i
        return -1

    def _match_cred(self):
        """按当前所选服务器在 ini 中匹配凭据（手动 ini 优先，其次注册表 ini）。"""
        target = self.cmb_server.get().strip()
        prov = (self.cfg.get("provider") or "MSSQL").upper()
        sources = []
        if self.cfg.get("auth_ini"):
            sources.append(self.cfg["auth_ini"])
        reg_ini = db.read_auth_ini_path_from_registry(prov)
        if reg_ini and reg_ini not in sources:
            sources.append(reg_ini)
        for source in sources:
            cred = db.get_credentials_from_ini(source, target, prov)
            if cred:
                self._set_cred(cred, source)
                return
        self._set_cred(None, "")

    def _set_cred(self, cred, source):
        self.cred = cred
        self.cred_source = source or ""
        if cred:
            name = Path(self.cred_source).name if self.cred_source else ""
            self.lbl_cred.configure(
                text=f"{cred[1]}（来自 {name}）" if name else cred[1], fg="#1a7f37")
        else:
            self.lbl_cred.configure(
                text="所选服务器在认证 ini 中无账号条目，请点击「选择认证 ini 文件...」",
                fg="#b54708")

    def _on_pick_ini(self):
        """手动指定认证 ini：其服务器条目并入下拉，并优先采用该 ini 的凭据。"""
        start = self.cfg.get("auth_ini") or ""
        path = filedialog.askopenfilename(
            parent=self, title="选择认证 ini 文件", initialdir=start or ".",
            filetypes=[("ini 文件", "*.ini"), ("所有文件", "*.*")])
        if not path:
            return
        prov = (self.cfg.get("provider") or "MSSQL").upper()
        entries = db.parse_auth_ini(path)
        cred = db.get_credentials_from_ini(path, self.cmb_server.get().strip(), prov)
        if not cred and entries:
            # 当前服务器在该 ini 无条目 → 预选 ini 第一个有账号的服务器
            for e in entries:
                cred = db.get_credentials_from_ini(path, e["server"], prov)
                if cred:
                    self.cfg["auth_ini"] = path
                    self._populate_servers()
                    for i, s in enumerate(self.cmb_server.cget("values")):
                        if db._server_matches(e["server"], s):
                            self.cmb_server.current(i)
                            break
                    self._set_cred(cred, path)
                    return
        if not cred:
            messagebox.showwarning(
                "解析失败",
                "无法从该文件解析出账号密码。\n"
                "需要 ServicesAuthenticationSettings.exe 维护的格式："
                "'服务器 账号 密文'（空格分隔三列）。", parent=self)
            return
        self.cfg["auth_ini"] = path
        self._populate_servers()
        self._set_cred(cred, path)

    # ---- 测试 / 保存 ----
    def _on_test(self):
        if not self.cred:
            messagebox.showwarning("提示", "未找到认证文件，请先选择认证 ini 文件。",
                                   parent=self)
            return
        self._run_test()

    def _run_test(self):
        """GUI 线程内先取快照，子线程只使用快照值（不触碰控件）。"""
        server = self.cmb_server.get().strip()
        site_db = self.ed_site_db.get().strip()
        user, pwd = self.cred[1], self.cred[2]
        prov = (self.cfg.get("provider") or "MSSQL").upper()
        self.btn_test.configure(state="disabled", text="测试中...")
        self.btn_ok.configure(state="disabled")
        q = self._q

        def work():
            q.put(db.test_connection(server, site_db, user, pwd, prov))
        threading.Thread(target=work, daemon=True).start()

    def _on_test_result(self, err: str):
        if self._aborted:
            return  # 对话框已取消：丢弃迟到结果（不恢复按钮/不弹窗/不写盘）
        self.btn_test.configure(state="normal", text="测试连接")
        self.btn_ok.configure(state="normal")
        if self._saving:
            self._saving = False
            if err:
                messagebox.showwarning(
                    "无法保存", "连接测试未通过，配置未保存：\n" + err, parent=self)
                return
            self._do_save()
        elif err:
            messagebox.showwarning("连接失败", err, parent=self)
        else:
            messagebox.showinfo("成功", "连接成功！", parent=self)

    def _on_accept(self):
        if not self.cred:
            messagebox.showwarning(
                "无法保存",
                "未找到有效的认证文件（ini），无法获取数据库账号密码。\n"
                "请点击「选择认证 ini 文件...」指定后再保存。", parent=self)
            return
        # 保存前先验证连接：config.json 里记录的必须是连得上的配置，
        # 杜绝错误服务器记录被「上次配置优先」沿用导致自动连接失败
        self._saving = True
        self._run_test()

    def _do_save(self):
        self.cfg["server"] = self.cmb_server.get().strip()
        self.cfg["site_db"] = self.ed_site_db.get().strip()
        self.cfg["username"] = self.cred[1]
        self.cfg["password_enc"] = db.encrypt_password(self.cred[2])
        db.save_config(self.cfg)
        self.result = self.cfg
        self.destroy()


# ------------------------------------------------------------ 范围选择 ----

def open_filter_dialog(parent, conn_info: dict, plant: dict,
                       plant_name: str = "", last_name: str = ""):
    """模态过滤器选择；确定返回 (name, oids)，取消返回 None。"""
    dlg = _FilterDialog(parent, conn_info, plant, plant_name, last_name)
    parent.wait_window(dlg)
    return dlg.result


# ------------------------------------------------------------ 图纸目录 ----

def open_dwg_folder_dialog(parent, conn_info: dict, plant: dict,
                           plant_name: str = "", last_names=None):
    """模态图纸目录树选择（ISO 图纸报表专用）；确定返回 (names, oids)，
    取消返回 None；无勾选确定 = 不过滤（全部图纸）。"""
    dlg = _DwgFolderDialog(parent, conn_info, plant, plant_name,
                           last_names or [])
    parent.wait_window(dlg)
    return dlg.result


class _DwgFolderDialog(ttk.Toplevel):
    """图纸目录树（DrawingFolderHierarchy）复选框多选：单击行切换，
    父勾选传播全子树；搜索框过滤。SP3D FilterSelector 选不到该树。"""

    def __init__(self, parent, conn_info, plant, plant_name, last_names):
        super().__init__(parent)
        self.title(f"选择图纸目录 - {plant_name}")
        set_app_icon(self)
        self.transient(parent)
        self.geometry("560x680")
        self.minsize(460, 520)
        self.result = None       # (names, oids) 或 None
        self._checked = set()    # 勾选节点 oid
        self._name_of = {}       # iid -> (oid, name)
        self._parent_of = {}     # iid -> parent iid
        self._loading = True
        self._last_names = list(last_names)
        self._q = queue.Queue()

        body = ttk.Frame(self, padding=(16, 12, 16, 12))
        body.pack(fill="both", expand=True)
        self.lbl_tip = ttk.Label(body, text="正在加载图纸目录...", foreground="#175CD3")
        self.lbl_tip.pack(anchor="w")
        self.ed_search = ttk.Entry(body, state="disabled")
        self.ed_search.configure(bootstyle="secondary")
        self.ed_search.pack(fill="x", pady=(8, 6))

        box = ttk.Frame(body)
        box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(box, show="tree", selectmode="none")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.lbl_detail = ttk.Label(body, text="单击行勾选/取消（勾选父目录即含全部子级）；不勾选 = 全部图纸",
                                    foreground="#6b7280", wraplength=520,
                                    justify="left")
        self.lbl_detail.pack(anchor="w", fill="x", pady=(6, 0))

        bb = ttk.Frame(body)
        bb.pack(fill="x", pady=(10, 0))
        self.btn_ok = ttk.Button(bb, text="确定", width=10, bootstyle="primary",
                                 command=self._on_ok)
        self.btn_clear = ttk.Button(bb, text="全部清除", width=10,
                                    command=self._on_clear)
        self.btn_clear.configure(bootstyle="secondary-outline")
        self.btn_cancel = ttk.Button(bb, text="取消", width=10,
                                     command=self._cancel)
        self.btn_cancel.configure(bootstyle="secondary-outline")
        # 加载期间禁用全部按钮（加载线程仍持有数据库连接）
        for b in (self.btn_ok, self.btn_clear, self.btn_cancel):
            b.configure(state="disabled")
        self.btn_ok.pack(side="right")
        self.btn_cancel.pack(side="right", padx=(0, 8))
        self.btn_clear.pack(side="right", padx=(0, 8))

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.tree.bind("<Button-1>", self._on_click)
        self.ed_search.bind("<KeyRelease>", lambda _e: self._on_search())
        self.after(POLL_MS, self._poll)

        db_r = db.open_db(conn_info.get("provider", "MSSQL"),
                          conn_info["server"], plant["reports"],
                          conn_info["username"], conn_info["password"])
        self._db = db_r
        q = self._q

        def work():
            try:
                q.put(("ok", mat_core.load_drawing_tree(db_r, plant)))
            except Exception as ex:
                q.put(("fail", str(ex)))
        threading.Thread(target=work, daemon=True).start()
        self.wait_visibility()
        self.grab_set()

    # ---- 加载轮询 / 关闭 ----
    def _poll(self):
        if not self.winfo_exists():
            return
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "ok":
                    self._on_loaded(payload)
                else:
                    self._on_failed(payload)
        except queue.Empty:
            pass
        self.after(POLL_MS, self._poll)

    def _cancel(self):
        if self._loading:
            return
        self.result = None
        self._close_db()
        self.destroy()

    def _close_db(self):
        d, self._db = getattr(self, "_db", None), None
        if d is not None:
            try:
                d.close()
            except Exception:
                pass

    def _on_loaded(self, result):
        self._loading = False
        self.lbl_tip.pack_forget()
        for b in (self.btn_ok, self.btn_clear, self.btn_cancel):
            b.configure(state="normal")
        self.ed_search.configure(state="normal")
        roots, _edges = result
        self._close_db()
        if not roots:
            messagebox.showwarning("提示", "该 Plant 没有图纸目录树。", parent=self)
            self.destroy()
            return
        for r in roots:
            self._populate(r, "")
        # 按名称恢复上次勾选
        last = set(self._last_names)
        if last:
            for iid, (_oid, name) in self._name_of.items():
                if name in last:
                    self._set_check(iid, True, propagate=True)

    def _on_failed(self, msg):
        self._loading = False
        self.lbl_tip.pack_forget()
        self._close_db()
        messagebox.showwarning("加载图纸目录失败", msg, parent=self)
        self.destroy()

    def _populate(self, node: dict, parent_iid: str):
        it = self.tree.insert(parent_iid, "end", text=f"☐ {node['name']}")
        self._name_of[it] = (node["oid"], node["name"])
        self._parent_of[it] = parent_iid
        for ch in node["children"]:
            self._populate(ch, it)

    # ---- 勾选 ----
    def _on_click(self, e):
        if self._loading:
            return
        # 点在展开/折叠箭头上不切换勾选（否则打开目录会翻转整棵子树）
        try:
            if "indicator" in str(self.tree.identify("element", e.x, e.y)):
                return
        except tk.TclError:
            pass
        iid = self.tree.identify_row(e.y)
        if not iid or iid not in self._name_of:
            return
        oid = self._name_of[iid][0]
        self._set_check(iid, oid not in self._checked, propagate=True)

    def _set_check(self, iid, on: bool, propagate=False):
        oid = self._name_of[iid][0]
        if on:
            self._checked.add(oid)
        else:
            self._checked.discard(oid)
        self.tree.item(iid, text=("☑ " if on else "☐ ")
                       + self._name_of[iid][1])
        if propagate:
            for ch in self.tree.get_children(iid):
                self._set_check(ch, on, propagate=True)
            self._propagate_up(iid)
        self._update_detail()

    def _propagate_up(self, iid):
        # 向上联动：子级全勾 → 父级自动勾上（递归到根）；
        # 子级出现未勾 → 已勾的祖先逐级取消（其余子级保持原状）
        p = self._parent_of.get(iid)
        if not p or p not in self._name_of:
            return
        kids = self.tree.get_children(p)
        all_on = bool(kids) and all(self._name_of[k][0] in self._checked
                                    for k in kids)
        poid = self._name_of[p][0]
        if all_on and poid not in self._checked:
            self._checked.add(poid)
            self.tree.item(p, text="☑ " + self._name_of[p][1])
            self._propagate_up(p)
        elif not all_on and poid in self._checked:
            self._checked.discard(poid)
            self.tree.item(p, text="☐ " + self._name_of[p][1])
            self._propagate_up(p)

    def _update_detail(self):
        if self._checked:
            self.lbl_detail.configure(
                text=f"已勾选 {len(self._checked)} 个目录节点（含子级）",
                foreground="#1a7f37")
        else:
            self.lbl_detail.configure(
                text="未勾选（全部图纸）", foreground="#6b7280")

    def _on_clear(self):
        for iid in self._name_of:
            self._set_check(iid, False, propagate=False)

    # ---- 搜索 ----
    def _on_search(self):
        t = self.ed_search.get().strip().lower()

        def visit(iid) -> bool:
            text = self.tree.item(iid, "text").lower()
            keep = t in text
            for ch in self.tree.get_children(iid):
                if visit(ch):
                    keep = True
            if not keep:
                self.tree.detach(iid)
            elif t:
                self.tree.item(iid, open=True)
            return keep

        # 先恢复原位再过滤（同 _FilterDialog）
        for iid, parent in self._parent_of.items():
            try:
                self.tree.move(iid, parent, "end")
            except tk.TclError:
                pass
        if t:
            for top in list(self.tree.get_children("")):
                visit(top)
        else:
            for top in self.tree.get_children(""):
                self._collapse_all(top)

    def _collapse_all(self, iid):
        self.tree.item(iid, open=False)
        for ch in self.tree.get_children(iid):
            self._collapse_all(ch)

    # ---- 确定 ----
    def _on_ok(self):
        # 只返回勾选中最高层的节点（其祖先未被勾选），避免 oid 冗余
        oid_parent = {v[0]: self._name_of.get(self._parent_of[iid], (None,))[0]
                      for iid, v in self._name_of.items()}
        top_oids, names = [], []
        for oid in self._checked:
            p = oid_parent.get(oid)
            if p is None or p not in self._checked:
                top_oids.append(oid)
        name_by_oid = {v[0]: v[1] for v in self._name_of.values()}
        names = [name_by_oid.get(o, "") for o in top_oids]
        self.result = (names, top_oids)
        self.destroy()


class _FilterDialog(ttk.Toplevel):
    """过滤器树：文件夹仅分组；含系统 oid 的可选，其余置灰。"""

    def __init__(self, parent, conn_info, plant, plant_name, last_name):
        super().__init__(parent)
        self.title(f"选择过滤器 - {plant_name}")
        set_app_icon(self)
        self.transient(parent)
        self.geometry("540x660")
        self.minsize(460, 520)
        self.result = None       # (name, oids) 或 None
        self._conn_info = conn_info   # SQL 过滤器确定时开 RDB 执行
        self._plant = plant
        self._data = {}          # iid -> 节点 dict（仅可选过滤器节点）
        self._order = {}         # iid -> (parent_iid, index)：搜索过滤后恢复原位
        self._loading = True
        self._last_name = last_name
        self._q = queue.Queue()

        body = ttk.Frame(self, padding=(16, 12, 16, 12))
        body.pack(fill="both", expand=True)
        self.lbl_tip = ttk.Label(body, text="正在加载过滤器...", foreground="#175CD3")
        self.lbl_tip.pack(anchor="w")
        self.ed_search = ttk.Entry(body, state="disabled")
        self.ed_search.configure(bootstyle="secondary")
        self.ed_search.pack(fill="x", pady=(8, 6))

        box = ttk.Frame(body)
        box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(box, show="tree", selectmode="browse")
        self.tree.tag_configure("disabled", foreground="#9ca3af")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.lbl_detail = ttk.Label(body, text="", foreground="#6b7280",
                                    wraplength=500, justify="left")
        self.lbl_detail.pack(anchor="w", fill="x", pady=(6, 0))

        bb = ttk.Frame(body)
        bb.pack(fill="x", pady=(10, 0))
        self.btn_ok = ttk.Button(bb, text="确定", width=10, bootstyle="primary",
                                 command=self._on_ok)
        self.btn_cancel = ttk.Button(bb, text="取消", width=10, command=self._cancel)
        self.btn_cancel.configure(bootstyle="secondary-outline")
        # 加载期间禁用全部按钮（加载线程仍持有数据库连接，不允许提前关闭）
        self.btn_ok.configure(state="disabled")
        self.btn_cancel.configure(state="disabled")
        self.btn_ok.pack(side="right")
        self.btn_cancel.pack(side="right", padx=(0, 8))

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._on_select())
        self.ed_search.bind("<KeyRelease>", lambda _e: self._on_search())
        self.after(POLL_MS, self._poll)

        # 后台解析用户过滤器（Db 惰连接，真正连库发生在本线程首次查询）
        db_mdb = db.open_db(conn_info.get("provider", "MSSQL"),
                            conn_info["server"], plant["model"],
                            conn_info["username"], conn_info["password"])
        self._db = db_mdb
        q = self._q

        def work():
            try:
                q.put(("ok", mat_core.load_user_filters(db_mdb, getpass.getuser())))
            except Exception as ex:
                q.put(("fail", str(ex)))
        threading.Thread(target=work, daemon=True).start()
        self.wait_visibility()
        self.grab_set()

    # ---- 加载轮询 ----
    def _poll(self):
        if not self.winfo_exists():  # 已销毁：不再消费队列/触碰控件
            return
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "ok":
                    self._on_loaded(payload)
                else:
                    self._on_failed(payload)
        except queue.Empty:
            pass
        self.after(POLL_MS, self._poll)

    def _cancel(self):
        # 加载中不允许关闭：后台线程还在使用数据库连接
        if self._loading:
            return
        self.result = None
        self._close_db()
        self.destroy()

    def _close_db(self):
        # 加载完成（线程已结束）后由主线程关闭连接
        d, self._db = getattr(self, "_db", None), None
        if d is not None:
            try:
                d.close()
            except Exception:
                pass

    def _on_loaded(self, result):
        self._loading = False
        self.lbl_tip.pack_forget()
        self.btn_cancel.configure(state="normal")
        self.ed_search.configure(state="normal")
        tree, my_tree = result
        if my_tree:
            # 尾部追加「My Filters」区块（folder 节点仅作分组标题不可选）
            tree = tree + [{"type": "folder", "name": "My Filters",
                            "children": my_tree}]
        if not tree:
            self._close_db()
            messagebox.showwarning(
                "提示",
                "该 Plant 没有可用过滤器\n（只列 Plant Filters / My Filters 区的过滤器）。",
                parent=self)
            self.destroy()
            return
        self._close_db()
        self._populate(tree, "")
        self._record_order()
        # 预选上次选择的过滤器（按原始名字匹配，不受消歧后缀影响），
        # 上次的已被删/改名时回退第一个可用
        first = self._find_by_name(self._last_name) if self._last_name else None
        first = first or self._first_usable()
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)
            self.tree.see(first)

    def _on_failed(self, msg):
        self._loading = False
        self.lbl_tip.pack_forget()
        self._close_db()
        messagebox.showwarning("加载过滤器失败", msg, parent=self)
        self.destroy()

    def _populate(self, nodes: list, parent_iid: str):
        """按 weld_core 返回的树结构填充；同文件夹重名项加注节点摘要。"""
        from collections import Counter
        dup = {name for name, k in
               Counter(n["name"] for n in nodes
                       if n["type"] == "filter").items() if k > 1}
        for n in nodes:
            if n["type"] == "folder":
                it = self.tree.insert(parent_iid, "end",
                                      text="📁 " + n["name"], open=False)
                self._populate(n["children"], it)
                continue
            label = n["name"]
            if n["name"] in dup and n["sys_names"]:
                shown = "、".join(n["sys_names"][:3])
                more = "…" if len(n["sys_names"]) > 3 else ""
                label = f"{n['name']}（{shown}{more}）"
            # 图标：📁 文件夹 / 🔍 普通过滤器 / 📝 SQL 过滤器
            label = ("📝 " if n["kind"] == "sql" else "🔍 ") + label
            if (n["kind"] == "sys" and n["oids"]) or n["kind"] == "sql":
                it = self.tree.insert(parent_iid, "end", text=label)
                self._data[it] = n
            else:
                it = self.tree.insert(parent_iid, "end", text=label,
                                      tags=("disabled",))

    def _record_order(self):
        """记录每个节点的原位置，供搜索过滤后恢复。"""
        def visit(iid):
            self._order[iid] = (self.tree.parent(iid), self.tree.index(iid))
            for ch in self.tree.get_children(iid):
                visit(ch)
        for top in self.tree.get_children(""):
            visit(top)

    def _find_by_name(self, name: str):
        """递归查找名字匹配的可选过滤器节点（同名取先出现的）。"""
        for iid, n in self._data.items():
            if n.get("name") == name:
                return iid
        return None

    def _first_usable(self):
        for iid in self._data:
            return iid
        return None

    def _on_search(self):
        t = self.ed_search.get().strip().lower()

        def visit(iid) -> bool:
            text = self.tree.item(iid, "text").lower()
            hit = t in text
            keep = hit
            for ch in self.tree.get_children(iid):
                if visit(ch):
                    keep = True
            if not keep:
                self.tree.detach(iid)
            elif t:
                self.tree.item(iid, open=True)
            return keep

        # 先恢复原位再过滤，避免上次过滤把父节点藏掉导致子节点丢失
        for iid, (parent, idx) in self._order.items():
            try:
                self.tree.move(iid, parent, idx)
            except tk.TclError:
                pass
        if t:
            for top in list(self.tree.get_children("")):
                visit(top)
        else:
            for top in self.tree.get_children(""):
                self._collapse_all(top)

    def _collapse_all(self, iid):
        self.tree.item(iid, open=False)
        for ch in self.tree.get_children(iid):
            self._collapse_all(ch)

    def _on_select(self):
        iid = self.tree.focus()
        f = self._data.get(iid)
        if not f:
            self.lbl_detail.configure(text="")
            if not self._loading:
                self.btn_ok.configure(state="disabled")
            return
        if f["kind"] == "sql":
            self.lbl_detail.configure(
                text="📝 SQL 过滤器：确定时在报表库执行其查询，"
                     "返回的对象直接作为统计范围（不经系统树展开）。")
            self.btn_ok.configure(state="normal")
            return
        names = f["sys_names"]
        shown = "、".join(names[:8]) or "（无名字记录）"
        more = f" 等共 {len(names)} 个" if len(names) > 8 else ""
        self.lbl_detail.configure(text=f"包含系统节点：{shown}{more}")
        self.btn_ok.configure(state="normal")

    def _exec_sql_filter(self, f: dict):
        """在 RDB 执行 SQL 过滤器查询，返回 oid 集合；失败弹窗返回 None。

        结果列取 ResultColumnName 指定列（无则第一列）；空集合返回空
        set（调用方拦截——空结果绝不能当作「未设置范围」回退全量）。
        """
        info = self._conn_info
        try:
            dbr = db.open_db(info.get("provider", "MSSQL"),
                             info["server"], self._plant["reports"],
                             info["username"], info["password"])
        except Exception as ex:
            messagebox.showerror("连接失败", "打开报表库失败：\n" + str(ex),
                                 parent=self)
            return None
        try:
            rows = dbr.query(f["sql_text"])
        except Exception as ex:
            messagebox.showwarning(
                "SQL 过滤器执行失败",
                "「%s」的查询执行出错：\n%s" % (f["name"], ex), parent=self)
            return None
        finally:
            dbr.close()
        idx = 0
        rc = f.get("sql_result_col") or ""
        if rc and rows:
            # db.query 行是普通 list（无列名元数据），拿不到列名时退第一列
            desc = getattr(rows[0], "cursor_description", None) or []
            cols = [str(c[0]).lower() for c in desc]
            if rc.lower() in cols:
                idx = cols.index(rc.lower())
        out = set()
        for r in rows:
            v = r[idx] if idx < len(r) else None
            if v is not None and str(v).strip():
                out.add(str(v))
        return out

    def _on_ok(self):
        iid = self.tree.focus()
        f = self._data.get(iid)
        if not f:
            return
        if f["kind"] == "sql":
            oids = self._exec_sql_filter(f)
            if oids is None:
                return          # 执行失败已提示，留在对话框
            if not oids:
                messagebox.showwarning(
                    "查询结果为空",
                    "「%s」的查询没有返回对象。\n"
                    "不设置范围（取消选择）。" % f["name"], parent=self)
                return
            self.result = (f["name"], sorted(oids))
            self.destroy()
            return
        self.result = (f["name"], list(f["oids"]))
        self.destroy()


# ------------------------------------------------------------ 快照管理 ----

def open_snapshot_dialog(parent, cfg: dict, flt: dict = None) -> None:
    """快照管理（模态）：列表 / 比较导出 / 单快照导出 / 删除。

    flt = {"report", "plant"} 时只列出同类型、同项目的快照
    （范围不限——扩大范围比增量是合理用法）。"""
    dlg = _SnapshotDialog(parent, cfg, flt)
    parent.wait_window(dlg)


class _SnapshotDialog(ttk.Toplevel):
    """快照列表（时间升序）：选 2 份同类型比较导出增加/减少两 sheet，
    选 1 份全量导出；快照自描述，全程不依赖数据库连接。"""

    _TIP = "选两份同类型快照「比较导出」增减差异；选一份可「导出快照」全量。"

    def __init__(self, parent, cfg: dict, flt: dict = None):
        super().__init__(parent)
        self.title("快照管理 - 存档与比较")
        set_app_icon(self)
        self.transient(parent)
        self.geometry("780x480")
        self.minsize(640, 380)
        self.cfg = cfg
        self._flt = flt or None  # 主界面当前设置过滤（类型/项目/范围）
        self._meta = {}          # iid -> 列表元数据（path/report/plant/time/scope/by_line）

        body = ttk.Frame(self, padding=(16, 12, 16, 12))
        body.pack(fill="both", expand=True)
        self.lbl_tip = ttk.Label(body, text=self._TIP, foreground="#6b7280")
        self.lbl_tip.pack(anchor="w")

        box = ttk.Frame(body)
        box.pack(fill="both", expand=True, pady=(8, 0))
        cols = ("report", "plant", "byline", "time", "scope")
        self.tree = ttk.Treeview(box, columns=cols, show="headings",
                                 selectmode="extended")
        for c, txt, w in (("report", "类型", 90), ("plant", "项目", 90),
                          ("byline", "记录方式", 110), ("time", "时间", 150),
                          ("scope", "范围", 250)):
            self.tree.heading(c, text=txt)
            self.tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        bb = ttk.Frame(body)
        bb.pack(fill="x", pady=(10, 0))
        btn_close = ttk.Button(bb, text="关闭", width=10,
                               bootstyle="secondary-outline",
                               command=self.destroy)
        btn_export = ttk.Button(bb, text="导出快照", width=12,
                                bootstyle="secondary-outline",
                                command=self._on_export)
        btn_cmp = ttk.Button(bb, text="比较导出", width=12,
                             bootstyle="primary", command=self._on_compare)
        btn_del = ttk.Button(bb, text="删除", width=10,
                             bootstyle="danger-outline", command=self._on_delete)
        btn_refresh = ttk.Button(bb, text="刷新", width=10,
                                 bootstyle="secondary-outline",
                                 command=self._refresh)
        btn_close.pack(side="right")
        btn_export.pack(side="right", padx=(0, 8))
        btn_cmp.pack(side="right", padx=(0, 8))
        btn_refresh.pack(side="left")
        btn_del.pack(side="left", padx=(8, 0))

        self._refresh()
        self.wait_visibility()
        self.grab_set()

    # ---- 列表 ----
    @staticmethod
    def _byline_text(m: dict) -> str:
        """记录方式列：产出快照时固化的统计口径。"""
        if m["report"] == "焊接量":
            w = m.get("weld_opts")
            if not w:
                return "未记录"
            parts = [("不分组", "按材料等级", "按材料组")[w.get("met", 0)],
                     "忽略壁厚" if w.get("thk") else "含壁厚"]
            if w.get("skip_pipes"):
                parts.append("忽略直管")
            return "；".join(parts)
        v = m.get("by_line")
        if v is True:
            return "按管线号"
        if v is False:
            return "不按管线号"
        return "未记录"

    def _refresh(self):
        self.tree.delete(*self.tree.get_children(""))
        self._meta.clear()
        snaps = snapshots.list_snapshots()
        if self._flt:
            snaps = [m for m in snaps
                     if m["report"] == self._flt["report"]
                     and m["plant"] == self._flt["plant"]]
            # 按主界面 GUI 当前口径过滤：勾了「按管线号分组」只列分组口径
            # 的快照，没勾只列不分组口径；焊接量按当前模式/壁厚/直管设置。
            # 同口径才可比，列表里天然只剩能互比的快照。
            if "by_line" in self._flt:
                snaps = [m for m in snaps
                         if m.get("by_line") == self._flt["by_line"]]
            if "weld_opts" in self._flt:
                snaps = [m for m in snaps
                         if m.get("weld_opts") == self._flt["weld_opts"]]
            base = "仅列出当前项目（{}）「{}」类型、当前口径（{}）的快照，范围不限".format(
                self._flt["plant"], self._flt["report"],
                self._byline_text({"report": self._flt["report"],
                                   "by_line": self._flt.get("by_line"),
                                   "weld_opts": self._flt.get("weld_opts")})
                if ("by_line" in self._flt or "weld_opts" in self._flt)
                else "全部")
        else:
            base = self._TIP
        for i, m in enumerate(snaps):
            iid = str(i)
            self.tree.insert("", "end", iid=iid,
                             values=(m["report"], m["plant"],
                                     self._byline_text(m), m["time"],
                                     m["scope"] or "—"))
            self._meta[iid] = m
        if snaps:
            self.lbl_tip.configure(text=base, foreground="#6b7280")
        else:
            self.lbl_tip.configure(
                text="当前设置下暂无快照。先在主界面统计后点「存为快照」。",
                foreground="#b54708")

    def _selected(self) -> list:
        return [self._meta[i] for i in self.tree.selection() if i in self._meta]

    # ---- 比较导出 ----
    def _on_compare(self):
        ms = self._selected()
        if len(ms) != 2:
            messagebox.showwarning("提示", "请选择两份快照进行比较。", parent=self)
            return
        if ms[0]["report"] != ms[1]["report"]:
            messagebox.showwarning("提示", "只能比较同类型快照。", parent=self)
            return
        a, b = sorted(ms, key=lambda m: m["time"])   # 时间早 = 旧快照
        if (a.get("by_line") != b.get("by_line")
                or a.get("weld_opts") != b.get("weld_opts")):
            messagebox.showwarning(
                "提示",
                "两份快照统计口径不同（{} / {}），不能比较。\n"
                "请选同口径的两份快照。".format(
                    self._byline_text(a), self._byline_text(b)),
                parent=self)
            return
        try:
            headers, added, removed = snapshots.compare(
                snapshots.load(a["path"]), snapshots.load(b["path"]))
        except Exception as ex:
            messagebox.showerror("比较失败", str(ex), parent=self)
            return
        default = "{}比较_{}_{}.xlsx".format(
            a["report"], a["time"][:10], b["time"][:10])
        path = filedialog.asksaveasfilename(
            parent=self, title="保存比较结果",
            initialdir=self.cfg.get("last_dir") or str(Path.home()),
            initialfile=default, defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            n_add, n_rem = excel_out.write_compare_excel(
                a["report"], a["time"], b["time"], headers, added, removed,
                path)
        except Exception as ex:
            messagebox.showerror("导出失败", str(ex), parent=self)
            return
        self.cfg["last_dir"] = str(Path(path).parent)
        db.save_config(self.cfg)
        messagebox.showinfo(
            "导出完成", f"增加 {n_add} 行，减少 {n_rem} 行。\n{path}",
            parent=self)

    # ---- 单快照导出 ----
    def _on_export(self):
        ms = self._selected()
        if len(ms) != 1:
            messagebox.showwarning("提示", "请选择一份快照导出。", parent=self)
            return
        m = ms[0]
        try:
            meta = snapshots.load(m["path"])
        except Exception as ex:
            messagebox.showerror("读取快照失败", str(ex), parent=self)
            return
        default = "{}_{}_{}.xlsx".format(
            m["report"], m["plant"], m["time"][:10])
        path = filedialog.asksaveasfilename(
            parent=self, title="导出快照",
            initialdir=self.cfg.get("last_dir") or str(Path.home()),
            initialfile=default, defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            n = excel_out.write_snapshot_excel(
                m["report"], meta["headers"], meta["rows"], path)
        except Exception as ex:
            messagebox.showerror("导出失败", str(ex), parent=self)
            return
        self.cfg["last_dir"] = str(Path(path).parent)
        db.save_config(self.cfg)
        messagebox.showinfo("导出完成", f"导出完成，共 {n} 行。\n{path}",
                            parent=self)

    # ---- 删除 ----
    def _on_delete(self):
        ms = self._selected()
        if not ms:
            return
        if not messagebox.askyesno("删除快照",
                                   f"确定删除选中的 {len(ms)} 份快照？",
                                   parent=self):
            return
        errs = []
        for m in ms:
            try:
                snapshots.delete(m["path"])
            except OSError as ex:
                errs.append(f"{Path(m['path']).name}：{ex}")
        self._refresh()
        if errs:
            messagebox.showerror(
                "删除失败",
                "以下文件未能删除（可能被其他程序占用）：\n" + "\n".join(errs),
                parent=self)
