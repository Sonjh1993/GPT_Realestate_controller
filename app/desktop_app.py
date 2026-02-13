from __future__ import annotations

import os
import platform
import shutil
import webbrowser
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .matching import match_properties
from .proposal import build_kakao_message, generate_proposal_pdf
from .tasks_engine import reconcile_auto_tasks
from .sheet_sync import SyncSettings, upload_visible_data
from .storage import (
    PROPERTY_TABS,
    TAB_COMPLEX_NAME,
    add_customer,
    add_photo,
    add_property,
    add_viewing,
    delete_photo,
    delete_viewing,
    get_customer,
    get_property,
    get_setting,
    init_db,
    list_customers,
    list_photos,
    list_photos_all,
    list_properties,
    list_viewings,
    set_setting,
    soft_delete_customer,
    soft_delete_property,
    toggle_customer_hidden,
    toggle_property_hidden,
    update_customer,
    update_property,
    update_viewing_status,
    add_task,
    delete_task,
    get_viewing,
    list_tasks,
    mark_task_done,
    get_task,
 )


# 단지별 면적타입 (드롭다운)
UNIT_TYPES_BY_TAB = {
    "아파트단지1": [
        "82A㎡",
        "82C㎡",
        "82A1㎡",
        "82C1㎡",
        "82B㎡",
        "82D㎡",
        "98C㎡",
        "98C1㎡",
        "98D㎡",
        "98A㎡",
        "98C2㎡",
        "99B㎡",
        "113B㎡",
        "114A㎡",
        "114C㎡",
        "145㎡",
    ],
    "아파트단지2": [
        "83A㎡",
        "84C㎡",
        "84D㎡",
        "84B-1㎡",
        "84B㎡",
        "100A㎡",
        "100B㎡",
        "100D㎡",
        "100C㎡",
        "101E㎡",
        "116D㎡",
        "116B㎡",
        "116E㎡",
        "116A㎡",
        "117C㎡",
        "117C-1㎡",
        "150㎡",
    ],
}


def _parse_area_from_unit_type(text: str) -> float | None:
    # "84B-1㎡" -> 84
    s = (text or "").strip()
    if not s:
        return None
    num = ""
    for ch in s:
        if ch.isdigit() or ch == ".":
            num += ch
        else:
            break
    try:
        return float(num) if num else None
    except Exception:
        return None


def _m2_to_pyeong(area_m2: float) -> float:
    return round(area_m2 / 3.305785, 1)


def _open_folder(path: Path) -> None:
    if not path.exists():
        messagebox.showinfo("안내", f"폴더가 없습니다: {path}")
        return
    system = platform.system().lower()
    try:
        if system.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif system == "darwin":
            os.system(f"open '{path}'")
        else:
            os.system(f"xdg-open '{path}'")
    except Exception as exc:
        messagebox.showerror("오류", f"폴더 열기 실패: {exc}")


@dataclass
class AppSettings:
    sync_dir: Path
    webhook_url: str


class LedgerDesktopApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("엄마 부동산 장부 (오프라인 우선)")
        self.root.geometry("1520x930")

        init_db()
        self.settings = self._load_settings()

        self.main = ttk.Notebook(root)
        self.main.pack(fill="both", expand=True)

        self.dashboard_tab = ttk.Frame(self.main)
        self.property_tab = ttk.Frame(self.main)
        self.customer_tab = ttk.Frame(self.main)
        self.matching_tab = ttk.Frame(self.main)
        self.settings_tab = ttk.Frame(self.main)

        self.main.add(self.dashboard_tab, text="오늘")
        self.main.add(self.property_tab, text="물건")
        self.main.add(self.customer_tab, text="고객")
        self.main.add(self.matching_tab, text="매칭")
        self.main.add(self.settings_tab, text="설정")

        self._build_dashboard_ui()
        self._build_property_ui()
        self._build_customer_ui()
        self._build_matching_ui()
        self._build_settings_ui()

        self.refresh_all()

    # -----------------
    # Settings
    # -----------------
    def _load_settings(self) -> AppSettings:
        # DB 설정이 우선, 없으면 env, 그것도 없으면 기본값
        sync_dir = get_setting("GOOGLE_DRIVE_SYNC_DIR", "").strip()
        webhook = get_setting("GOOGLE_SHEETS_WEBHOOK_URL", "").strip()

        env = SyncSettings.from_env()
        sync_dir_path = Path(sync_dir).expanduser() if sync_dir else env.sync_dir
        webhook_url = webhook if webhook else env.webhook_url
        return AppSettings(sync_dir=sync_dir_path, webhook_url=webhook_url)

    def _persist_settings(self) -> None:
        set_setting("GOOGLE_DRIVE_SYNC_DIR", str(self.settings.sync_dir))
        set_setting("GOOGLE_SHEETS_WEBHOOK_URL", self.settings.webhook_url)

    # -----------------
    # Dashboard
    # -----------------
    def _build_dashboard_ui(self):
        top = ttk.LabelFrame(self.dashboard_tab, text="오늘의 운영")
        top.pack(fill="x", padx=10, pady=10)

        self.dash_vars = {
            "props_total": tk.StringVar(value="0"),
            "props_hidden": tk.StringVar(value="0"),
            "customers_total": tk.StringVar(value="0"),
            "viewings_7d": tk.StringVar(value="0"),
            "tasks_open": tk.StringVar(value="0"),
        }

        row = 0
        ttk.Label(top, text="물건(전체)").grid(row=row, column=0, padx=6, pady=6, sticky="e")
        ttk.Label(top, textvariable=self.dash_vars["props_total"]).grid(row=row, column=1, padx=6, pady=6, sticky="w")

        ttk.Label(top, text="물건(숨김)").grid(row=row, column=2, padx=6, pady=6, sticky="e")
        ttk.Label(top, textvariable=self.dash_vars["props_hidden"]).grid(row=row, column=3, padx=6, pady=6, sticky="w")

        ttk.Label(top, text="고객(전체)").grid(row=row, column=4, padx=6, pady=6, sticky="e")
        ttk.Label(top, textvariable=self.dash_vars["customers_total"]).grid(row=row, column=5, padx=6, pady=6, sticky="w")

        ttk.Label(top, text="7일 일정").grid(row=row, column=6, padx=6, pady=6, sticky="e")
        ttk.Label(top, textvariable=self.dash_vars["viewings_7d"]).grid(row=row, column=7, padx=6, pady=6, sticky="w")

        ttk.Label(top, text="할 일(OPEN)").grid(row=row, column=8, padx=6, pady=6, sticky="e")
        ttk.Label(top, textvariable=self.dash_vars["tasks_open"]).grid(row=row, column=9, padx=6, pady=6, sticky="w")

        btns = ttk.Frame(top)
        btns.grid(row=1, column=0, columnspan=10, sticky="w", padx=6, pady=6)
        ttk.Button(btns, text="내보내기/동기화(Drive/웹훅)", command=self.export_sync).pack(side="left", padx=4)
        ttk.Button(btns, text="내보내기 폴더 열기", command=self.open_export_folder).pack(side="left", padx=4)

        # Main area: Tasks + Upcoming viewings
        pw = ttk.PanedWindow(self.dashboard_tab, orient=tk.HORIZONTAL)
        pw.pack(fill="both", expand=True, padx=10, pady=10)

        task_frame = ttk.LabelFrame(pw, text="할 일(Next Action)")
        schedule_frame = ttk.LabelFrame(pw, text="다가오는 일정(최근)")
        pw.add(task_frame, weight=1)
        pw.add(schedule_frame, weight=2)

        # ---- Tasks
        tctrl = ttk.Frame(task_frame)
        tctrl.pack(fill="x", padx=6, pady=6)
        ttk.Button(tctrl, text="새 할 일", command=self.open_add_task_window).pack(side="left", padx=4)
        ttk.Button(tctrl, text="완료", command=self.mark_selected_task_done).pack(side="left", padx=4)
        ttk.Button(tctrl, text="관련 열기", command=self.open_selected_task_related).pack(side="left", padx=4)
        ttk.Button(tctrl, text="새로고침", command=self.refresh_tasks).pack(side="left", padx=4)

        tcols = ("id", "due_at", "title", "entity", "status")
        self.tasks_tree = ttk.Treeview(task_frame, columns=tcols, show="headings", height=18)
        twidths = {"id": 55, "due_at": 140, "title": 520, "entity": 120, "status": 80}
        for c in tcols:
            self.tasks_tree.heading(c, text=c)
            self.tasks_tree.column(c, width=twidths.get(c, 120))
        self.tasks_tree.pack(fill="both", expand=True, padx=6, pady=6)
        self.tasks_tree.bind("<Double-1>", self._on_double_click_task)

        # ---- Upcoming viewings
        cols = ("id", "start", "end", "title", "property_id", "customer_id", "status")
        self.upcoming_tree = ttk.Treeview(schedule_frame, columns=cols, show="headings", height=18)
        widths = {"id": 60, "start": 150, "end": 150, "title": 360, "property_id": 90, "customer_id": 90, "status": 90}
        for c in cols:
            self.upcoming_tree.heading(c, text=c)
            self.upcoming_tree.column(c, width=widths.get(c, 120))
        self.upcoming_tree.pack(fill="both", expand=True, padx=6, pady=6)
        self.upcoming_tree.bind("<Double-1>", self._on_double_click_viewing)

    def refresh_dashboard(self):
        props = list_properties(include_deleted=False)
        custs = list_customers(include_deleted=False)
        self.dash_vars["props_total"].set(str(len(props)))
        self.dash_vars["props_hidden"].set(str(sum(1 for p in props if p.get("hidden"))))
        self.dash_vars["customers_total"].set(str(len(custs)))

        # open tasks count (MANUAL + AUTO)
        try:
            self.dash_vars["tasks_open"].set(str(len(list_tasks(include_done=False))))
        except Exception:
            self.dash_vars["tasks_open"].set("0")

        # next 7 days viewings
        viewings = list_viewings()
        now = datetime.now()
        end = now + timedelta(days=7)
        upcoming = []
        for v in viewings:
            try:
                st = datetime.fromisoformat(str(v.get("start_at")))
            except Exception:
                try:
                    st = datetime.strptime(str(v.get("start_at")), "%Y-%m-%d %H:%M")
                except Exception:
                    continue
            if now <= st <= end:
                upcoming.append(v)
        upcoming.sort(key=lambda x: x.get("start_at", ""))

        self.dash_vars["viewings_7d"].set(str(len(upcoming)))

        for i in self.upcoming_tree.get_children():
            self.upcoming_tree.delete(i)
        for v in upcoming[:200]:
            self.upcoming_tree.insert(
                "",
                "end",
                values=(
                    v.get("id"),
                    v.get("start_at"),
                    v.get("end_at"),
                    v.get("title"),
                    v.get("property_id"),
                    v.get("customer_id") or "",
                    v.get("status"),
                ),
            )

def refresh_tasks(self):
    # 자동 태스크 최신화
    try:
        reconcile_auto_tasks()
    except Exception:
        pass

    rows = list_tasks(include_done=False, limit=400)
    self.dash_vars.get("tasks_open", tk.StringVar()).set(str(len(rows)))

    for i in self.tasks_tree.get_children():
        self.tasks_tree.delete(i)

    for r in rows:
        entity = ""
        et = str(r.get("entity_type") or "").strip()
        eid = r.get("entity_id")
        if et and eid is not None:
            entity = f"{et}#{eid}"

        self.tasks_tree.insert(
            "",
            "end",
            values=(
                r.get("id"),
                r.get("due_at") or "",
                r.get("title") or "",
                entity,
                r.get("status") or "",
            ),
        )

def _selected_task_id(self) -> int | None:
    sel = self.tasks_tree.selection()
    if not sel:
        return None
    try:
        return int(self.tasks_tree.item(sel[0], "values")[0])
    except Exception:
        return None

def mark_selected_task_done(self):
    tid = self._selected_task_id()
    if tid is None:
        return
    try:
        mark_task_done(tid)
    except Exception as exc:
        messagebox.showerror("오류", f"완료 처리 실패: {exc}")
        return
    self.refresh_tasks()
    self.refresh_dashboard()

def open_selected_task_related(self):
    tid = self._selected_task_id()
    if tid is None:
        return
    row = None
    try:
        # list_tasks는 필터가 있어 get_task 대신 재조회
        tasks = list_tasks(include_done=True, limit=500)
        for t in tasks:
            if int(t.get("id")) == tid:
                row = t
                break
    except Exception:
        row = None

    if not row:
        return

    et = str(row.get("entity_type") or "").strip()
    eid = row.get("entity_id")

    try:
        if et == "PROPERTY" and eid is not None:
            self.open_property_detail(int(eid))
        elif et == "CUSTOMER" and eid is not None:
            self.open_customer_detail(int(eid))
        elif et == "VIEWING" and eid is not None:
            v = get_viewing(int(eid))
            if v and v.get("property_id"):
                self.open_property_detail(int(v.get("property_id")))
        else:
            messagebox.showinfo("안내", "연결된 대상이 없는 할 일입니다.")
    except Exception as exc:
        messagebox.showerror("오류", f"관련 열기 실패: {exc}")

def _on_double_click_task(self, _event):
    self.open_selected_task_related()

def open_add_task_window(self, *, default_entity_type: str = "", default_entity_id: int | None = None):
    win = tk.Toplevel(self.root)
    win.title("새 할 일 추가")
    win.geometry("520x320")

    vars_ = {
        "title": tk.StringVar(value=""),
        "due_at": tk.StringVar(value=""),
        "entity_type": tk.StringVar(value=default_entity_type),
        "entity_id": tk.StringVar(value=str(default_entity_id) if default_entity_id is not None else ""),
        "note": tk.StringVar(value=""),
    }

    frm = ttk.Frame(win)
    frm.pack(fill="both", expand=True, padx=12, pady=12)

    def row(r, label, widget):
        ttk.Label(frm, text=label).grid(row=r, column=0, padx=6, pady=6, sticky="e")
        widget.grid(row=r, column=1, padx=6, pady=6, sticky="w")

    row(0, "제목(필수)", ttk.Entry(frm, textvariable=vars_["title"], width=44))
    row(1, "기한(선택, YYYY-MM-DD HH:MM)", ttk.Entry(frm, textvariable=vars_["due_at"], width=44))

    et_combo = ttk.Combobox(frm, textvariable=vars_["entity_type"], values=["", "PROPERTY", "CUSTOMER", "VIEWING"], width=41, state="readonly")
    row(2, "연결대상(선택)", et_combo)
    row(3, "대상 ID(선택)", ttk.Entry(frm, textvariable=vars_["entity_id"], width=44))
    row(4, "메모(선택)", ttk.Entry(frm, textvariable=vars_["note"], width=44))

    hint = ttk.Label(frm, text="TIP: 연결대상+ID를 넣으면 '관련 열기'로 바로 이동합니다.", foreground="#666")
    hint.grid(row=5, column=0, columnspan=2, sticky="w", padx=6, pady=6)

    def save():
        title = vars_["title"].get().strip()
        if not title:
            messagebox.showwarning("확인", "제목은 필수입니다.")
            return

        due_at = vars_["due_at"].get().strip() or None
        if due_at:
            try:
                datetime.strptime(due_at, "%Y-%m-%d %H:%M")
            except Exception:
                messagebox.showwarning("확인", "기한 형식이 올바르지 않습니다. 예: 2026-02-13 14:30")
                return

        et = vars_["entity_type"].get().strip() or None
        eid_txt = vars_["entity_id"].get().strip()
        eid = int(eid_txt) if eid_txt.isdigit() else None
        note = vars_["note"].get().strip()

        try:
            add_task(title=title, due_at=due_at, entity_type=et, entity_id=eid, note=note, kind="MANUAL", status="OPEN")
        except Exception as exc:
            messagebox.showerror("오류", f"저장 실패: {exc}")
            return

        self.refresh_tasks()
        self.refresh_dashboard()
        win.destroy()

    btns = ttk.Frame(frm)
    btns.grid(row=6, column=0, columnspan=2, sticky="w", padx=6, pady=12)
    ttk.Button(btns, text="저장", command=save).pack(side="left", padx=4)
    ttk.Button(btns, text="취소", command=win.destroy).pack(side="left", padx=4)

    # -----------------
    # Properties
    # -----------------
    def _build_property_ui(self):
        top = ttk.LabelFrame(self.property_tab, text="물건 등록(빠른 입력)")
        top.pack(fill="x", padx=10, pady=8)

        self.pvars: dict[str, tk.Variable] = {
            "tab": tk.StringVar(value=PROPERTY_TABS[0]),
            "complex_name": tk.StringVar(value=TAB_COMPLEX_NAME.get(PROPERTY_TABS[0], "")),
            "unit_type": tk.StringVar(),
            "area": tk.StringVar(),
            "pyeong": tk.StringVar(),
            "address_detail": tk.StringVar(),
            "floor": tk.StringVar(),
            "total_floor": tk.StringVar(),
            "view": tk.StringVar(),
            "orientation": tk.StringVar(),
            "condition": tk.StringVar(value="중"),
            "repair_needed": tk.BooleanVar(value=False),
            "tenant_info": tk.StringVar(),
            "naver_link": tk.StringVar(),
            "special_notes": tk.StringVar(),
            "note": tk.StringVar(),
            "status": tk.StringVar(value="신규등록"),
        }

        self._p_widgets: dict[str, tk.Widget] = {}

        fields = [
            ("탭", "tab", "combo", PROPERTY_TABS),
            ("단지명", "complex_name", "entry", None),
            ("면적타입", "unit_type", "combo_free", None),
            ("동/호(상세)", "address_detail", "entry", None),
            ("면적(㎡)", "area", "entry", None),
            ("평형", "pyeong", "entry", None),
            ("층수", "floor", "entry", None),
            ("총층", "total_floor", "entry", None),
            ("조망/뷰", "view", "entry", None),
            ("향", "orientation", "entry", None),
            ("컨디션", "condition", "combo", ["상", "중", "하"]),
            ("상태", "status", "combo", ["신규등록", "검증필요", "광고중", "문의응대", "현장안내", "협상중", "계약중", "거래완료", "보관"]),
            ("세입자정보", "tenant_info", "entry", None),
            ("네이버링크", "naver_link", "entry", None),
            ("특이사항", "special_notes", "entry", None),
            ("별도기재", "note", "entry", None),
        ]

        for i, (label, key, kind, options) in enumerate(fields):
            ttk.Label(top, text=label).grid(row=i // 4, column=(i % 4) * 2, padx=6, pady=4, sticky="e")

            if kind == "combo":
                w: tk.Widget = ttk.Combobox(top, textvariable=self.pvars[key], values=options, width=22, state="readonly")
            elif kind == "combo_free":
                # 단지별로 리스트를 바꾸되, 임의 입력도 허용
                w = ttk.Combobox(top, textvariable=self.pvars[key], values=[], width=22, state="normal")
            else:
                w = ttk.Entry(top, textvariable=self.pvars[key], width=26)

            w.grid(row=i // 4, column=(i % 4) * 2 + 1, padx=6, pady=4, sticky="w")
            self._p_widgets[key] = w

        ttk.Checkbutton(top, text="수리필요", variable=self.pvars["repair_needed"]).grid(row=4, column=0, padx=6, pady=6, sticky="w")
        ttk.Button(top, text="물건 등록", command=self.create_property).grid(row=4, column=1, padx=6, pady=6, sticky="w")
        ttk.Button(top, text="내보내기/동기화", command=self.export_sync).grid(row=4, column=2, padx=6, pady=6, sticky="w")

        # bind tab change to update unit types / complex name
        if isinstance(self._p_widgets["tab"], ttk.Combobox):
            self._p_widgets["tab"].bind("<<ComboboxSelected>>", lambda _e: self._on_property_tab_changed())
        if isinstance(self._p_widgets["unit_type"], ttk.Combobox):
            self._p_widgets["unit_type"].bind("<<ComboboxSelected>>", lambda _e: self._on_unit_type_changed())

        self._on_property_tab_changed()

        # list area
        self.inner_tabs = ttk.Notebook(self.property_tab)
        self.inner_tabs.pack(fill="both", expand=True, padx=10, pady=8)
        self.prop_trees: dict[str, ttk.Treeview] = {}

        cols = ("id", "hidden", "status", "complex_name", "address_detail", "unit_type", "floor", "condition", "updated_at")
        col_defs = [
            ("id", 55),
            ("hidden", 70),
            ("status", 90),
            ("complex_name", 220),
            ("address_detail", 170),
            ("unit_type", 110),
            ("floor", 70),
            ("condition", 70),
            ("updated_at", 150),
        ]

        for tab_name in PROPERTY_TABS:
            frame = ttk.Frame(self.inner_tabs)
            display = tab_name
            if tab_name in TAB_COMPLEX_NAME:
                display = f"{tab_name} ({TAB_COMPLEX_NAME[tab_name]})"
            self.inner_tabs.add(frame, text=display)

            tree = ttk.Treeview(frame, columns=cols, show="headings", height=17)
            for c, w in col_defs:
                tree.heading(c, text=c)
                tree.column(c, width=w)
            tree.pack(fill="both", expand=True)
            tree.bind("<Double-1>", lambda e, t=tab_name: self._on_double_click_property(e, t))
            self.prop_trees[tab_name] = tree

            btns = ttk.Frame(frame)
            btns.pack(fill="x", pady=4)
            ttk.Button(btns, text="상세", command=lambda t=tab_name: self.open_selected_property_detail(t)).pack(side="left", padx=4)
            ttk.Button(btns, text="숨김/보임", command=lambda t=tab_name: self.toggle_selected_property(t)).pack(side="left", padx=4)
            ttk.Button(btns, text="삭제", command=lambda t=tab_name: self.delete_selected_property(t)).pack(side="left", padx=4)

    def _on_property_tab_changed(self):
        tab = str(self.pvars["tab"].get())
        # 단지명 자동
        if tab in TAB_COMPLEX_NAME:
            self.pvars["complex_name"].set(TAB_COMPLEX_NAME[tab])
            if isinstance(self._p_widgets.get("complex_name"), ttk.Entry):
                self._p_widgets["complex_name"].configure(state="readonly")
        else:
            if isinstance(self._p_widgets.get("complex_name"), ttk.Entry):
                self._p_widgets["complex_name"].configure(state="normal")

        # unit types drop-down
        unit_combo = self._p_widgets.get("unit_type")
        if isinstance(unit_combo, ttk.Combobox):
            values = UNIT_TYPES_BY_TAB.get(tab, [])
            unit_combo.configure(values=values)

    def _on_unit_type_changed(self):
        ut = str(self.pvars["unit_type"].get())
        area = _parse_area_from_unit_type(ut)
        if area is not None:
            self.pvars["area"].set(str(area))
            self.pvars["pyeong"].set(str(_m2_to_pyeong(area)))

    def create_property(self):
        data = {}
        for k, v in self.pvars.items():
            if isinstance(v, tk.BooleanVar):
                data[k] = bool(v.get())
            else:
                data[k] = str(v.get()).strip()
        add_property(data)
        self.refresh_all()

    def refresh_properties(self):
        for tab in PROPERTY_TABS:
            tree = self.prop_trees[tab]
            for i in tree.get_children():
                tree.delete(i)
            for row in list_properties(tab):
                tree.insert(
                    "",
                    "end",
                    values=(
                        row.get("id"),
                        "숨김" if row.get("hidden") else "보임",
                        row.get("status"),
                        row.get("complex_name"),
                        row.get("address_detail"),
                        row.get("unit_type"),
                        row.get("floor"),
                        row.get("condition"),
                        row.get("updated_at"),
                    ),
                )

    def _selected_id_from_tree(self, tree: ttk.Treeview) -> int | None:
        selected = tree.selection()
        if not selected:
            return None
        try:
            return int(tree.item(selected[0], "values")[0])
        except Exception:
            return None

    def toggle_selected_property(self, tab_name: str):
        tree = self.prop_trees[tab_name]
        pid = self._selected_id_from_tree(tree)
        if pid is None:
            return
        toggle_property_hidden(pid)
        self.refresh_all()

    def delete_selected_property(self, tab_name: str):
        tree = self.prop_trees[tab_name]
        pid = self._selected_id_from_tree(tree)
        if pid is None:
            return
        if messagebox.askyesno("확인", "해당 물건을 삭제 처리(복구 가능)할까요?"):
            soft_delete_property(pid)
            self.refresh_all()

    def open_selected_property_detail(self, tab_name: str):
        tree = self.prop_trees[tab_name]
        pid = self._selected_id_from_tree(tree)
        if pid is None:
            return
        self.open_property_detail(pid)

    def _on_double_click_property(self, _event, tab_name: str):
        self.open_selected_property_detail(tab_name)

    # -----------------
    # Customers
    # -----------------
    def _build_customer_ui(self):
        top = ttk.LabelFrame(self.customer_tab, text="고객 요구사항 등록")
        top.pack(fill="x", padx=10, pady=8)

        self.cvars = {
            "customer_name": tk.StringVar(),
            "phone": tk.StringVar(),
            "preferred_tab": tk.StringVar(value=PROPERTY_TABS[0]),
            "preferred_area": tk.StringVar(),
            "preferred_pyeong": tk.StringVar(),
            "budget": tk.StringVar(),
            "move_in_period": tk.StringVar(),
            "view_preference": tk.StringVar(),
            "location_preference": tk.StringVar(),
            "floor_preference": tk.StringVar(),
            "extra_needs": tk.StringVar(),
            "status": tk.StringVar(value="진행"),
        }

        fields = [
            ("고객명", "customer_name", "entry", None),
            ("전화번호", "phone", "entry", None),
            ("희망탭", "preferred_tab", "combo", PROPERTY_TABS),
            ("희망면적", "preferred_area", "entry", None),
            ("희망평형", "preferred_pyeong", "entry", None),
            ("예산", "budget", "entry", None),
            ("기간", "move_in_period", "entry", None),
            ("뷰", "view_preference", "entry", None),
            ("위치", "location_preference", "entry", None),
            ("층수선호", "floor_preference", "entry", None),
            ("기타요청", "extra_needs", "entry", None),
            ("상태", "status", "combo", ["진행", "보류", "완료"]),
        ]

        for i, (label, key, kind, options) in enumerate(fields):
            ttk.Label(top, text=label).grid(row=i // 4, column=(i % 4) * 2, padx=6, pady=4, sticky="e")
            if kind == "combo":
                w = ttk.Combobox(top, textvariable=self.cvars[key], values=options, width=22, state="readonly")
            else:
                w = ttk.Entry(top, textvariable=self.cvars[key], width=26)
            w.grid(row=i // 4, column=(i % 4) * 2 + 1, padx=6, pady=4, sticky="w")

        ttk.Button(top, text="고객 등록", command=self.create_customer).grid(row=3, column=0, padx=6, pady=6, sticky="w")
        ttk.Button(top, text="내보내기/동기화", command=self.export_sync).grid(row=3, column=1, padx=6, pady=6, sticky="w")

        cols = (
            "id",
            "hidden",
            "customer_name",
            "phone",
            "preferred_tab",
            "preferred_area",
            "budget",
            "floor_preference",
            "status",
            "updated_at",
        )
        self.customer_tree = ttk.Treeview(self.customer_tab, columns=cols, show="headings", height=22)
        col_defs = [
            ("id", 55),
            ("hidden", 70),
            ("customer_name", 120),
            ("phone", 120),
            ("preferred_tab", 110),
            ("preferred_area", 110),
            ("budget", 120),
            ("floor_preference", 120),
            ("status", 80),
            ("updated_at", 150),
        ]
        for c, w in col_defs:
            self.customer_tree.heading(c, text=c)
            self.customer_tree.column(c, width=w)
        self.customer_tree.pack(fill="both", expand=True, padx=10, pady=8)
        self.customer_tree.bind("<Double-1>", self._on_double_click_customer)

        btns = ttk.Frame(self.customer_tab)
        btns.pack(fill="x", padx=10, pady=4)
        ttk.Button(btns, text="상세", command=self.open_selected_customer_detail).pack(side="left", padx=4)
        ttk.Button(btns, text="숨김/보임", command=self.toggle_selected_customer).pack(side="left", padx=4)
        ttk.Button(btns, text="삭제", command=self.delete_selected_customer).pack(side="left", padx=4)

    def create_customer(self):
        if not self.cvars["customer_name"].get().strip():
            messagebox.showwarning("확인", "고객명은 필수입니다.")
            return
        add_customer({k: v.get() for k, v in self.cvars.items()})
        self.refresh_all()

    def refresh_customers(self):
        for i in self.customer_tree.get_children():
            self.customer_tree.delete(i)
        for row in list_customers():
            self.customer_tree.insert(
                "",
                "end",
                values=(
                    row.get("id"),
                    "숨김" if row.get("hidden") else "보임",
                    row.get("customer_name"),
                    row.get("phone"),
                    row.get("preferred_tab"),
                    row.get("preferred_area"),
                    row.get("budget"),
                    row.get("floor_preference"),
                    row.get("status"),
                    row.get("updated_at"),
                ),
            )

    def toggle_selected_customer(self):
        cid = self._selected_id_from_tree(self.customer_tree)
        if cid is None:
            return
        toggle_customer_hidden(cid)
        self.refresh_all()

    def delete_selected_customer(self):
        cid = self._selected_id_from_tree(self.customer_tree)
        if cid is None:
            return
        if messagebox.askyesno("확인", "해당 고객 요청을 삭제 처리(복구 가능)할까요?"):
            soft_delete_customer(cid)
            self.refresh_all()

    def open_selected_customer_detail(self):
        cid = self._selected_id_from_tree(self.customer_tree)
        if cid is None:
            return
        self.open_customer_detail(cid)

    def _on_double_click_customer(self, _event):
        self.open_selected_customer_detail()

    # -----------------
    # Matching
    # -----------------
    def _build_matching_ui(self):
        top = ttk.LabelFrame(self.matching_tab, text="고객 → 물건 매칭(규칙 기반)")
        top.pack(fill="x", padx=10, pady=10)

        ttk.Label(top, text="고객 선택").grid(row=0, column=0, padx=6, pady=6, sticky="e")
        self.match_customer_var = tk.StringVar()
        self.match_customer_combo = ttk.Combobox(top, textvariable=self.match_customer_var, values=[], width=50, state="readonly")
        self.match_customer_combo.grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ttk.Button(top, text="새로고침", command=self.refresh_matching_customers).grid(row=0, column=2, padx=6, pady=6)
        ttk.Button(top, text="추천 생성", command=self.run_matching).grid(row=0, column=3, padx=6, pady=6)
        ttk.Button(top, text="제안서 PDF(선택/상위)", command=self.generate_proposal_from_matching).grid(row=0, column=4, padx=6, pady=6)
        ttk.Button(top, text="제안문구 복사", command=self.copy_message_from_matching).grid(row=0, column=5, padx=6, pady=6)
        ttk.Button(top, text="제안서 폴더", command=self.open_proposals_folder).grid(row=0, column=6, padx=6, pady=6)

        bottom = ttk.Frame(self.matching_tab)
        bottom.pack(fill="both", expand=True, padx=10, pady=10)

        cols = ("score", "property_id", "tab", "complex", "address", "unit_type", "floor", "condition", "reasons")
        self.match_tree = ttk.Treeview(bottom, columns=cols, show="headings", height=22, selectmode="extended")
        col_defs = [
            ("score", 60),
            ("property_id", 90),
            ("tab", 110),
            ("complex", 220),
            ("address", 160),
            ("unit_type", 110),
            ("floor", 70),
            ("condition", 80),
            ("reasons", 420),
        ]
        for c, w in col_defs:
            self.match_tree.heading(c, text=c)
            self.match_tree.column(c, width=w)
        self.match_tree.pack(fill="both", expand=True)
        self.match_tree.bind("<Double-1>", self._on_double_click_match)

        self.refresh_matching_customers()

    def refresh_matching_customers(self):
        rows = [r for r in list_customers() if not r.get("hidden") and not r.get("deleted")]
        values = [f"{r['id']} | {r.get('customer_name','')} | {r.get('phone','')}" for r in rows]
        self._match_customer_index = {values[i]: rows[i]["id"] for i in range(len(values))}
        self.match_customer_combo.configure(values=values)
        if values and not self.match_customer_var.get():
            self.match_customer_var.set(values[0])

    def run_matching(self):
        key = self.match_customer_var.get()
        if not key:
            return
        cid = self._match_customer_index.get(key)
        if not cid:
            return
        customer = get_customer(cid)
        if not customer:
            messagebox.showerror("오류", "고객 정보를 찾을 수 없습니다.")
            return

        props = [p for p in list_properties(include_deleted=False) if not p.get("hidden")]
        results = match_properties(customer, props, limit=60)

        for i in self.match_tree.get_children():
            self.match_tree.delete(i)
        for r in results:
            p = r.property_row
            self.match_tree.insert(
                "",
                "end",
                values=(
                    r.score,
                    p.get("id"),
                    p.get("tab"),
                    p.get("complex_name"),
                    p.get("address_detail"),
                    p.get("unit_type"),
                    p.get("floor"),
                    p.get("condition"),
                    ", ".join(r.reasons),
                ),
            )

    def _on_double_click_match(self, _event):
        selected = self.match_tree.selection()
        if not selected:
            return
        try:
            pid = int(self.match_tree.item(selected[0], "values")[1])
        except Exception:
            return
        self.open_property_detail(pid)


def _get_current_matching_customer(self) -> dict | None:
    key = self.match_customer_var.get()
    if not key:
        return None
    cid = getattr(self, "_match_customer_index", {}).get(key)
    if not cid:
        return None
    return get_customer(int(cid))

def _get_selected_match_property_ids(self, *, fallback_top_n: int = 5) -> list[int]:
    ids: list[int] = []
    for item in self.match_tree.selection():
        try:
            pid = int(self.match_tree.item(item, "values")[1])
            ids.append(pid)
        except Exception:
            continue
    if ids:
        return ids

    # fallback: first N rows in the tree
    for item in self.match_tree.get_children()[:fallback_top_n]:
        try:
            pid = int(self.match_tree.item(item, "values")[1])
            ids.append(pid)
        except Exception:
            continue
    return ids

def open_proposals_folder(self):
    self.settings = self._load_settings()
    out_dir = self.settings.sync_dir / "exports" / "proposals"
    out_dir.mkdir(parents=True, exist_ok=True)
    _open_folder(out_dir)

def generate_proposal_from_matching(self):
    customer = self._get_current_matching_customer()
    if not customer:
        messagebox.showwarning("확인", "고객을 먼저 선택해주세요.")
        return

    pids = self._get_selected_match_property_ids(fallback_top_n=5)
    if not pids:
        messagebox.showwarning("확인", "추천 결과가 없습니다. 먼저 '추천 생성'을 눌러주세요.")
        return

    props: list[dict] = []
    photos_by_property: dict[int, list[str]] = {}

    for pid in pids:
        row = get_property(pid)
        if not row:
            continue
        props.append(row)
        photos_by_property[pid] = [r.get("file_path") for r in list_photos(pid) if r.get("file_path")]

    self.settings = self._load_settings()
    out_dir = self.settings.sync_dir / "exports" / "proposals"

    try:
        out = generate_proposal_pdf(
            customer=customer,
            properties=props,
            photos_by_property=photos_by_property,
            output_dir=out_dir,
            title="매물 제안서",
        )
    except Exception as exc:
        messagebox.showerror("오류", f"제안서 생성 실패: {exc}")
        return

    messagebox.showinfo("완료", f"제안서 생성 완료\n- PDF: {out.pdf_path.name}\n- TXT: {out.txt_path.name}")
    try:
        _open_folder(out_dir)
    except Exception:
        pass

def copy_message_from_matching(self):
    customer = self._get_current_matching_customer()
    if not customer:
        messagebox.showwarning("확인", "고객을 먼저 선택해주세요.")
        return

    pids = self._get_selected_match_property_ids(fallback_top_n=5)
    if not pids:
        messagebox.showwarning("확인", "추천 결과가 없습니다. 먼저 '추천 생성'을 눌러주세요.")
        return

    props: list[dict] = []
    for pid in pids:
        row = get_property(pid)
        if row:
            props.append(row)

    msg = build_kakao_message(customer, props, include_links=True)
    try:
        self.root.clipboard_clear()
        self.root.clipboard_append(msg)
        messagebox.showinfo("복사 완료", "제안문구를 클립보드에 복사했습니다. (카톡/문자에 붙여넣기)")
    except Exception as exc:
        messagebox.showerror("오류", f"클립보드 복사 실패: {exc}")

    # -----------------
    # Settings tab
    # -----------------
    def _build_settings_ui(self):
        box = ttk.LabelFrame(self.settings_tab, text="동기화 설정(관리자)")
        box.pack(fill="x", padx=10, pady=10)

        self.set_sync_dir_var = tk.StringVar(value=str(self.settings.sync_dir))
        self.set_webhook_var = tk.StringVar(value=self.settings.webhook_url)

        ttk.Label(box, text="Drive 동기화 폴더").grid(row=0, column=0, padx=6, pady=6, sticky="e")
        ttk.Entry(box, textvariable=self.set_sync_dir_var, width=80).grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ttk.Button(box, text="폴더 선택", command=self.browse_sync_dir).grid(row=0, column=2, padx=6, pady=6)

        ttk.Label(box, text="(선택) Sheets 웹훅 URL").grid(row=1, column=0, padx=6, pady=6, sticky="e")
        ttk.Entry(box, textvariable=self.set_webhook_var, width=80).grid(row=1, column=1, padx=6, pady=6, sticky="w")

        btns = ttk.Frame(box)
        btns.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=6)
        ttk.Button(btns, text="저장", command=self.save_settings).pack(side="left", padx=4)
        ttk.Button(btns, text="내보내기/동기화 실행", command=self.export_sync).pack(side="left", padx=4)
        ttk.Button(btns, text="폴더 열기", command=lambda: _open_folder(Path(self.set_sync_dir_var.get()).expanduser())).pack(side="left", padx=4)

        hint = ttk.Label(
            self.settings_tab,
            text=(
                "TIP: Drive 동기화 폴더를 'Google Drive' 안의 폴더로 지정하면,\n"
                "     내보낸 CSV/JSON/ICS/사진을 어디서나 열람할 수 있습니다. (API 없이도 운영 가능)"
            ),
        )
        hint.pack(anchor="w", padx=14)

    def browse_sync_dir(self):
        d = filedialog.askdirectory(title="동기화 폴더 선택")
        if d:
            self.set_sync_dir_var.set(d)

    def save_settings(self):
        sync_dir = Path(self.set_sync_dir_var.get()).expanduser()
        webhook = self.set_webhook_var.get().strip()
        self.settings = AppSettings(sync_dir=sync_dir, webhook_url=webhook)
        self._persist_settings()
        messagebox.showinfo("완료", "설정이 저장되었습니다.")

    # -----------------
    # Sync / Export
    # -----------------
    def export_sync(self):
        self.settings = self._load_settings()  # 최신 반영
        props = list_properties(include_deleted=False)
        custs = list_customers(include_deleted=False)
        photos = list_photos_all()
        viewings = list_viewings()

        ok, msg = upload_visible_data(
            props,
            custs,
            photos=photos,
            viewings=viewings,
            tasks=tasks,
            settings=SyncSettings(webhook_url=self.settings.webhook_url, sync_dir=self.settings.sync_dir),
        )
        if ok:
            messagebox.showinfo("동기화", msg)
        else:
            messagebox.showerror("동기화 실패", msg)

    def open_export_folder(self):
        self.settings = self._load_settings()
        _open_folder(self.settings.sync_dir / "exports")

    # -----------------
    # Detail windows
    # -----------------
    def open_property_detail(self, property_id: int):
        row = get_property(property_id)
        if not row:
            messagebox.showerror("오류", "물건 정보를 찾을 수 없습니다.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"물건 상세 - ID {property_id}")
        win.geometry("1100x760")

        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True)

        tab_basic = ttk.Frame(nb)
        tab_photos = ttk.Frame(nb)
        tab_viewings = ttk.Frame(nb)
        nb.add(tab_basic, text="기본")
        nb.add(tab_photos, text="사진")
        nb.add(tab_viewings, text="일정")

        # ---- basic form
        vars_: dict[str, tk.Variable] = {
            "tab": tk.StringVar(value=str(row.get("tab", ""))),
            "complex_name": tk.StringVar(value=str(row.get("complex_name", ""))),
            "unit_type": tk.StringVar(value=str(row.get("unit_type", ""))),
            "area": tk.StringVar(value=str(row.get("area", "") or "")),
            "pyeong": tk.StringVar(value=str(row.get("pyeong", "") or "")),
            "address_detail": tk.StringVar(value=str(row.get("address_detail", ""))),
            "floor": tk.StringVar(value=str(row.get("floor", ""))),
            "total_floor": tk.StringVar(value=str(row.get("total_floor", ""))),
            "view": tk.StringVar(value=str(row.get("view", ""))),
            "orientation": tk.StringVar(value=str(row.get("orientation", ""))),
            "condition": tk.StringVar(value=str(row.get("condition", ""))),
            "repair_needed": tk.BooleanVar(value=bool(row.get("repair_needed"))),
            "tenant_info": tk.StringVar(value=str(row.get("tenant_info", ""))),
            "naver_link": tk.StringVar(value=str(row.get("naver_link", ""))),
            "special_notes": tk.StringVar(value=str(row.get("special_notes", ""))),
            "note": tk.StringVar(value=str(row.get("note", ""))),
            "status": tk.StringVar(value=str(row.get("status", ""))),
        }

        form = ttk.LabelFrame(tab_basic, text="물건 정보")
        form.pack(fill="x", padx=10, pady=10)

        def add_row(r, c, label, widget):
            ttk.Label(form, text=label).grid(row=r, column=c, padx=6, pady=4, sticky="e")
            widget.grid(row=r, column=c + 1, padx=6, pady=4, sticky="w")

        # row 0
        add_row(0, 0, "탭", ttk.Entry(form, textvariable=vars_["tab"], width=24, state="readonly"))
        add_row(0, 2, "단지명", ttk.Entry(form, textvariable=vars_["complex_name"], width=36))
        add_row(0, 4, "상태", ttk.Combobox(form, textvariable=vars_["status"], width=18, state="readonly",
                                             values=["신규등록", "검증필요", "광고중", "문의응대", "현장안내", "협상중", "계약중", "거래완료", "보관"]))

        # row 1
        add_row(1, 0, "동/호(상세)", ttk.Entry(form, textvariable=vars_["address_detail"], width=24))
        add_row(1, 2, "면적타입", ttk.Entry(form, textvariable=vars_["unit_type"], width=18))
        add_row(1, 4, "층/총층", ttk.Entry(form, textvariable=vars_["floor"], width=7))
        ttk.Label(form, text="/").grid(row=1, column=6, sticky="w")
        ttk.Entry(form, textvariable=vars_["total_floor"], width=7).grid(row=1, column=7, padx=2, pady=4, sticky="w")

        # row 2
        add_row(2, 0, "면적(㎡)", ttk.Entry(form, textvariable=vars_["area"], width=24))
        add_row(2, 2, "평형", ttk.Entry(form, textvariable=vars_["pyeong"], width=18))
        add_row(2, 4, "컨디션", ttk.Combobox(form, textvariable=vars_["condition"], width=18, state="readonly", values=["상", "중", "하"]))

        # row 3
        add_row(3, 0, "조망/뷰", ttk.Entry(form, textvariable=vars_["view"], width=24))
        add_row(3, 2, "향", ttk.Entry(form, textvariable=vars_["orientation"], width=18))
        ttk.Checkbutton(form, text="수리필요", variable=vars_["repair_needed"]).grid(row=3, column=4, padx=6, pady=4, sticky="w")

        # row 4
        add_row(4, 0, "세입자정보", ttk.Entry(form, textvariable=vars_["tenant_info"], width=60))

        # row 5
        add_row(5, 0, "네이버링크", ttk.Entry(form, textvariable=vars_["naver_link"], width=60))

        # row 6
        add_row(6, 0, "특이사항", ttk.Entry(form, textvariable=vars_["special_notes"], width=60))

        # row 7
        add_row(7, 0, "별도기재", ttk.Entry(form, textvariable=vars_["note"], width=60))

        actions = ttk.Frame(tab_basic)
        actions.pack(fill="x", padx=10, pady=10)

        def save_changes():
            data = {k: (v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get())) for k, v in vars_.items()}
            # area/pyeong numeric as strings ok; storage converts
            update_property(property_id, data)
            messagebox.showinfo("완료", "저장되었습니다.")
            self.refresh_all()

        def toggle_hide():
            toggle_property_hidden(property_id)
            messagebox.showinfo("완료", "상태가 변경되었습니다(숨김/보임).")
            self.refresh_all()

        def soft_delete():
            if messagebox.askyesno("확인", "삭제 처리(복구 가능) 하시겠습니까?"):
                soft_delete_property(property_id)
                self.refresh_all()
                win.destroy()

        def open_link():
            url = str(vars_["naver_link"].get()).strip()
            if not url:
                return
            webbrowser.open(url)

        ttk.Button(actions, text="저장", command=save_changes).pack(side="left", padx=4)
        ttk.Button(actions, text="네이버 링크 열기", command=open_link).pack(side="left", padx=4)
        ttk.Button(actions, text="숨김/보임", command=toggle_hide).pack(side="left", padx=4)
        ttk.Button(actions, text="삭제", command=soft_delete).pack(side="left", padx=4)
        ttk.Button(actions, text="할 일 추가", command=lambda: self.open_add_task_window(default_entity_type="PROPERTY", default_entity_id=property_id)).pack(side="left", padx=4)

        # ---- photos tab
        ph_box = ttk.LabelFrame(tab_photos, text="사진")
        ph_box.pack(fill="both", expand=True, padx=10, pady=10)

        ph_cols = ("id", "tag", "file_path", "created_at")
        ph_tree = ttk.Treeview(ph_box, columns=ph_cols, show="headings", height=18)
        ph_defs = [("id", 60), ("tag", 140), ("file_path", 660), ("created_at", 160)]
        for c, w in ph_defs:
            ph_tree.heading(c, text=c)
            ph_tree.column(c, width=w)
        ph_tree.pack(fill="both", expand=True)

        ph_controls = ttk.Frame(ph_box)
        ph_controls.pack(fill="x", pady=6)
        ph_tag_var = tk.StringVar()
        ttk.Label(ph_controls, text="태그").pack(side="left", padx=4)
        ttk.Entry(ph_controls, textvariable=ph_tag_var, width=20).pack(side="left", padx=4)

        def refresh_photos():
            for i in ph_tree.get_children():
                ph_tree.delete(i)
            for r in list_photos(property_id):
                ph_tree.insert("", "end", values=(r.get("id"), r.get("tag"), r.get("file_path"), r.get("created_at")))

        def add_photo_ui():
            src = filedialog.askopenfilename(title="사진 선택")
            if not src:
                return
            dst = self._copy_photo_to_library(Path(src), property_id)
            add_photo(property_id, str(dst), tag=ph_tag_var.get().strip())
            ph_tag_var.set("")
            refresh_photos()

        def open_photo():
            sel = ph_tree.selection()
            if not sel:
                return
            path = ph_tree.item(sel[0], "values")[2]
            if path:
                try:
                    _open_folder(Path(path).parent)
                except Exception:
                    pass

        def remove_photo():
            sel = ph_tree.selection()
            if not sel:
                return
            pid_ = int(ph_tree.item(sel[0], "values")[0])
            if messagebox.askyesno("확인", "사진 기록을 삭제할까요? (파일은 남습니다)"):
                delete_photo(pid_)
                refresh_photos()

        ttk.Button(ph_controls, text="추가", command=add_photo_ui).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="폴더 열기", command=open_photo).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="기록 삭제", command=remove_photo).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="내보내기/동기화", command=self.export_sync).pack(side="right", padx=4)

        refresh_photos()

        # ---- viewings tab
        vw_box = ttk.LabelFrame(tab_viewings, text="일정(캘린더 API 없이도 저장/내보내기 가능)")
        vw_box.pack(fill="both", expand=True, padx=10, pady=10)

        vw_cols = ("id", "start_at", "end_at", "title", "customer_id", "status")
        vw_tree = ttk.Treeview(vw_box, columns=vw_cols, show="headings", height=18)
        vw_defs = [("id", 60), ("start_at", 160), ("end_at", 160), ("title", 360), ("customer_id", 90), ("status", 90)]
        for c, w in vw_defs:
            vw_tree.heading(c, text=c)
            vw_tree.column(c, width=w)
        vw_tree.pack(fill="both", expand=True)

        vw_controls = ttk.Frame(vw_box)
        vw_controls.pack(fill="x", pady=6)

        def refresh_viewings():
            for i in vw_tree.get_children():
                vw_tree.delete(i)
            rows = list_viewings(property_id=property_id)
            for r in rows:
                vw_tree.insert(
                    "",
                    "end",
                    values=(r.get("id"), r.get("start_at"), r.get("end_at"), r.get("title"), r.get("customer_id") or "", r.get("status")),
                )

        def add_viewing_ui():
            self._open_add_viewing_window(property_id, refresh_viewings)

        def mark_done():
            sel = vw_tree.selection()
            if not sel:
                return
            vid = int(vw_tree.item(sel[0], "values")[0])
            update_viewing_status(vid, "완료")
            refresh_viewings()
            self.refresh_all()

        def remove_viewing():
            sel = vw_tree.selection()
            if not sel:
                return
            vid = int(vw_tree.item(sel[0], "values")[0])
            if messagebox.askyesno("확인", "일정 기록을 삭제할까요?"):
                delete_viewing(vid)
                refresh_viewings()
                self.refresh_all()

        ttk.Button(vw_controls, text="일정 추가", command=add_viewing_ui).pack(side="left", padx=4)
        ttk.Button(vw_controls, text="완료 처리", command=mark_done).pack(side="left", padx=4)
        ttk.Button(vw_controls, text="삭제", command=remove_viewing).pack(side="left", padx=4)
        ttk.Button(vw_controls, text="내보내기/동기화", command=self.export_sync).pack(side="right", padx=4)

        refresh_viewings()

    def _copy_photo_to_library(self, src: Path, property_id: int) -> Path:
        self.settings = self._load_settings()
        base = self.settings.sync_dir
        dst_dir = base / "Photos" / f"PR_{property_id:06d}"
        dst_dir.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = src.name.replace(" ", "_")
        dst = dst_dir / f"{property_id}_{ts}_{safe_name}"
        shutil.copy2(src, dst)
        return dst

    def _open_add_viewing_window(self, property_id: int, on_saved):
        win = tk.Toplevel(self.root)
        win.title("일정 추가")
        win.geometry("520x300")

        vars_ = {
            "start": tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M")),
            "end": tk.StringVar(value=(datetime.now() + timedelta(hours=1)).strftime("%Y-%m-%d %H:%M")),
            "customer_id": tk.StringVar(value=""),
            "title": tk.StringVar(value="현장/상담"),
            "memo": tk.StringVar(value=""),
        }

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        def row(r, label, key, width=42):
            ttk.Label(frm, text=label).grid(row=r, column=0, padx=6, pady=6, sticky="e")
            ttk.Entry(frm, textvariable=vars_[key], width=width).grid(row=r, column=1, padx=6, pady=6, sticky="w")

        row(0, "시작(YYYY-MM-DD HH:MM)", "start")
        row(1, "종료(YYYY-MM-DD HH:MM)", "end")
        row(2, "고객 ID(선택)", "customer_id")
        row(3, "제목", "title")
        row(4, "메모", "memo")

        def save():
            start = vars_["start"].get().strip()
            end = vars_["end"].get().strip()
            try:
                datetime.strptime(start, "%Y-%m-%d %H:%M")
                datetime.strptime(end, "%Y-%m-%d %H:%M")
            except Exception:
                messagebox.showwarning("확인", "시간 형식이 올바르지 않습니다. 예: 2026-02-13 14:30")
                return

            cid_text = vars_["customer_id"].get().strip()
            cid = int(cid_text) if cid_text.isdigit() else None
            title = vars_["title"].get().strip() or "현장/상담"
            memo = vars_["memo"].get().strip()

            add_viewing(
                property_id=property_id,
                customer_id=cid,
                start_at=start,
                end_at=end,
                title=title,
                memo=memo,
                status="예정",
            )
            on_saved()
            self.refresh_all()
            win.destroy()

        ttk.Button(frm, text="저장", command=save).grid(row=5, column=0, padx=6, pady=12)
        ttk.Button(frm, text="취소", command=win.destroy).grid(row=5, column=1, padx=6, pady=12, sticky="w")

    def open_customer_detail(self, customer_id: int):
        row = get_customer(customer_id)
        if not row:
            messagebox.showerror("오류", "고객 정보를 찾을 수 없습니다.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"고객 상세 - ID {customer_id}")
        win.geometry("980x560")

        vars_ = {k: tk.StringVar(value=str(row.get(k, "") or "")) for k in row.keys()}
        # boolean hidden not edited here

        form = ttk.LabelFrame(win, text="고객 정보")
        form.pack(fill="x", padx=10, pady=10)

        fields = [
            ("고객명", "customer_name"),
            ("전화", "phone"),
            ("희망탭", "preferred_tab"),
            ("희망면적", "preferred_area"),
            ("희망평형", "preferred_pyeong"),
            ("예산", "budget"),
            ("기간", "move_in_period"),
            ("뷰", "view_preference"),
            ("위치", "location_preference"),
            ("층수선호", "floor_preference"),
            ("기타요청", "extra_needs"),
            ("상태", "status"),
        ]

        for i, (label, key) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i // 3, column=(i % 3) * 2, padx=6, pady=6, sticky="e")
            if key == "preferred_tab":
                w = ttk.Combobox(form, textvariable=vars_[key], values=PROPERTY_TABS, width=22, state="readonly")
            elif key == "status":
                w = ttk.Combobox(form, textvariable=vars_[key], values=["진행", "보류", "완료"], width=22, state="readonly")
            else:
                w = ttk.Entry(form, textvariable=vars_[key], width=26)
            w.grid(row=i // 3, column=(i % 3) * 2 + 1, padx=6, pady=6, sticky="w")

        actions = ttk.Frame(win)
        actions.pack(fill="x", padx=10, pady=10)

        def save_changes():
            data = {k: v.get() for k, v in vars_.items()}
            update_customer(customer_id, data)
            messagebox.showinfo("완료", "저장되었습니다.")
            self.refresh_all()

        def toggle_hide():
            toggle_customer_hidden(customer_id)
            messagebox.showinfo("완료", "숨김/보임이 변경되었습니다.")
            self.refresh_all()

        def soft_delete():
            if messagebox.askyesno("확인", "삭제 처리(복구 가능) 하시겠습니까?"):
                soft_delete_customer(customer_id)
                self.refresh_all()
                win.destroy()

        ttk.Button(actions, text="저장", command=save_changes).pack(side="left", padx=4)
        ttk.Button(actions, text="숨김/보임", command=toggle_hide).pack(side="left", padx=4)
        ttk.Button(actions, text="삭제", command=soft_delete).pack(side="left", padx=4)
        ttk.Button(actions, text="매칭 보기", command=lambda: self._open_matching_for_customer(customer_id)).pack(side="left", padx=4)

    def _open_matching_for_customer(self, customer_id: int):
        # 매칭 탭으로 전환하고 고객 선택
        self.main.select(self.matching_tab)
        self.refresh_matching_customers()
        # combobox value 찾아서 set
        for label, cid in getattr(self, "_match_customer_index", {}).items():
            if cid == customer_id:
                self.match_customer_var.set(label)
                break
        self.run_matching()


def _top_matches_for_customer(self, customer_id: int, top_n: int = 5) -> tuple[dict | None, list[dict]]:
    customer = get_customer(customer_id)
    if not customer:
        return None, []
    props = [p for p in list_properties(include_deleted=False) if not p.get("hidden")]
    results = match_properties(customer, props, limit=max(1, int(top_n)))
    matched = [r.property_row for r in results]
    return customer, matched

def generate_proposal_for_customer(self, customer_id: int, top_n: int = 5):
    customer, props = self._top_matches_for_customer(customer_id, top_n=top_n)
    if not customer:
        messagebox.showerror("오류", "고객 정보를 찾을 수 없습니다.")
        return
    if not props:
        messagebox.showwarning("안내", "추천할 물건이 없습니다. 고객 조건을 조금 완화하거나 물건을 추가해 주세요.")
        return

    photos_by_property: dict[int, list[str]] = {}
    for p in props:
        try:
            pid = int(p.get("id"))
        except Exception:
            continue
        photos_by_property[pid] = [r.get("file_path") for r in list_photos(pid) if r.get("file_path")]

    self.settings = self._load_settings()
    out_dir = self.settings.sync_dir / "exports" / "proposals"

    try:
        out = generate_proposal_pdf(
            customer=customer,
            properties=props,
            photos_by_property=photos_by_property,
            output_dir=out_dir,
            title="매물 제안서",
        )
    except Exception as exc:
        messagebox.showerror("오류", f"제안서 생성 실패: {exc}")
        return

    messagebox.showinfo("완료", f"제안서 생성 완료\n- PDF: {out.pdf_path.name}\n- TXT: {out.txt_path.name}")
    try:
        _open_folder(out_dir)
    except Exception:
        pass

def copy_message_for_customer(self, customer_id: int, top_n: int = 5):
    customer, props = self._top_matches_for_customer(customer_id, top_n=top_n)
    if not customer:
        messagebox.showerror("오류", "고객 정보를 찾을 수 없습니다.")
        return
    if not props:
        messagebox.showwarning("안내", "추천할 물건이 없습니다. 고객 조건을 조금 완화하거나 물건을 추가해 주세요.")
        return

    msg = build_kakao_message(customer, props, include_links=True)
    try:
        self.root.clipboard_clear()
        self.root.clipboard_append(msg)
        messagebox.showinfo("복사 완료", "제안문구를 클립보드에 복사했습니다. (카톡/문자에 붙여넣기)")
    except Exception as exc:
        messagebox.showerror("오류", f"클립보드 복사 실패: {exc}")

    def _on_double_click_viewing(self, _event):
        # 현재는 별도 상세창은 생략(필요하면 확장)
        pass

    # -----------------
    # Refresh
    # -----------------
    def refresh_all(self):
        # 자동 '할 일' 갱신(조용히)
        try:
            reconcile_auto_tasks()
        except Exception:
            pass

        self.refresh_properties()
        self.refresh_customers()
        self.refresh_tasks()
        self.refresh_dashboard()
        self.refresh_matching_customers()


def run_desktop_app():
    root = tk.Tk()
    LedgerDesktopApp(root)
    root.mainloop()
