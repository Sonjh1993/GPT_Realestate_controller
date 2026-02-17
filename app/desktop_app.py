from __future__ import annotations

import os
import platform
import re
import shutil
import webbrowser
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from .matching import match_properties
from .proposal import build_kakao_message, generate_proposal_pdf
from .tasks_engine import reconcile_auto_tasks
from .sheet_sync import SyncSettings, upload_visible_data
from . import unit_master
from . import money_utils
from .storage import (
    CUSTOMER_STATUS_VALUES,
    PHOTO_TAG_VALUES,
    PROPERTY_STATUS_VALUES,
    PROPERTY_TABS,
    TAB_COMPLEX_NAME,
    TASK_TYPE_VALUES,
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


def _open_file(path: Path) -> None:
    if not path.exists():
        messagebox.showinfo("안내", f"파일이 없습니다: {path}")
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
        messagebox.showerror("오류", f"파일 열기 실패: {exc}")


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
        self.root.geometry("1920x1080")
        try:
            self.root.state("zoomed")
        except Exception:
            pass
        self._apply_ui_scale(1.5)

        self.sort_state: dict[tuple[int, str], bool] = {}
        self.task_sort_col = "due_at"
        self.task_sort_desc = False

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
        self.start_auto_sync_loop()

    def _apply_ui_scale(self, scale: float = 1.5) -> None:
        try:
            self.root.tk.call("tk", "scaling", scale)
        except Exception:
            pass

        # 기본 Tk 폰트들을 2배 확장
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"):
            try:
                f = tkfont.nametofont(name)
                size = abs(int(f.cget("size") or 10))
                f.configure(size=max(10, int(size * scale)))
            except Exception:
                continue

        # ttk 위젯 스타일 확대
        style = ttk.Style(self.root)
        base_size = 15
        style.configure("TLabel", font=("맑은 고딕", base_size))
        style.configure("TButton", font=("맑은 고딕", base_size), padding=(9, 7))
        style.configure("TEntry", font=("맑은 고딕", base_size), padding=(6, 6))
        style.configure("TCombobox", font=("맑은 고딕", base_size), padding=(6, 6))
        style.configure("TCheckbutton", font=("맑은 고딕", base_size))
        style.configure("TRadiobutton", font=("맑은 고딕", base_size))
        style.configure("TNotebook.Tab", font=("맑은 고딕", base_size), padding=(14, 9))
        style.configure("Treeview", font=("맑은 고딕", base_size), rowheight=33)
        style.configure("Treeview.Heading", font=("맑은 고딕", base_size, "bold"))
        style.configure("TLabelframe.Label", font=("맑은 고딕", base_size, "bold"))

    def _fit_toplevel(self, win: tk.Toplevel, width: int, height: int) -> None:
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            max_w = max(520, sw - 20)
            max_h = max(460, sh - 40)
            w = min(max(width, int(sw * 0.65)), max_w)
            h = min(max(height, int(sh * 0.82)), max_h)
            x = max(0, (sw - w) // 2)
            y = max(8, (sh - h) // 5)
            win.geometry(f"{w}x{h}+{x}+{y}")
            win.minsize(min(760, max_w), min(520, max_h))
            win.resizable(True, True)
        except Exception:
            pass

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

        self.sync_status_var = tk.StringVar(value="마지막 동기화: 없음")
        ttk.Label(top, textvariable=self.sync_status_var, foreground="#2f4f4f").grid(row=1, column=0, columnspan=10, sticky="w", padx=6, pady=2)

        btns = ttk.Frame(top)
        btns.grid(row=2, column=0, columnspan=10, sticky="w", padx=6, pady=6)
        ttk.Button(btns, text="내보내기/동기화(Drive/웹훅)", command=self.export_sync).pack(side="left", padx=4)
        ttk.Button(btns, text="내보내기 폴더 열기", command=self.open_export_folder).pack(side="left", padx=4)

        # Main area: Tasks + Upcoming viewings
        pw = ttk.PanedWindow(self.dashboard_tab, orient=tk.HORIZONTAL)
        pw.pack(fill="both", expand=True, padx=10, pady=10)

        task_frame = ttk.LabelFrame(pw, text="할 일(Next Action)")
        pw.add(task_frame, weight=1)

        # ---- Tasks
        tctrl = ttk.Frame(task_frame)
        tctrl.pack(fill="x", padx=6, pady=6)
        ttk.Button(tctrl, text="새 할 일", command=self.open_add_task_window).pack(side="left", padx=4)
        ttk.Button(tctrl, text="완료", command=self.mark_selected_task_done).pack(side="left", padx=4)
        ttk.Button(tctrl, text="관련 열기", command=self.open_selected_task_related).pack(side="left", padx=4)
        ttk.Button(tctrl, text="새로고침", command=self.refresh_tasks).pack(side="left", padx=4)

        tcols = ("id", "due_at", "title", "entity", "status")
        self.tasks_tree = ttk.Treeview(task_frame, columns=tcols, show="headings", height=18)
        twidths = {"id": 55, "due_at": 170, "title": 430, "entity": 340, "status": 100}
        tlabels = {"id": "ID", "due_at": "일정일시", "title": "할 일", "entity": "고객/물건", "status": "상태"}
        for c in tcols:
            self.tasks_tree.heading(c, text=tlabels.get(c, c), command=lambda col=c: self._sort_tasks_by(col))
            self.tasks_tree.column(c, width=twidths.get(c, 120))
        self.tasks_tree.pack(fill="both", expand=True, padx=6, pady=6)
        t_x = ttk.Scrollbar(task_frame, orient="horizontal", command=self.tasks_tree.xview)
        self.tasks_tree.configure(xscrollcommand=t_x.set)
        t_x.pack(fill="x", padx=6, pady=(0,6))
        self.tasks_tree.bind("<Double-1>", self._on_double_click_task)


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

        try:
            now = datetime.now()
            end = now + timedelta(days=7)
            vcount = 0
            for v in list_viewings():
                dt = None
                try:
                    dt = datetime.fromisoformat(str(v.get("start_at") or ""))
                except Exception:
                    try:
                        dt = datetime.strptime(str(v.get("start_at") or ""), "%Y-%m-%d %H:%M")
                    except Exception:
                        dt = None
                if dt and now <= dt <= end:
                    vcount += 1
            self.dash_vars["viewings_7d"].set(str(vcount))
        except Exception:
            self.dash_vars["viewings_7d"].set("0")

    def _sort_tasks_by(self, col: str):
        if col not in {"id", "due_at", "title", "entity", "status"}:
            return
        if self.task_sort_col == col:
            self.task_sort_desc = not self.task_sort_desc
        else:
            self.task_sort_col = col
            self.task_sort_desc = False
        self.refresh_tasks()

    def _sort_task_rows(self, rows: list[dict]) -> list[dict]:
        col = getattr(self, "task_sort_col", "due_at")
        desc = bool(getattr(self, "task_sort_desc", False))

        def key_due(r: dict):
            raw = str(r.get("due_at") or "")
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
                try:
                    return datetime.strptime(raw, fmt)
                except Exception:
                    pass
            return datetime.max if not desc else datetime.min

        key_map = {
            "id": lambda r: int(r.get("id") or 0),
            "due_at": key_due,
            "title": lambda r: str(r.get("title") or ""),
            "entity": lambda r: self._task_entity_label(r),
            "status": lambda r: str(r.get("status") or ""),
        }
        return sorted(rows, key=key_map.get(col, key_due), reverse=desc)

    def _task_entity_label(self, row: dict) -> str:
        et = str(row.get("entity_type") or "").strip()
        eid = row.get("entity_id")
        if not et or eid is None:
            return ""
        try:
            eid_i = int(eid)
        except Exception:
            return f"{et}#{eid}"
        if et == "PROPERTY":
            p = get_property(eid_i)
            if p:
                return f"{p.get('complex_name') or ''} {p.get('address_detail') or ''}".strip() or f"물건#{eid_i}"
        if et == "CUSTOMER":
            c = get_customer(eid_i)
            if c:
                return f"고객 {c.get('customer_name') or eid_i}".strip()
        if et == "VIEWING":
            v = get_viewing(eid_i)
            if v and v.get("property_id"):
                p = get_property(int(v.get("property_id")))
                if p:
                    return f"일정 {p.get('complex_name') or ''} {p.get('address_detail') or ''}".strip()
            return f"일정#{eid_i}"
        return f"{et}#{eid_i}"

    def refresh_tasks(self):
        # 자동 태스크 최신화
        try:
            reconcile_auto_tasks()
        except Exception:
            pass
    
        rows = list_tasks(include_done=False, limit=400)
        rows = self._sort_task_rows(rows)
        self.dash_vars.get("tasks_open", tk.StringVar()).set(str(len(rows)))

        for i in self.tasks_tree.get_children():
            self.tasks_tree.delete(i)

        for r in rows:
            entity = self._task_entity_label(r)

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
        self.complete_task_and_handoff(tid)

    def _task_handoff_options(self, title: str) -> list[str]:
        t = (title or "").strip()
        if "상담" in t and "예약" in t:
            return ["집/상가 방문", "약속 어레인지", "광고 등록", "기타"]
        if "방문" in t:
            return ["계약 / 잔금 일정", "약속 어레인지", "기타"]
        if "계약" in t or "잔금" in t:
            return ["후속(서류/정산/보관)", "기타", "종료"]
        return ["기타"]

    def open_task_handoff_dialog(self, task_row: dict):
        title = str(task_row.get("title") or "")
        options = self._task_handoff_options(title)

        win = tk.Toplevel(self.root)
        win.title("다음 할 일 추천")
        self._fit_toplevel(win, 620, 420)

        create_next = tk.BooleanVar(value=True)
        next_type = tk.StringVar(value=options[0] if options else "기타")
        note_var = tk.StringVar(value="")

        now = datetime.now()
        y = tk.IntVar(value=now.year)
        m = tk.IntVar(value=now.month)
        d = tk.IntVar(value=now.day + 1 if now.day < 28 else now.day)
        hh = tk.IntVar(value=10)
        mm = tk.IntVar(value=0)

        selected_customer = tk.StringVar(value="")
        selected_property = tk.StringVar(value="")

        # 기존 컨텍스트 자동 채움
        et = str(task_row.get("entity_type") or "")
        eid = task_row.get("entity_id")
        mobj = re.search(r"고객\s*(\d+)", title)
        if mobj:
            selected_customer.set(mobj.group(1))
        if et == "CUSTOMER" and eid is not None:
            selected_customer.set(str(eid))
        if et == "PROPERTY" and eid is not None:
            selected_property.set(str(eid))

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="완료된 할 일").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        ttk.Label(frm, text=title or "(제목 없음)", foreground="#2f4f4f").grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Checkbutton(frm, text="다음 할 일 만들기", variable=create_next).grid(row=1, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(frm, text="다음 할 일").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        ttk.Combobox(frm, textvariable=next_type, values=options, state="readonly", width=30).grid(row=2, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(frm, text="일시").grid(row=3, column=0, sticky="e", padx=6, pady=6)
        dt = ttk.Frame(frm)
        dt.grid(row=3, column=1, sticky="w")
        ttk.Spinbox(dt, from_=2020, to=2100, textvariable=y, width=6).pack(side="left")
        ttk.Spinbox(dt, from_=1, to=12, textvariable=m, width=4).pack(side="left", padx=2)
        ttk.Spinbox(dt, from_=1, to=31, textvariable=d, width=4).pack(side="left", padx=2)
        ttk.Spinbox(dt, from_=0, to=23, textvariable=hh, width=4).pack(side="left", padx=2)
        ttk.Spinbox(dt, from_=0, to=59, increment=5, textvariable=mm, width=4).pack(side="left", padx=2)

        # 매칭(고객+물건) 선택 영역
        ttk.Label(frm, text="고객ID").grid(row=4, column=0, sticky="e", padx=6, pady=6)
        cwrap = ttk.Frame(frm)
        cwrap.grid(row=4, column=1, sticky="w", padx=6, pady=6)
        ttk.Entry(cwrap, textvariable=selected_customer, width=12).pack(side="left")

        def pick_customer():
            rows = [c for c in list_customers(include_deleted=False) if not c.get("hidden")]
            pop = tk.Toplevel(win)
            pop.title("고객 선택")
            self._fit_toplevel(pop, 760, 560)
            tree = ttk.Treeview(pop, columns=("id", "name", "phone"), show="headings", height=12)
            for c in ("id", "name", "phone"):
                tree.heading(c, text=c)
            tree.pack(fill="both", expand=True, padx=8, pady=8)
            for r in rows:
                tree.insert("", "end", values=(r.get("id"), r.get("customer_name"), r.get("phone")))
            def done():
                sel = tree.selection()
                if sel:
                    selected_customer.set(str(tree.item(sel[0], "values")[0]))
                pop.destroy()
            ttk.Button(pop, text="선택", command=done).pack(pady=6)

        ttk.Button(cwrap, text="고객선택", command=pick_customer).pack(side="left", padx=4)

        ttk.Label(frm, text="물건ID").grid(row=5, column=0, sticky="e", padx=6, pady=6)
        pwrap = ttk.Frame(frm)
        pwrap.grid(row=5, column=1, sticky="w", padx=6, pady=6)
        ttk.Entry(pwrap, textvariable=selected_property, width=12).pack(side="left")

        def pick_property():
            rows = [p for p in list_properties(include_deleted=False) if not p.get("hidden")]
            pop = tk.Toplevel(win)
            pop.title("물건 선택")
            self._fit_toplevel(pop, 860, 620)
            tree = ttk.Treeview(pop, columns=("id", "complex", "addr"), show="headings", height=12)
            for c in ("id", "complex", "addr"):
                tree.heading(c, text=c)
            tree.pack(fill="both", expand=True, padx=8, pady=8)
            for r in rows:
                tree.insert("", "end", values=(r.get("id"), r.get("complex_name"), r.get("address_detail")))
            def done():
                sel = tree.selection()
                if sel:
                    selected_property.set(str(tree.item(sel[0], "values")[0]))
                pop.destroy()
            ttk.Button(pop, text="선택", command=done).pack(pady=6)

        ttk.Button(pwrap, text="물건선택", command=pick_property).pack(side="left", padx=4)

        ttk.Label(frm, text="메모").grid(row=6, column=0, sticky="e", padx=6, pady=6)
        ttk.Entry(frm, textvariable=note_var, width=34).grid(row=6, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(frm, text="완료되었습니다. 다음 액션을 추천합니다.", foreground="#8b4513").grid(row=7, column=0, columnspan=2, sticky="w", padx=6, pady=8)

        def confirm():
            if not create_next.get():
                win.destroy()
                return
            try:
                due_at = datetime(y.get(), m.get(), d.get(), hh.get(), mm.get()).strftime("%Y-%m-%d %H:%M")
            except Exception:
                messagebox.showwarning("확인", "일시 값이 올바르지 않습니다.")
                return

            mapped_title = next_type.get().strip()
            if mapped_title == "종료":
                win.destroy()
                return

            entity_type = task_row.get("entity_type")
            entity_id = task_row.get("entity_id")
            title_to_save = mapped_title

            if mapped_title == "약속 어레인지":
                if not selected_customer.get().strip() or not selected_property.get().strip():
                    messagebox.showwarning("확인", "약속 어레인지는 고객/물건을 모두 선택해주세요.")
                    return
                entity_type = "PROPERTY"
                entity_id = int(selected_property.get())
                title_to_save = f"약속 어레인지 (고객 {selected_customer.get().strip()})"
            elif mapped_title in ("광고 등록", "후속(서류/정산/보관)") and selected_property.get().strip():
                entity_type = "PROPERTY"
                entity_id = int(selected_property.get())

            try:
                add_task(
                    title=title_to_save,
                    due_at=due_at,
                    entity_type=entity_type,
                    entity_id=entity_id,
                    note=note_var.get().strip(),
                    kind="MANUAL",
                    status="OPEN",
                )
            except Exception as exc:
                messagebox.showerror("오류", f"다음 할 일 생성 실패: {exc}")
                return
            self.refresh_tasks()
            self.refresh_dashboard()
            win.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=1, sticky="w", padx=6, pady=10)
        ttk.Button(btns, text="확인", command=confirm).pack(side="left", padx=4)
        ttk.Button(btns, text="닫기", command=win.destroy).pack(side="left", padx=4)

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
        tid = self._selected_task_id()
        if tid is None:
            return
        self.open_task_detail(tid)

    def complete_task_and_handoff(self, task_id: int, *, parent_win: tk.Toplevel | None = None):
        row = None
        try:
            tasks = list_tasks(include_done=True, limit=800)
            for t in tasks:
                if int(t.get("id") or 0) == int(task_id):
                    row = t
                    break
        except Exception:
            row = None

        try:
            mark_task_done(task_id)
        except Exception as exc:
            messagebox.showerror("오류", f"완료 처리 실패: {exc}")
            return

        self.refresh_tasks()
        self.refresh_dashboard()
        if parent_win is not None:
            try:
                parent_win.destroy()
            except Exception:
                pass

        messagebox.showinfo("완료", "할 일을 완료했습니다.")
        if row:
            self.open_task_handoff_dialog(row)

    def open_task_detail(self, task_id: int):
        row = get_task(task_id)
        if not row:
            messagebox.showwarning("안내", "할 일 정보를 찾을 수 없습니다.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"할 일 상세 - #{task_id}")
        self._fit_toplevel(win, 560, 360)

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="제목").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        ttk.Label(frm, text=str(row.get("title") or "")).grid(row=0, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(frm, text="일시").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        ttk.Label(frm, text=str(row.get("due_at") or "")).grid(row=1, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(frm, text="상태").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        ttk.Label(frm, text=str(row.get("status") or "")).grid(row=2, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(frm, text="메모").grid(row=3, column=0, sticky="ne", padx=6, pady=6)
        ttk.Label(frm, text=str(row.get("note") or "")).grid(row=3, column=1, sticky="w", padx=6, pady=6)

        def open_related_from_detail():
            et = str(row.get("entity_type") or "").strip()
            eid = row.get("entity_id")
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

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=1, sticky="w", padx=6, pady=12)
        ttk.Button(btns, text="관련 열기", command=open_related_from_detail).pack(side="left", padx=4)
        if str(row.get("status") or "") == "OPEN":
            ttk.Button(btns, text="완료", command=lambda: self.complete_task_and_handoff(task_id, parent_win=win)).pack(side="left", padx=4)
        else:
            ttk.Label(btns, text="이미 완료된 할 일입니다.", foreground="#8b4513").pack(side="left", padx=4)
        ttk.Button(btns, text="닫기", command=win.destroy).pack(side="left", padx=4)

    def open_add_task_window(self, *, default_entity_type: str = "", default_entity_id: int | None = None):
        win = tk.Toplevel(self.root)
        win.title("새 할 일 추가")
        self._fit_toplevel(win, 700, 520)

        selected_properties: list[int] = []
        selected_customer: int | None = None
        if default_entity_type == "PROPERTY" and default_entity_id:
            selected_properties = [int(default_entity_id)]
        elif default_entity_type == "CUSTOMER" and default_entity_id:
            selected_customer = int(default_entity_id)

        task_type_var = tk.StringVar(value="상담 예약")
        target_type_var = tk.StringVar(value="물건")
        note_var = tk.StringVar(value="")
        now = datetime.now()
        date_vars = {
            "y": tk.IntVar(value=now.year),
            "m": tk.IntVar(value=now.month),
            "d": tk.IntVar(value=now.day),
            "hh": tk.IntVar(value=now.hour),
            "mm": tk.IntVar(value=(now.minute // 10) * 10),
        }
        selected_summary = tk.StringVar(value="고객: - / 물건: -")
        arrange_guide_var = tk.StringVar(value="")

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="할 일 종류").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        task_type_cb = ttk.Combobox(
            frm,
            textvariable=task_type_var,
            state="readonly",
            width=36,
            values=TASK_TYPE_VALUES,
        )
        task_type_cb.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(frm, text="일시").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        dt = ttk.Frame(frm)
        dt.grid(row=1, column=1, sticky="w")
        ttk.Spinbox(dt, from_=2020, to=2100, textvariable=date_vars["y"], width=6).pack(side="left")
        ttk.Spinbox(dt, from_=1, to=12, textvariable=date_vars["m"], width=4).pack(side="left", padx=2)
        ttk.Spinbox(dt, from_=1, to=31, textvariable=date_vars["d"], width=4).pack(side="left", padx=2)
        ttk.Spinbox(dt, from_=0, to=23, textvariable=date_vars["hh"], width=4).pack(side="left", padx=2)
        ttk.Spinbox(dt, from_=0, to=59, increment=5, textvariable=date_vars["mm"], width=4).pack(side="left", padx=2)

        quick = ttk.Frame(frm)
        quick.grid(row=2, column=1, sticky="w", padx=6, pady=2)

        def set_quick(days: int):
            t = datetime.now() + timedelta(days=days)
            date_vars["y"].set(t.year)
            date_vars["m"].set(t.month)
            date_vars["d"].set(t.day)

        ttk.Button(quick, text="오늘", command=lambda: set_quick(0)).pack(side="left", padx=2)
        ttk.Button(quick, text="내일", command=lambda: set_quick(1)).pack(side="left", padx=2)
        ttk.Button(quick, text="+7일", command=lambda: set_quick(7)).pack(side="left", padx=2)

        ttk.Label(frm, text="연결대상").grid(row=3, column=0, sticky="e", padx=6, pady=6)
        ttk.Combobox(frm, textvariable=target_type_var, state="readonly", width=36, values=["물건", "고객", "매칭(고객+물건)"]).grid(row=3, column=1, sticky="w", padx=6, pady=6)

        def render_summary():
            selected_summary.set(f"고객: {selected_customer or '-'} / 물건: {', '.join(map(str, selected_properties)) or '-'}")

        def update_arrange_guide(_=None):
            if task_type_var.get().strip() == "약속 어레인지":
                arrange_guide_var.set("안내: 1) 고객 1명 선택 → 2) 물건 1개 이상 체크 → 저장 시 물건 수만큼 할 일이 생성됩니다.")
            else:
                arrange_guide_var.set("")

        def choose_property():
            nonlocal selected_properties
            rows = [p for p in list_properties(include_deleted=False) if not p.get("hidden")]
            if not rows:
                popup = tk.Toplevel(win)
                popup.title("물건 선택")
                self._fit_toplevel(popup, 520, 220)
                ttk.Label(popup, text="등록된 물건이 없습니다. 먼저 물건 등록을 해주세요.").pack(padx=12, pady=12, anchor="w")
                btns = ttk.Frame(popup)
                btns.pack(pady=8)
                ttk.Button(btns, text="물건 등록 열기", command=lambda: (popup.destroy(), self.open_property_wizard())).pack(side="left", padx=4)
                ttk.Button(btns, text="닫기", command=popup.destroy).pack(side="left", padx=4)
                return

            popup = tk.Toplevel(win)
            popup.title("물건 선택")
            self._fit_toplevel(popup, 820, 620)
            search_var = tk.StringVar(value="")
            selected_count_var = tk.StringVar(value=f"선택 {len(selected_properties)}건")
            picked_ids = set(selected_properties)

            ttk.Entry(popup, textvariable=search_var, width=40).pack(padx=8, pady=4, anchor="w")
            tree = ttk.Treeview(popup, columns=("pick", "id", "tab", "addr"), show="headings")
            tree.heading("pick", text="선택")
            tree.heading("id", text="번호")
            tree.heading("tab", text="단지")
            tree.heading("addr", text="주소")
            tree.column("pick", width=56, anchor="center")
            tree.column("id", width=64, anchor="center")
            tree.column("tab", width=180)
            tree.column("addr", width=260)
            tree.pack(fill="both", expand=True, padx=8, pady=8)

            item_to_id: dict[str, int] = {}

            def render():
                item_to_id.clear()
                for item in tree.get_children():
                    tree.delete(item)
                q = search_var.get().strip().lower()
                for r in rows:
                    hay = f"{r.get('tab','')} {r.get('address_detail','')} {r.get('owner_phone','')}".lower()
                    if q and q not in hay:
                        continue
                    pid = int(r.get("id") or 0)
                    mark = "☑" if pid in picked_ids else "☐"
                    iid = tree.insert("", "end", values=(mark, pid, r.get("tab"), r.get("address_detail")))
                    item_to_id[iid] = pid
                selected_count_var.set(f"선택 {len(picked_ids)}건")

            def toggle_item(item: str):
                pid = item_to_id.get(item)
                if not pid:
                    return
                if pid in picked_ids:
                    picked_ids.remove(pid)
                else:
                    picked_ids.add(pid)
                mark = "☑" if pid in picked_ids else "☐"
                vals = list(tree.item(item, "values"))
                vals[0] = mark
                tree.item(item, values=vals)
                selected_count_var.set(f"선택 {len(picked_ids)}건")

            def on_click(event):
                item = tree.identify_row(event.y)
                if item:
                    toggle_item(item)

            search_var.trace_add("write", lambda *_: render())
            tree.bind("<Button-1>", on_click)
            render()

            bottom = ttk.Frame(popup)
            bottom.pack(fill="x", padx=8, pady=6)
            ttk.Label(bottom, textvariable=selected_count_var, foreground="#2f4f4f").pack(side="left")

            def done():
                nonlocal selected_properties
                selected_properties = sorted(picked_ids)
                render_summary()
                popup.destroy()

            ttk.Button(bottom, text="선택 완료", command=done).pack(side="right")

        def choose_customer():
            nonlocal selected_customer
            rows = [c for c in list_customers(include_deleted=False) if not c.get("hidden") and c.get("status") in ("문의", "임장예약", "계약진행", "대기", "")]
            popup = tk.Toplevel(win)
            popup.title("고객 선택")
            search_var = tk.StringVar(value="")
            ttk.Entry(popup, textvariable=search_var, width=36).pack(padx=8, pady=4, anchor="w")
            tree = ttk.Treeview(popup, columns=("id", "name", "phone"), show="headings")
            for c in ("id", "name", "phone"):
                tree.heading(c, text=c)
            tree.pack(fill="both", expand=True, padx=8, pady=8)

            def render():
                for item in tree.get_children():
                    tree.delete(item)
                q = search_var.get().strip()
                qd = "".join(ch for ch in q if ch.isdigit())
                for r in rows:
                    phone = "".join(ch for ch in str(r.get("phone") or "") if ch.isdigit())
                    hay = f"{r.get('customer_name','')} {r.get('phone','')}".lower()
                    if q and q.lower() not in hay and (not qd or qd not in phone):
                        continue
                    tree.insert("", "end", values=(r.get("id"), r.get("customer_name"), r.get("phone")))

            search_var.trace_add("write", lambda *_: render())
            render()

            def done():
                nonlocal selected_customer
                sel = tree.selection()
                if not sel:
                    return
                selected_customer = int(tree.item(sel[0], "values")[0])
                render_summary()
                popup.destroy()

            ttk.Button(popup, text="선택 완료", command=done).pack(pady=6)

        sel_btns = ttk.Frame(frm)
        sel_btns.grid(row=4, column=1, sticky="w", padx=6, pady=6)
        ttk.Button(sel_btns, text="고객 선택", command=choose_customer).pack(side="left", padx=2)
        ttk.Button(sel_btns, text="물건 선택", command=choose_property).pack(side="left", padx=2)
        ttk.Label(frm, textvariable=selected_summary, foreground="#555").grid(row=5, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(frm, textvariable=arrange_guide_var, foreground="#8b4513").grid(row=6, column=1, sticky="w", padx=6, pady=2)

        ttk.Label(frm, text="메모").grid(row=7, column=0, sticky="e", padx=6, pady=6)
        ttk.Entry(frm, textvariable=note_var, width=42).grid(row=7, column=1, sticky="w", padx=6, pady=6)

        def save():
            title = task_type_var.get().strip()
            try:
                due_at = datetime(
                    date_vars["y"].get(),
                    date_vars["m"].get(),
                    date_vars["d"].get(),
                    date_vars["hh"].get(),
                    date_vars["mm"].get(),
                ).strftime("%Y-%m-%d %H:%M")
            except Exception:
                messagebox.showwarning("확인", "일시 값이 올바르지 않습니다.")
                return

            if title == "상담 예약" and not (selected_properties or selected_customer):
                messagebox.showwarning("확인", "상담 예약은 고객 또는 물건을 선택해야 합니다.")
                return
            if title == "집/상가 방문" and not selected_properties:
                messagebox.showwarning("확인", "방문 일정은 물건 1개 이상이 필요합니다.")
                return
            if title == "계약 / 잔금 일정" and not selected_properties:
                messagebox.showwarning("확인", "계약/잔금 일정은 물건 1개 이상이 필요합니다.")
                return
            if title in ("광고 등록", "후속(서류/정산/보관)") and not selected_properties:
                messagebox.showwarning("확인", f"{title} 할 일은 물건 1개 이상이 필요합니다.")
                return
            if title == "약속 어레인지":
                if not selected_customer or not selected_properties:
                    messagebox.showwarning("확인", "약속 어레인지는 고객 1명 + 물건 1개 이상이 필요합니다.")
                    return
                for pid in selected_properties:
                    add_task(title=f"{title} (고객 {selected_customer})", due_at=due_at, entity_type="PROPERTY", entity_id=pid, note=note_var.get().strip(), kind="MANUAL", status="OPEN")
            elif selected_properties and title in ("집/상가 방문", "계약 / 잔금 일정", "광고 등록", "후속(서류/정산/보관)", "기타"):
                for pid in selected_properties:
                    add_task(title=title, due_at=due_at, entity_type="PROPERTY", entity_id=pid, note=note_var.get().strip(), kind="MANUAL", status="OPEN")
            else:
                entity_type = None
                entity_id = None
                if selected_properties:
                    entity_type, entity_id = "PROPERTY", selected_properties[0]
                elif selected_customer:
                    entity_type, entity_id = "CUSTOMER", selected_customer
                add_task(title=title, due_at=due_at, entity_type=entity_type, entity_id=entity_id, note=note_var.get().strip(), kind="MANUAL", status="OPEN")

            self.refresh_tasks()
            self.refresh_dashboard()
            win.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=1, sticky="w", padx=6, pady=12)
        ttk.Button(btns, text="저장", command=save).pack(side="left", padx=4)
        ttk.Button(btns, text="취소", command=win.destroy).pack(side="left", padx=4)

        task_type_cb.bind("<<ComboboxSelected>>", update_arrange_guide)
        render_summary()
        update_arrange_guide()

    def _build_property_ui(self):
        top = ttk.LabelFrame(self.property_tab, text="물건")
        top.pack(fill="x", padx=10, pady=8)

        ttk.Button(top, text="+ 물건 등록", command=self.open_property_wizard).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="내보내기/동기화", command=self.export_sync).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="숨김함", command=self.open_hidden_properties_window).pack(side="left", padx=4, pady=6)

        ttk.Label(top, text="검색").pack(side="left", padx=(24, 4))
        self.property_search_var = tk.StringVar(value="")
        ent = ttk.Entry(top, textvariable=self.property_search_var, width=30)
        ent.pack(side="left", padx=4)
        ent.bind("<KeyRelease>", lambda _e: self.refresh_properties())

        self.inner_tabs = ttk.Notebook(self.property_tab)
        self.inner_tabs.pack(fill="both", expand=True, padx=10, pady=8)
        self.prop_trees: dict[str, ttk.Treeview] = {}

        cols = ("status", "complex_name", "address_detail", "unit_type", "floor", "price_summary", "updated_at", "id")
        col_defs = [
            ("status", 90),
            ("complex_name", 180),
            ("address_detail", 170),
            ("unit_type", 100),
            ("floor", 70),
            ("price_summary", 250),
            ("updated_at", 145),
            ("id", 1),
        ]
        col_labels = {
            "status": "상태",
            "complex_name": "단지",
            "address_detail": "동/호",
            "unit_type": "타입",
            "floor": "층",
            "price_summary": "가격요약",
            "updated_at": "업데이트",
            "id": "",
        }

        for tab_name in PROPERTY_TABS:
            frame = ttk.Frame(self.inner_tabs)
            self.inner_tabs.add(frame, text=tab_name)

            tree = ttk.Treeview(frame, columns=cols, show="headings", height=17)
            for c, w in col_defs:
                tree.heading(c, text=col_labels.get(c, c), command=lambda col=c, t=tree: self.sort_tree(t, col))
                tree.column(c, width=w, stretch=(c != "id"))
                if c == "id":
                    tree.column(c, minwidth=0, width=0, stretch=False)
            tree.pack(fill="both", expand=True)
            xsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=xsb.set)
            xsb.pack(fill="x")
            tree.bind("<Double-1>", lambda e, t=tab_name: self._on_double_click_property(e, t))
            self.prop_trees[tab_name] = tree

            btns = ttk.Frame(frame)
            btns.pack(fill="x", pady=4)
            ttk.Button(btns, text="상세", command=lambda t=tab_name: self.open_selected_property_detail(t)).pack(side="left", padx=4)
            ttk.Button(btns, text="숨김/보임", command=lambda t=tab_name: self.toggle_selected_property(t)).pack(side="left", padx=4)
            ttk.Button(btns, text="삭제", command=lambda t=tab_name: self.delete_selected_property(t)).pack(side="left", padx=4)

    def sort_tree(self, tree: ttk.Treeview, col: str):
        key = (id(tree), col)
        asc = not self.sort_state.get(key, False)
        items = list(tree.get_children(""))

        def val(item: str):
            raw = str(tree.set(item, col) or "")
            if col in {"id", "floor"}:
                try:
                    return int(raw)
                except Exception:
                    return 0
            if col == "updated_at":
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
                    try:
                        return datetime.strptime(raw, fmt)
                    except Exception:
                        pass
                return datetime.min
            return raw

        items.sort(key=val, reverse=not asc)
        for idx, item in enumerate(items):
            tree.move(item, "", idx)
        self.sort_state[key] = asc

    def _calc_price_summary(self, row: dict) -> str:
        return money_utils.property_price_summary(row)

    def refresh_properties(self):
        q = self.property_search_var.get().strip().lower() if hasattr(self, "property_search_var") else ""
        for tab in PROPERTY_TABS:
            tree = self.prop_trees[tab]
            for i in tree.get_children():
                tree.delete(i)
            rows = list_properties(tab)
            rows.sort(key=lambda r: str(r.get("updated_at") or ""), reverse=True)
            for row in rows:
                if row.get("hidden"):
                    continue
                if q:
                    hay = " ".join([str(row.get("address_detail") or ""), str(row.get("owner_phone") or ""), str(row.get("special_notes") or "")]).lower()
                    if q not in hay:
                        continue
                tree.insert(
                    "",
                    "end",
                    values=(
                        row.get("status"),
                        row.get("complex_name"),
                        row.get("address_detail"),
                        row.get("unit_type"),
                        row.get("floor"),
                        self._calc_price_summary(row),
                        row.get("updated_at"),
                        row.get("id"),
                    ),
                )

    def open_property_wizard(self):
        win = tk.Toplevel(self.root)
        win.title("물건 등록")
        self._fit_toplevel(win, 860, 700)

        vars_ = {
            "tab": tk.StringVar(value=PROPERTY_TABS[0]),
            "unit_type": tk.StringVar(value=""),
            "dong": tk.StringVar(value=""),
            "floor": tk.StringVar(value=""),
            "ho": tk.StringVar(value=""),
            "area": tk.StringVar(value=""),
            "pyeong": tk.StringVar(value=""),
            "deal_sale": tk.BooleanVar(value=False),
            "deal_jeonse": tk.BooleanVar(value=False),
            "deal_wolse": tk.BooleanVar(value=False),
            "price_sale_10m": tk.StringVar(value="0"),
            "price_jeonse_10m": tk.StringVar(value="0"),
            "wolse_deposit_10m": tk.StringVar(value="0"),
            "wolse_rent_10man": tk.StringVar(value="0"),
            "condition": tk.StringVar(value="중"),
            "view": tk.StringVar(value="탁 트인 뷰"),
            "orientation": tk.StringVar(value="남향"),
            "status": tk.StringVar(value="신규등록"),
            "repair_needed": tk.BooleanVar(value=False),
            "repair_items": tk.StringVar(value=""),
            "owner_name": tk.StringVar(value=""),
            "owner_phone": tk.StringVar(value=""),
            "owner_status": tk.StringVar(value=""),
            "resident_type": tk.StringVar(value="주인거주"),
            "tenant_phone": tk.StringVar(value=""),
            "visit_coop": tk.StringVar(value="협조"),
            "contact_coop": tk.StringVar(value="협조"),
            "visit_condition": tk.StringVar(value="미리 약속 필요"),
            "move_available_date": tk.StringVar(value=""),
            "special_notes": tk.StringVar(value=""),
            "note": tk.StringVar(value=""),
        }

        header = ttk.Frame(win)
        header.pack(fill="x", padx=10, pady=(10, 4))
        progress_var = tk.StringVar(value="1/3 단계: 기본 정보")
        ttk.Label(header, textvariable=progress_var).pack(anchor="w")

        body = ttk.Frame(win)
        body.pack(fill="both", expand=True, padx=10, pady=6)

        step1 = ttk.Frame(body)
        step2 = ttk.Frame(body)
        step3 = ttk.Frame(body)
        steps = [step1, step2, step3]
        step_titles = ["기본 정보", "거래/가격", "연락/메모"]

        # Step1
        ttk.Label(step1, text="단지").grid(row=0, column=0, padx=6, pady=6, sticky="e")
        tab_cb = ttk.Combobox(step1, textvariable=vars_["tab"], values=PROPERTY_TABS, state="readonly", width=26)
        tab_cb.grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step1, text="동").grid(row=1, column=0, padx=6, pady=6, sticky="e")
        dong_cb = ttk.Combobox(step1, textvariable=vars_["dong"], state="readonly", width=26)
        dong_cb.grid(row=1, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step1, text="층").grid(row=2, column=0, padx=6, pady=6, sticky="e")
        floor_cb = ttk.Combobox(step1, textvariable=vars_["floor"], state="readonly", width=26)
        floor_cb.grid(row=2, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step1, text="호").grid(row=3, column=0, padx=6, pady=6, sticky="e")
        ho_cb = ttk.Combobox(step1, textvariable=vars_["ho"], state="readonly", width=26)
        ho_cb.grid(row=3, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step1, text="면적타입").grid(row=4, column=0, padx=6, pady=6, sticky="e")
        ttk.Label(step1, textvariable=vars_["unit_type"]).grid(row=4, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step1, text="면적(㎡)").grid(row=5, column=0, padx=6, pady=6, sticky="e")
        ttk.Label(step1, textvariable=vars_["area"]).grid(row=5, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step1, text="평형").grid(row=6, column=0, padx=6, pady=6, sticky="e")
        ttk.Label(step1, textvariable=vars_["pyeong"]).grid(row=6, column=1, padx=6, pady=6, sticky="w")

        manual_hint = tk.StringVar(value="")
        ttk.Label(step1, textvariable=manual_hint, foreground="#555").grid(row=7, column=0, columnspan=2, padx=6, pady=2, sticky="w")

        # Step2
        ttk.Checkbutton(step2, text="매매", variable=vars_["deal_sale"]).grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(step2, textvariable=vars_["price_sale_10m"], width=12).grid(row=0, column=1)
        ttk.Label(step2, text="천만원").grid(row=0, column=2)

        ttk.Checkbutton(step2, text="전세", variable=vars_["deal_jeonse"]).grid(row=1, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(step2, textvariable=vars_["price_jeonse_10m"], width=12).grid(row=1, column=1)
        ttk.Label(step2, text="천만원").grid(row=1, column=2)

        ttk.Checkbutton(step2, text="월세", variable=vars_["deal_wolse"]).grid(row=2, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(step2, textvariable=vars_["wolse_deposit_10m"], width=12).grid(row=2, column=1)
        ttk.Label(step2, text="천만원").grid(row=2, column=2)
        ttk.Entry(step2, textvariable=vars_["wolse_rent_10man"], width=12).grid(row=2, column=3)
        ttk.Label(step2, text="십만원").grid(row=2, column=4)

        ttk.Label(step2, text="상태").grid(row=3, column=0, padx=6, pady=6, sticky="e")
        ttk.Combobox(step2, textvariable=vars_["status"], values=PROPERTY_STATUS_VALUES, state="readonly", width=20).grid(row=3, column=1, columnspan=2, padx=6, pady=6, sticky="w")
        ttk.Label(step2, text="컨디션").grid(row=4, column=0, padx=6, pady=6, sticky="e")
        ttk.Combobox(step2, textvariable=vars_["condition"], values=["상", "중", "하"], state="readonly", width=20).grid(row=4, column=1, columnspan=2, padx=6, pady=6, sticky="w")

        # Step3
        ttk.Label(step3, text="집주인명").grid(row=0, column=0, padx=6, pady=6, sticky="e")
        ttk.Entry(step3, textvariable=vars_["owner_name"], width=26).grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step3, text="집주인 전화*").grid(row=1, column=0, padx=6, pady=6, sticky="e")
        ttk.Entry(step3, textvariable=vars_["owner_phone"], width=26).grid(row=1, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step3, text="세입자 전화").grid(row=2, column=0, padx=6, pady=6, sticky="e")
        ttk.Entry(step3, textvariable=vars_["tenant_phone"], width=26).grid(row=2, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step3, text="입주 가능일").grid(row=3, column=0, padx=6, pady=6, sticky="e")

        move_mode = tk.StringVar(value="협의")
        move_y = tk.IntVar(value=datetime.now().year)
        move_m = tk.IntVar(value=datetime.now().month)
        move_d = tk.IntVar(value=datetime.now().day)
        move_wrap = ttk.Frame(step3)
        move_wrap.grid(row=3, column=1, padx=6, pady=6, sticky="w")
        move_mode_cb = ttk.Combobox(move_wrap, textvariable=move_mode, values=["즉시", "협의", "날짜선택"], state="readonly", width=10)
        move_mode_cb.pack(side="left")
        date_wrap = ttk.Frame(move_wrap)
        ttk.Spinbox(date_wrap, from_=2020, to=2100, textvariable=move_y, width=6).pack(side="left")
        ttk.Spinbox(date_wrap, from_=1, to=12, textvariable=move_m, width=4).pack(side="left", padx=2)
        ttk.Spinbox(date_wrap, from_=1, to=31, textvariable=move_d, width=4).pack(side="left", padx=2)

        def refresh_move_ui(*_a):
            if move_mode.get() == "날짜선택":
                if not date_wrap.winfo_ismapped():
                    date_wrap.pack(side="left", padx=6)
            else:
                if date_wrap.winfo_ismapped():
                    date_wrap.pack_forget()

        move_mode_cb.bind("<<ComboboxSelected>>", refresh_move_ui)
        refresh_move_ui()

        ttk.Label(step3, text="메모").grid(row=4, column=0, padx=6, pady=6, sticky="ne")
        ttk.Entry(step3, textvariable=vars_["special_notes"], width=40).grid(row=4, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(step3, text="사진은 저장 후 상세창 [사진] 탭에서 업로드", foreground="#555").grid(row=5, column=0, columnspan=2, padx=6, pady=6, sticky="w")

        def reset_auto_values():
            vars_["unit_type"].set("")
            vars_["area"].set("")
            vars_["pyeong"].set("")

        def refresh_unit_info():
            tab = vars_["tab"].get().strip()
            if not unit_master.has_master(tab):
                reset_auto_values()
                return
            dong = vars_["dong"].get().strip()
            ho = vars_["ho"].get().strip()
            floor_num = int(vars_["floor"].get() or 0)
            info = unit_master.get_unit_info(tab, dong, floor_num, ho)
            vars_["unit_type"].set(str(info.get("type") or ""))
            area = float(info.get("supply_m2") or 0)
            pyeong = float(info.get("pyeong") or 0)
            vars_["area"].set(f"{area:g}" if area else "")
            vars_["pyeong"].set(f"{pyeong:g}" if pyeong else "")

        def on_floor_change(_=None):
            tab = vars_["tab"].get().strip()
            if not unit_master.has_master(tab):
                return
            dong = vars_["dong"].get().strip()
            floor = vars_["floor"].get().strip()
            floors = unit_master.get_floors(tab, dong)
            floor_cb.configure(values=[str(f) for f in floors])
            if not floor and floors:
                floor = str(floors[0])
                vars_["floor"].set(floor)
            hos = unit_master.get_hos(tab, dong, int(floor) if floor else 0)
            ho_cb.configure(values=hos)
            vars_["ho"].set(hos[0] if hos else "")
            refresh_unit_info()

        def on_dong_change(_=None):
            tab = vars_["tab"].get().strip()
            if not unit_master.has_master(tab):
                return
            dong = vars_["dong"].get().strip()
            floors = unit_master.get_floors(tab, dong)
            floor_cb.configure(values=[str(f) for f in floors])
            vars_["floor"].set(str(floors[0]) if floors else "")
            on_floor_change()

        def set_manual_mode(hint: str):
            manual_hint.set(hint)
            dong_cb.configure(state="normal", values=[])
            floor_cb.configure(state="normal", values=[])
            ho_cb.configure(state="normal", values=[])
            vars_["dong"].set("")
            vars_["floor"].set("")
            vars_["ho"].set("")
            reset_auto_values()

        def on_tab_change(_=None):
            tab = vars_["tab"].get().strip()
            if unit_master.has_master(tab):
                dongs = unit_master.get_dongs(tab)
                if not dongs:
                    set_manual_mode("CSV 데이터가 비어 있어 수동 입력 모드로 전환했습니다.")
                    return
                manual_hint.set("아파트는 동→층→호만 선택하면 면적/타입이 자동 입력됩니다.")
                dong_cb.configure(state="readonly", values=dongs)
                floor_cb.configure(state="readonly")
                ho_cb.configure(state="readonly")
                vars_["dong"].set(dongs[0])
                on_dong_change()
            else:
                set_manual_mode("상가/단독주택은 동/층/호를 직접 입력하세요.")

        tab_cb.bind("<<ComboboxSelected>>", on_tab_change)
        dong_cb.bind("<<ComboboxSelected>>", on_dong_change)
        floor_cb.bind("<<ComboboxSelected>>", on_floor_change)
        ho_cb.bind("<<ComboboxSelected>>", lambda _e: refresh_unit_info())
        on_tab_change()

        def _is_non_negative_int(v: str) -> bool:
            v = v.strip()
            return v.isdigit() if v else False

        def validate_step(idx: int) -> bool:
            tab = vars_["tab"].get().strip()
            if idx == 0:
                if unit_master.has_master(tab):
                    if not (vars_["dong"].get().strip() and vars_["floor"].get().strip() and vars_["ho"].get().strip()):
                        messagebox.showwarning("입력 확인", "아파트 물건은 단지/동/층/호를 모두 선택해주세요.")
                        return False
                return True
            if idx == 1:
                if not (vars_["deal_sale"].get() or vars_["deal_jeonse"].get() or vars_["deal_wolse"].get()):
                    messagebox.showwarning("입력 확인", "거래유형을 최소 1개 선택해주세요.")
                    return False
                nums = []
                if vars_["deal_sale"].get():
                    nums += [vars_["price_sale_10m"].get()]
                if vars_["deal_jeonse"].get():
                    nums += [vars_["price_jeonse_10m"].get()]
                if vars_["deal_wolse"].get():
                    nums += [vars_["wolse_deposit_10m"].get(), vars_["wolse_rent_10man"].get()]
                if any(not _is_non_negative_int(n) for n in nums):
                    messagebox.showwarning("입력 확인", "가격 입력은 숫자(0 포함)만 가능합니다.")
                    return False
                return True
            if idx == 2:
                if not vars_["owner_phone"].get().strip():
                    messagebox.showwarning("입력 확인", "집주인 전화번호(필수)를 입력해주세요.")
                    return False
                return True
            return True

        current = {"idx": 0}

        def show_step(idx: int):
            for frm in steps:
                frm.pack_forget()
            steps[idx].pack(fill="both", expand=True)
            progress_var.set(f"{idx + 1}/3 단계: {step_titles[idx]}")
            prev_btn.configure(state=("normal" if idx > 0 else "disabled"))
            next_btn.configure(state=("normal" if idx < len(steps) - 1 else "disabled"))
            save_btn.configure(state=("normal" if idx == len(steps) - 1 else "disabled"))

        def go_prev():
            if current["idx"] > 0:
                current["idx"] -= 1
                show_step(current["idx"])

        def go_next():
            idx = current["idx"]
            if not validate_step(idx):
                return
            if idx < len(steps) - 1:
                current["idx"] += 1
                show_step(current["idx"])

        def save():
            if not validate_step(current["idx"]):
                return

            tab = vars_["tab"].get().strip()
            dong = vars_["dong"].get().strip()
            floor = vars_["floor"].get().strip()
            ho = vars_["ho"].get().strip()

            data = {k: (v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get())) for k, v in vars_.items()}
            mode = move_mode.get().strip()
            data["move_available_date"] = f"{move_y.get():04d}-{move_m.get():02d}-{move_d.get():02d}" if mode == "날짜선택" else mode
            data["tab"] = tab
            data["complex_name"] = tab

            sale_eok, sale_che = money_utils.ten_million_to_eok_che(vars_["price_sale_10m"].get())
            jeonse_eok, jeonse_che = money_utils.ten_million_to_eok_che(vars_["price_jeonse_10m"].get())
            wolse_eok, wolse_che = money_utils.ten_million_to_eok_che(vars_["wolse_deposit_10m"].get())
            data["price_sale_eok"], data["price_sale_che"] = sale_eok, sale_che
            data["price_jeonse_eok"], data["price_jeonse_che"] = jeonse_eok, jeonse_che
            data["wolse_deposit_eok"], data["wolse_deposit_che"] = wolse_eok, wolse_che
            data["wolse_rent_man"] = money_utils.ten_man_to_man(vars_["wolse_rent_10man"].get())

            if unit_master.has_master(tab):
                floor_num = int(floor or 0)
                info = unit_master.get_unit_info(tab, dong, floor_num, ho)
                if not info.get("type"):
                    messagebox.showwarning("입력 확인", "선택한 동/층/호에 대한 타입 정보를 찾지 못했습니다.")
                    current["idx"] = 0
                    show_step(0)
                    return
                data["unit_type"] = str(info.get("type") or "")
                data["area"] = info.get("supply_m2") or 0
                data["pyeong"] = info.get("pyeong") or 0
                data["floor"] = str(floor_num) if floor_num else ""
                total_floor = unit_master.get_total_floor(tab, dong)
                data["total_floor"] = str(total_floor) if total_floor else ""
                data["address_detail"] = f"{dong} {ho}호" if dong and ho else ""
            else:
                data["address_detail"] = f"{dong} {ho}호" if dong and ho else ""

            try:
                add_property(data)
            except Exception as exc:
                messagebox.showerror("저장 실패", str(exc))
                return

            self.refresh_all()
            messagebox.showinfo("저장 완료", "물건이 저장되었습니다.")
            win.destroy()

        footer = ttk.Frame(win)
        footer.pack(fill="x", padx=10, pady=10)
        prev_btn = ttk.Button(footer, text="이전", command=go_prev)
        next_btn = ttk.Button(footer, text="다음", command=go_next)
        save_btn = ttk.Button(footer, text="저장", command=save)
        cancel_btn = ttk.Button(footer, text="취소", command=win.destroy)
        prev_btn.pack(side="left", padx=4)
        next_btn.pack(side="left", padx=4)
        save_btn.pack(side="left", padx=4)
        cancel_btn.pack(side="left", padx=4)

        show_step(0)

    def _selected_id_from_tree(self, tree: ttk.Treeview, *, value_index: int = -1) -> int | None:
        selected = tree.selection()
        if not selected:
            return None
        try:
            values = list(tree.item(selected[0], "values"))
            if not values:
                return None
            return int(values[value_index])
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
    def open_hidden_properties_window(self):
        win = tk.Toplevel(self.root)
        win.title("물건 숨김함")
        self._fit_toplevel(win, 760, 560)
        tree = ttk.Treeview(win, columns=("id", "tab", "address"), show="headings", height=14)
        for c in ("id", "tab", "address"):
            tree.heading(c, text=c)
        tree.pack(fill="both", expand=True, padx=8, pady=8)
        hidden_rows = [r for r in list_properties(include_deleted=False) if r.get("hidden")]
        for r in hidden_rows:
            tree.insert("", "end", values=(r.get("id"), r.get("tab"), r.get("address_detail")))

        def _selected_id():
            sel = tree.selection()
            if not sel:
                return None
            try:
                return int(tree.item(sel[0], "values")[0])
            except Exception:
                return None

        def restore():
            pid = _selected_id()
            if pid is None:
                return
            toggle_property_hidden(pid)
            self.refresh_properties()
            win.destroy()

        def open_detail():
            pid = _selected_id()
            if pid is None:
                return
            self.open_property_detail(pid)

        def delete_it():
            pid = _selected_id()
            if pid is None:
                return
            soft_delete_property(pid)
            self.refresh_all()
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(pady=6)
        ttk.Button(btns, text="복구", command=restore).pack(side="left", padx=4)
        ttk.Button(btns, text="상세", command=open_detail).pack(side="left", padx=4)
        ttk.Button(btns, text="삭제", command=delete_it).pack(side="left", padx=4)

    def _parse_preferred_tabs(self, text: str) -> list[str]:
        return [t.strip() for t in str(text or "").split(",") if t.strip()]

    def _normalize_property_tab(self, text: str) -> str:
        t = str(text or "").strip()
        alias = {
            "아파트단지1": "봉담자이 프라이드시티",
            "아파트단지2": "힐스테이트봉담프라이드시티",
            "힐스": "힐스테이트봉담프라이드시티",
            "자이": "봉담자이 프라이드시티",
        }
        return alias.get(t, t)

    def _join_preferred_tabs(self, tabs: list[str]) -> str:
        ordered = [t for t in PROPERTY_TABS if t in set(tabs)]
        return ", ".join(ordered)

    def _open_tab_multi_select(self, parent: tk.Misc, current_text: str) -> str | None:
        popup = tk.Toplevel(parent)
        popup.title("희망 유형 선택")
        self._fit_toplevel(popup, 420, 320)

        current = set(self._parse_preferred_tabs(current_text))
        vars_map = {tab: tk.BooleanVar(value=(tab in current)) for tab in PROPERTY_TABS}

        frm = ttk.Frame(popup)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="복수 선택 가능").pack(anchor="w", pady=(0, 8))
        for tab in PROPERTY_TABS:
            ttk.Checkbutton(frm, text=tab, variable=vars_map[tab]).pack(anchor="w", pady=2)

        result: dict[str, str | None] = {"value": None}

        def done():
            selected = [tab for tab in PROPERTY_TABS if vars_map[tab].get()]
            result["value"] = self._join_preferred_tabs(selected)
            popup.destroy()

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="확인", command=done).pack(side="left", padx=4)
        ttk.Button(btns, text="취소", command=popup.destroy).pack(side="left", padx=4)

        popup.transient(parent)
        popup.grab_set()
        popup.wait_window()
        return result["value"]

    def _build_customer_ui(self):
        top = ttk.LabelFrame(self.customer_tab, text="고객")
        top.pack(fill="x", padx=10, pady=8)

        ttk.Button(top, text="+ 고객 등록", command=self.open_customer_wizard).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="숨김함", command=self.open_hidden_customers_window).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="내보내기/동기화", command=self.export_sync).pack(side="left", padx=4, pady=6)

        ttk.Label(top, text="거래유형").pack(side="left", padx=(18, 4))
        self.customer_deal_filter_var = tk.StringVar(value="전체")
        deal_cb = ttk.Combobox(top, textvariable=self.customer_deal_filter_var, values=["전체", "매매", "전세", "월세"], state="readonly", width=8)
        deal_cb.pack(side="left", padx=4)
        deal_cb.bind("<<ComboboxSelected>>", lambda _e: self.refresh_customers())

        ttk.Label(top, text="전화번호 검색").pack(side="left", padx=(18, 4))
        self.customer_phone_query = tk.StringVar(value="")
        ent = ttk.Entry(top, textvariable=self.customer_phone_query, width=24)
        ent.pack(side="left", padx=4)
        ent.bind("<KeyRelease>", lambda _e: self.refresh_customers())
        ent.bind("<Return>", lambda _e: self.refresh_customers())
        ttk.Button(top, text="초기화", command=lambda: (self.customer_phone_query.set(""), self.customer_deal_filter_var.set("전체"), self.refresh_customers())).pack(side="left", padx=4)

        cols = ("id", "customer_name", "phone", "preferred_tab", "deal_type", "budget", "size", "move_in", "floor_preference", "status", "updated_at")
        col_defs = [
            ("id", 55),
            ("customer_name", 110),
            ("phone", 120),
            ("preferred_tab", 170),
            ("deal_type", 80),
            ("budget", 170),
            ("size", 100),
            ("move_in", 120),
            ("floor_preference", 90),
            ("status", 80),
            ("updated_at", 150),
        ]
        col_labels = {
            "id": "",
            "customer_name": "고객명",
            "phone": "전화번호",
            "preferred_tab": "희망유형",
            "deal_type": "거래유형",
            "budget": "예산",
            "size": "희망크기",
            "move_in": "입주희망",
            "floor_preference": "층수선호",
            "status": "상태",
            "updated_at": "업데이트",
        }

        self.customer_nb = ttk.Notebook(self.customer_tab)
        self.customer_nb.pack(fill="both", expand=True, padx=10, pady=8)
        self.customer_trees: dict[str, ttk.Treeview] = {}

        for tab_name in ["전체", *PROPERTY_TABS]:
            frame = ttk.Frame(self.customer_nb)
            self.customer_nb.add(frame, text=tab_name)
            tree = ttk.Treeview(frame, columns=cols, show="headings", height=22)
            for c, w in col_defs:
                tree.heading(c, text=col_labels.get(c, c))
                tree.column(c, width=w, stretch=(c != "id"))
                if c == "id":
                    tree.column(c, minwidth=0, width=0, stretch=False)
            tree.pack(fill="both", expand=True)
            c_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=c_x.set)
            c_x.pack(fill="x", pady=(0, 6))
            tree.bind("<Double-1>", self._on_double_click_customer)
            self.customer_trees[tab_name] = tree

        btns = ttk.Frame(self.customer_tab)
        btns.pack(fill="x", padx=10, pady=4)
        ttk.Button(btns, text="상세", command=self.open_selected_customer_detail).pack(side="left", padx=4)
        ttk.Button(btns, text="숨김/보임", command=self.toggle_selected_customer).pack(side="left", padx=4)
        ttk.Button(btns, text="삭제", command=self.delete_selected_customer).pack(side="left", padx=4)

    def _current_customer_tree(self) -> ttk.Treeview:
        current_tab = self.customer_nb.tab(self.customer_nb.select(), "text") if hasattr(self, "customer_nb") else "전체"
        return self.customer_trees.get(current_tab) or self.customer_trees.get("전체")

    def open_customer_wizard(self):
        win = tk.Toplevel(self.root)
        win.title("고객 등록")
        self._fit_toplevel(win, 820, 660)

        vars_ = {
            "customer_name": tk.StringVar(),
            "phone": tk.StringVar(),
            "preferred_tab": tk.StringVar(value=""),
            "deal_type": tk.StringVar(value="매매"),
            "size_unit": tk.StringVar(value="㎡"),
            "size_value": tk.StringVar(),
            "budget_10m": tk.StringVar(value="0"),
            "wolse_deposit_10m": tk.StringVar(value="0"),
            "wolse_rent_10man": tk.StringVar(value="0"),
            "move_in_period": tk.StringVar(),
            "location_preference": tk.StringVar(),
            "view_preference": tk.StringVar(value="비중요"),
            "condition_preference": tk.StringVar(value="비중요"),
            "floor_preference": tk.StringVar(value="상관없음"),
            "has_pet": tk.StringVar(value="없음"),
            "extra_needs": tk.StringVar(),
            "status": tk.StringVar(value="문의"),
        }

        header = ttk.Frame(win)
        header.pack(fill="x", padx=10, pady=(10, 4))
        progress_var = tk.StringVar(value="1/3 단계: 기본")
        ttk.Label(header, textvariable=progress_var).pack(anchor="w")

        body = ttk.Frame(win)
        body.pack(fill="both", expand=True, padx=10, pady=6)
        s1 = ttk.Frame(body)
        s2 = ttk.Frame(body)
        s3 = ttk.Frame(body)
        steps = [s1, s2, s3]
        step_titles = ["기본", "희망조건", "일정/메모"]

        # Step1
        ttk.Label(s1, text="고객명").grid(row=0, column=0, padx=6, pady=8, sticky="e")
        ttk.Entry(s1, textvariable=vars_["customer_name"], width=32).grid(row=0, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s1, text="전화번호").grid(row=1, column=0, padx=6, pady=8, sticky="e")
        ttk.Entry(s1, textvariable=vars_["phone"], width=32).grid(row=1, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s1, text="상태").grid(row=2, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s1, textvariable=vars_["status"], values=CUSTOMER_STATUS_VALUES, state="readonly", width=29).grid(row=2, column=1, padx=6, pady=8, sticky="w")

        # Step2
        ttk.Label(s2, text="희망 유형").grid(row=0, column=0, padx=6, pady=8, sticky="e")
        pref_tab_wrap = ttk.Frame(s2)
        pref_tab_wrap.grid(row=0, column=1, padx=6, pady=8, sticky="w")
        ttk.Entry(pref_tab_wrap, textvariable=vars_["preferred_tab"], width=24, state="readonly").pack(side="left")
        ttk.Button(pref_tab_wrap, text="선택", command=lambda: (lambda v: vars_["preferred_tab"].set(v) if v is not None else None)(self._open_tab_multi_select(win, vars_["preferred_tab"].get()))).pack(side="left", padx=4)

        ttk.Label(s2, text="거래유형").grid(row=1, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s2, textvariable=vars_["deal_type"], values=["매매", "전세", "월세"], state="readonly", width=29).grid(row=1, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s2, text="희망 크기").grid(row=2, column=0, padx=6, pady=8, sticky="e")
        size_wrap = ttk.Frame(s2)
        size_wrap.grid(row=2, column=1, padx=6, pady=8, sticky="w")
        ttk.Entry(size_wrap, textvariable=vars_["size_value"], width=18).pack(side="left")
        ttk.Combobox(size_wrap, textvariable=vars_["size_unit"], values=["㎡", "평"], state="readonly", width=8).pack(side="left", padx=4)
        ttk.Label(s2, text="매매/전세 예산(천만원)").grid(row=3, column=0, padx=6, pady=8, sticky="e")
        ttk.Entry(s2, textvariable=vars_["budget_10m"], width=16).grid(row=3, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s2, text="월세 보증금(천만원)").grid(row=4, column=0, padx=6, pady=8, sticky="e")
        ttk.Entry(s2, textvariable=vars_["wolse_deposit_10m"], width=16).grid(row=4, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s2, text="월세액(십만원)").grid(row=5, column=0, padx=6, pady=8, sticky="e")
        ttk.Entry(s2, textvariable=vars_["wolse_rent_10man"], width=16).grid(row=5, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s2, text="선호 위치").grid(row=6, column=0, padx=6, pady=8, sticky="e")
        ttk.Entry(s2, textvariable=vars_["location_preference"], width=32).grid(row=6, column=1, padx=6, pady=8, sticky="w")

        # Step3
        ttk.Label(s3, text="입주 희망일").grid(row=0, column=0, padx=6, pady=8, sticky="e")
        move_mode = tk.StringVar(value="협의")
        move_y = tk.IntVar(value=datetime.now().year)
        move_m = tk.IntVar(value=datetime.now().month)
        move_d = tk.IntVar(value=datetime.now().day)
        move_wrap = ttk.Frame(s3)
        move_wrap.grid(row=0, column=1, padx=6, pady=8, sticky="w")
        move_mode_cb = ttk.Combobox(move_wrap, textvariable=move_mode, values=["즉시", "협의", "날짜선택"], state="readonly", width=10)
        move_mode_cb.pack(side="left")
        date_wrap = ttk.Frame(move_wrap)
        ttk.Spinbox(date_wrap, from_=2020, to=2100, textvariable=move_y, width=6).pack(side="left")
        ttk.Spinbox(date_wrap, from_=1, to=12, textvariable=move_m, width=4).pack(side="left", padx=2)
        ttk.Spinbox(date_wrap, from_=1, to=31, textvariable=move_d, width=4).pack(side="left", padx=2)

        def refresh_move_ui(*_a):
            if move_mode.get() == "날짜선택":
                if not date_wrap.winfo_ismapped():
                    date_wrap.pack(side="left", padx=6)
            else:
                if date_wrap.winfo_ismapped():
                    date_wrap.pack_forget()

        move_mode_cb.bind("<<ComboboxSelected>>", refresh_move_ui)
        refresh_move_ui()

        ttk.Label(s3, text="층수 선호").grid(row=1, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s3, textvariable=vars_["floor_preference"], values=["저", "중", "고", "상관없음"], state="readonly", width=29).grid(row=1, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s3, text="뷰 중요도").grid(row=2, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s3, textvariable=vars_["view_preference"], values=["중요", "비중요"], state="readonly", width=29).grid(row=2, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s3, text="컨디션 중요도").grid(row=3, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s3, textvariable=vars_["condition_preference"], values=["중요", "비중요"], state="readonly", width=29).grid(row=3, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s3, text="애완동물").grid(row=4, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s3, textvariable=vars_["has_pet"], values=["있음", "없음"], state="readonly", width=29).grid(row=4, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s3, text="추가 요청사항").grid(row=5, column=0, padx=6, pady=8, sticky="ne")
        ttk.Entry(s3, textvariable=vars_["extra_needs"], width=48).grid(row=5, column=1, padx=6, pady=8, sticky="w")

        def _is_non_negative_int(v: str) -> bool:
            v = v.strip()
            return v.isdigit() if v else False

        def validate_step(idx: int) -> bool:
            if idx == 0:
                if not (vars_["customer_name"].get().strip() or vars_["phone"].get().strip()):
                    messagebox.showwarning("입력 확인", "고객명 또는 전화번호 중 하나는 입력해주세요.")
                    return False
                return True
            if idx == 1:
                if not vars_["deal_type"].get().strip():
                    messagebox.showwarning("입력 확인", "거래유형을 선택해주세요.")
                    return False
                nums = [vars_["budget_10m"].get(), vars_["wolse_deposit_10m"].get(), vars_["wolse_rent_10man"].get()]
                if any(not _is_non_negative_int(n) for n in nums):
                    messagebox.showwarning("입력 확인", "예산(천만원/십만원)은 숫자만 입력해주세요.")
                    return False
                return True
            if idx == 2:
                if move_mode.get().strip() not in {"즉시", "협의", "날짜선택"}:
                    messagebox.showwarning("입력 확인", "입주희망일 방식을 선택해주세요.")
                    return False
                return True
            return True

        current = {"idx": 0}

        def show_step(idx: int):
            for frm in steps:
                frm.pack_forget()
            steps[idx].pack(fill="both", expand=True)
            progress_var.set(f"{idx + 1}/3 단계: {step_titles[idx]}")
            prev_btn.configure(state=("normal" if idx > 0 else "disabled"))
            next_btn.configure(state=("normal" if idx < len(steps) - 1 else "disabled"))
            save_btn.configure(state=("normal" if idx == len(steps) - 1 else "disabled"))

        def go_prev():
            if current["idx"] > 0:
                current["idx"] -= 1
                show_step(current["idx"])

        def go_next():
            idx = current["idx"]
            if not validate_step(idx):
                return
            if idx < len(steps) - 1:
                current["idx"] += 1
                show_step(current["idx"])

        def save():
            if not validate_step(current["idx"]):
                return
            payload = {k: v.get().strip() for k, v in vars_.items()}
            payload["move_in_period"] = f"{move_y.get():04d}-{move_m.get():02d}-{move_d.get():02d}" if move_mode.get() == "날짜선택" else move_mode.get().strip()
            deal = payload.get("deal_type", "")
            if deal == "월세":
                payload["budget"] = f"월세 {payload.get('wolse_deposit_10m','0')}천만원 / {payload.get('wolse_rent_10man','0')}십만원"
            elif deal == "전세":
                payload["budget"] = f"전세 {payload.get('budget_10m','0')}천만원"
            else:
                payload["budget"] = f"매매 {payload.get('budget_10m','0')}천만원"
            if payload["size_unit"] == "㎡":
                payload["preferred_area"] = payload["size_value"]
                payload["preferred_pyeong"] = ""
            else:
                payload["preferred_pyeong"] = payload["size_value"]
                payload["preferred_area"] = ""

            payload["budget_10m"] = int(payload.get("budget_10m", "0") or 0)
            payload["wolse_deposit_10m"] = int(payload.get("wolse_deposit_10m", "0") or 0)
            payload["wolse_rent_10man"] = int(payload.get("wolse_rent_10man", "0") or 0)

            try:
                add_customer(payload)
            except Exception as exc:
                messagebox.showerror("저장 실패", str(exc))
                return

            self.refresh_all()
            messagebox.showinfo("저장 완료", "고객이 저장되었습니다.")
            win.destroy()

        footer = ttk.Frame(win)
        footer.pack(fill="x", padx=10, pady=10)
        prev_btn = ttk.Button(footer, text="이전", command=go_prev)
        next_btn = ttk.Button(footer, text="다음", command=go_next)
        save_btn = ttk.Button(footer, text="저장", command=save)
        cancel_btn = ttk.Button(footer, text="취소", command=win.destroy)
        prev_btn.pack(side="left", padx=4)
        next_btn.pack(side="left", padx=4)
        save_btn.pack(side="left", padx=4)
        cancel_btn.pack(side="left", padx=4)

        show_step(0)

    def create_customer(self):
        self.open_customer_wizard()

    def refresh_customers(self):
        q = self.customer_phone_query.get().strip() if hasattr(self, "customer_phone_query") else ""
        deal_filter = self.customer_deal_filter_var.get().strip() if hasattr(self, "customer_deal_filter_var") else "전체"

        rows = list_customers(phone_query=q)
        order = {"문의": 0, "임장예약": 1, "계약진행": 2, "계약완료": 3, "입주": 4, "대기": 5}
        rows.sort(key=lambda r: (order.get(str(r.get("status") or ""), 9), -int(r.get("id") or 0)))

        for tree in getattr(self, "customer_trees", {}).values():
            for i in tree.get_children():
                tree.delete(i)

        for row in rows:
            if row.get("hidden"):
                continue
            if deal_filter != "전체" and str(row.get("deal_type") or "") != deal_filter:
                continue

            size = ""
            if row.get("size_value"):
                size = f"{row.get('size_value')} {row.get('size_unit') or ''}".strip()
            elif row.get("preferred_area"):
                size = f"{row.get('preferred_area')} ㎡"
            elif row.get("preferred_pyeong"):
                size = f"{row.get('preferred_pyeong')} 평"

            values = (
                row.get("id"),
                row.get("customer_name"),
                row.get("phone"),
                row.get("preferred_tab") or "",
                row.get("deal_type") or "",
                row.get("budget") or "",
                size,
                row.get("move_in_period"),
                row.get("floor_preference"),
                row.get("status"),
                row.get("updated_at"),
            )

            # 전체 탭
            if "전체" in self.customer_trees:
                self.customer_trees["전체"].insert("", "end", values=values)

            pref_tabs = {self._normalize_property_tab(t) for t in self._parse_preferred_tabs(str(row.get("preferred_tab") or ""))}
            for tab_name in PROPERTY_TABS:
                norm_tab = self._normalize_property_tab(tab_name)
                if norm_tab in pref_tabs and tab_name in self.customer_trees:
                    tab_values = list(values)
                    tab_values[3] = tab_name  # 탭별 리스트에서는 현재 탭명만 표시
                    self.customer_trees[tab_name].insert("", "end", values=tuple(tab_values))


    def toggle_selected_customer(self):
        cid = self._selected_id_from_tree(self._current_customer_tree(), value_index=0)
        if cid is None:
            return
        toggle_customer_hidden(cid)
        self.refresh_all()
    
    def delete_selected_customer(self):
        cid = self._selected_id_from_tree(self._current_customer_tree(), value_index=0)
        if cid is None:
            return
        if messagebox.askyesno("확인", "해당 고객 요청을 삭제 처리(복구 가능)할까요?"):
            soft_delete_customer(cid)
            self.refresh_all()
    
    def open_selected_customer_detail(self):
        cid = self._selected_id_from_tree(self._current_customer_tree(), value_index=0)
        if cid is None:
            return
        self.open_customer_detail(cid)
    
    def _on_double_click_customer(self, event):
        tree = event.widget if isinstance(event.widget, ttk.Treeview) else self._current_customer_tree()
        item = tree.identify_row(event.y)
        if item:
            try:
                cid = int(tree.item(item, "values")[0])
                self.open_customer_detail(cid)
                return
            except Exception:
                pass
        self.open_selected_customer_detail()
    
        # -----------------
        # Matching
        # -----------------
    def open_hidden_customers_window(self):
        win = tk.Toplevel(self.root)
        win.title("고객 숨김함")
        self._fit_toplevel(win, 760, 560)
        tree = ttk.Treeview(win, columns=("id", "name", "phone"), show="headings", height=14)
        for c in ("id", "name", "phone"):
            tree.heading(c, text=c)
        tree.pack(fill="both", expand=True, padx=8, pady=8)
        hidden_rows = [r for r in list_customers(include_deleted=False) if r.get("hidden")]
        for r in hidden_rows:
            tree.insert("", "end", values=(r.get("id"), r.get("customer_name"), r.get("phone")))

        def _selected_id():
            sel = tree.selection()
            if not sel:
                return None
            try:
                return int(tree.item(sel[0], "values")[0])
            except Exception:
                return None

        def restore():
            cid = _selected_id()
            if cid is None:
                return
            toggle_customer_hidden(cid)
            self.refresh_customers()
            win.destroy()

        def open_detail():
            cid = _selected_id()
            if cid is None:
                return
            self.open_customer_detail(cid)

        def delete_it():
            cid = _selected_id()
            if cid is None:
                return
            soft_delete_customer(cid)
            self.refresh_all()
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(pady=6)
        ttk.Button(btns, text="복구", command=restore).pack(side="left", padx=4)
        ttk.Button(btns, text="상세", command=open_detail).pack(side="left", padx=4)
        ttk.Button(btns, text="삭제", command=delete_it).pack(side="left", padx=4)

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

        cols = ("score", "summary", "deal", "reasons", "property_id")
        self.match_tree = ttk.Treeview(bottom, columns=cols, show="headings", height=22, selectmode="extended")
        col_defs = [("score", 70), ("summary", 300), ("deal", 180), ("reasons", 380), ("property_id", 0)]
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
                    f"{p.get('tab','')} / {p.get('address_detail','')} / {p.get('unit_type','')}",
                    money_utils.property_price_summary(p) or p.get("status", ""),
                    " / ".join(r.reasons[:2]),
                    p.get("id"),
                ),
            )
    
    def _on_double_click_match(self, _event):
        selected = self.match_tree.selection()
        if not selected:
            return
        try:
            pid = int(self.match_tree.item(selected[0], "values")[4])
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
                pid = int(self.match_tree.item(item, "values")[4])
                ids.append(pid)
            except Exception:
                continue
        if ids:
            return ids
    
        # fallback: first N rows in the tree
        for item in self.match_tree.get_children()[:fallback_top_n]:
            try:
                pid = int(self.match_tree.item(item, "values")[4])
                ids.append(pid)
            except Exception:
                continue
        return ids

    def _ranked_photos_for_property(self, property_id: int) -> list[dict[str, str]]:
        priority = ["거실", "안방", "작은방", "화장실", "주방", "현관", "뷰", "기타"]
        rank = {name: i for i, name in enumerate(priority)}
        rows = [
            {"file_path": str(r.get("file_path") or ""), "tag": str(r.get("tag") or "")}
            for r in list_photos(property_id)
            if r.get("file_path")
        ]
        rows.sort(key=lambda x: (rank.get(x.get("tag", ""), 99), x.get("tag", "")))
        return rows

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
        photos_by_property: dict[int, list[dict[str, str]]] = {}
    
        for pid in pids:
            row = get_property(pid)
            if not row:
                continue
            props.append(row)
            photos_by_property[pid] = self._ranked_photos_for_property(pid)
    
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
    def export_sync(self, *, silent: bool = False):
        self.settings = self._load_settings()  # 최신 반영
        props = list_properties(include_deleted=False)
        custs = list_customers(include_deleted=False)
        photos = list_photos_all()
        viewings = list_viewings()
        tasks = list_tasks(include_done=False)

        ok, msg = upload_visible_data(
            props,
            custs,
            photos=photos,
            viewings=viewings,
            tasks=tasks,
            settings=SyncSettings(webhook_url=self.settings.webhook_url, sync_dir=self.settings.sync_dir),
        )
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        if hasattr(self, "sync_status_var"):
            self.sync_status_var.set(f"마지막 동기화: {stamp} / {'성공' if ok else '실패'}")
        if silent:
            return
        if ok:
            messagebox.showinfo("동기화", msg)
        else:
            messagebox.showerror("동기화 실패", msg)
    
    def start_auto_sync_loop(self):
        def _run():
            try:
                self.export_sync(silent=True)
            finally:
                self.root.after(10 * 60 * 1000, _run)

        self.root.after(10 * 60 * 1000, _run)

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
        self._fit_toplevel(win, 1100, 760)

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
                                             values=PROPERTY_STATUS_VALUES))

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
            self.refresh_all()

            linked_tasks = list_tasks(include_done=False, entity_type="PROPERTY", entity_id=property_id)
            linked_customers = sorted({int(v.get("customer_id")) for v in list_viewings(property_id=property_id) if v.get("customer_id")})
            msg = f"저장 완료. 연결된 할일 {len(linked_tasks)}개 / 연결된 고객 {len(linked_customers)}명"
            if linked_tasks or linked_customers:
                if messagebox.askyesno("완료", msg + "\n바로 열까요?"):
                    self._open_related_navigation_popup(title="연결된 고객", customers=linked_customers)
            else:
                messagebox.showinfo("완료", "저장되었습니다.")

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
        ph_tree.bind("<Double-1>", lambda _e: open_photo())

        ph_controls = ttk.Frame(ph_box)
        ph_controls.pack(fill="x", pady=6)
        ph_tag_var = tk.StringVar(value=PHOTO_TAG_VALUES[0])
        ttk.Label(ph_controls, text="사진 구분").pack(side="left", padx=4)
        ttk.Combobox(ph_controls, textvariable=ph_tag_var, values=PHOTO_TAG_VALUES, state="readonly", width=14).pack(side="left", padx=4)

        def refresh_photos():
            for i in ph_tree.get_children():
                ph_tree.delete(i)
            for r in list_photos(property_id):
                ph_tree.insert("", "end", values=(r.get("id"), r.get("tag"), r.get("file_path"), r.get("created_at")))

        def add_photo_ui():
            src = filedialog.askopenfilename(title="사진 선택")
            if not src:
                return
            tag = ph_tag_var.get().strip()
            if not tag:
                messagebox.showwarning("확인", "사진 구분 태그를 선택해주세요.")
                return
            try:
                dst = self._copy_photo_to_library(Path(src), property_id, tag=tag)
                add_photo(property_id, str(dst), tag=tag)
                current = get_property(property_id)
                if current and str(current.get("status") or "") == "신규등록":
                    update_property(property_id, {"status": "검수완료(사진등록)"})
                ph_tag_var.set(PHOTO_TAG_VALUES[0])
                refresh_photos()
                self.refresh_properties()
                # 업로드 후에도 상세창/물건탭이 유지되도록 포커스 복귀
                try:
                    self.main.select(self.property_tab)
                    nb.select(tab_photos)
                    win.lift()
                    win.focus_force()
                except Exception:
                    pass
            except Exception as exc:
                messagebox.showerror("사진 업로드 실패", str(exc))

        def open_photo():
            sel = ph_tree.selection()
            if not sel:
                return
            path = ph_tree.item(sel[0], "values")[2]
            if path:
                _open_file(Path(path))

        def remove_photo():
            sel = ph_tree.selection()
            if not sel:
                return
            pid_ = int(ph_tree.item(sel[0], "values")[0])
            if messagebox.askyesno("확인", "사진 기록을 삭제할까요? (파일은 남습니다)"):
                delete_photo(pid_)
                refresh_photos()

        ttk.Button(ph_controls, text="추가", command=add_photo_ui).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="사진 보기", command=open_photo).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="폴더 열기", command=lambda: _open_folder(Path(ph_tree.item(ph_tree.selection()[0], "values")[2]).parent) if ph_tree.selection() else None).pack(side="left", padx=4)
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
    
    def _copy_photo_to_library(self, src: Path, property_id: int, *, tag: str = "") -> Path:
        self.settings = self._load_settings()
        base = self.settings.sync_dir
        dst_dir = base / "Photos" / f"PR_{property_id:06d}"
        dst_dir.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_tag = "".join(ch for ch in (tag or "기타") if ch.isalnum() or ch in "-_가-힣").strip() or "기타"
        ext = src.suffix.lower() or ".jpg"
        dst = dst_dir / f"PR_{property_id:06d}_{safe_tag}_{ts}{ext}"
        shutil.copy2(src, dst)
        return dst
    
    def _open_add_viewing_window(self, property_id: int, on_saved):
        win = tk.Toplevel(self.root)
        win.title("일정 추가")
        self._fit_toplevel(win, 620, 380)

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
        self._fit_toplevel(win, 980, 560)

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
                wrap = ttk.Frame(form)
                ttk.Entry(wrap, textvariable=vars_[key], width=20, state="readonly").pack(side="left")
                ttk.Button(wrap, text="선택", command=lambda k=key: (lambda v: vars_[k].set(v) if v is not None else None)(self._open_tab_multi_select(win, vars_[k].get()))).pack(side="left", padx=2)
                w = wrap
            elif key == "status":
                w = ttk.Combobox(form, textvariable=vars_[key], values=CUSTOMER_STATUS_VALUES, width=22, state="readonly")
            else:
                w = ttk.Entry(form, textvariable=vars_[key], width=26)
            w.grid(row=i // 3, column=(i % 3) * 2 + 1, padx=6, pady=6, sticky="w")

        actions = ttk.Frame(win)
        actions.pack(fill="x", padx=10, pady=10)

        def save_changes():
            data = {k: v.get() for k, v in vars_.items()}
            update_customer(customer_id, data)
            self.refresh_all()

            linked_tasks = list_tasks(include_done=False, entity_type="CUSTOMER", entity_id=customer_id)
            linked_props = sorted({int(v.get("property_id")) for v in list_viewings(customer_id=customer_id) if v.get("property_id")})
            msg = f"저장 완료. 연결된 할일 {len(linked_tasks)}개 / 연결된 물건 {len(linked_props)}개"
            if linked_tasks or linked_props:
                if messagebox.askyesno("완료", msg + "\n바로 열까요?"):
                    self._open_related_navigation_popup(title="연결된 물건", properties=linked_props)
            else:
                messagebox.showinfo("완료", "저장되었습니다.")

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
        ttk.Button(actions, text="제안서 PDF(상위5)", command=lambda: self.generate_proposal_for_customer(customer_id, top_n=5)).pack(side="left", padx=4)
        ttk.Button(actions, text="제안문구 복사(상위5)", command=lambda: self.copy_message_for_customer(customer_id, top_n=5)).pack(side="left", padx=4)
    
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


    def _open_related_navigation_popup(self, *, title: str, customers: list[int] | None = None, properties: list[int] | None = None):
        customers = customers or []
        properties = properties or []
        if not customers and not properties:
            return

        win = tk.Toplevel(self.root)
        win.title(title)
        self._fit_toplevel(win, 620, 460)

        tree = ttk.Treeview(win, columns=("type", "id", "name"), show="headings")
        tree.heading("type", text="유형")
        tree.heading("id", text="ID")
        tree.heading("name", text="이름/주소")
        tree.column("type", width=100, anchor="center")
        tree.column("id", width=80, anchor="center")
        tree.column("name", width=300)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for cid in customers:
            c = get_customer(cid)
            tree.insert("", "end", values=("고객", cid, str((c or {}).get("customer_name") or f"고객 {cid}")))
        for pid in properties:
            p = get_property(pid)
            name = f"{(p or {}).get('complex_name','')} {(p or {}).get('address_detail','')}".strip()
            tree.insert("", "end", values=("물건", pid, name or f"물건 {pid}"))

        def open_selected():
            sel = tree.selection()
            if not sel:
                return
            t, sid, _ = tree.item(sel[0], "values")
            try:
                sid_i = int(sid)
            except Exception:
                return
            if t == "고객":
                self.open_customer_detail(sid_i)
            else:
                self.open_property_detail(sid_i)

        tree.bind("<Double-1>", lambda _e: open_selected())
        ttk.Button(win, text="열기", command=open_selected).pack(pady=6)

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
    
        photos_by_property: dict[int, list[dict[str, str]]] = {}
        for p in props:
            try:
                pid = int(p.get("id"))
            except Exception:
                continue
            photos_by_property[pid] = self._ranked_photos_for_property(pid)
    
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
