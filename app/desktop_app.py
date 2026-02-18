from __future__ import annotations

import os
import json
import platform
import re
import shutil
import threading
import webbrowser
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, simpledialog, ttk

from .proposal import generate_proposal_pdf
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
    list_presets,
    set_setting,
    soft_delete_customer,
    soft_delete_property,
    restore_customer,
    restore_property,
    toggle_customer_hidden,
    toggle_property_hidden,
    update_customer,
    update_property,
    upsert_preset,
    delete_preset,
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

def fmt_phone(phone: str) -> str:
    d = "".join(ch for ch in str(phone or "") if ch.isdigit())
    if len(d) == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    if len(d) == 10:
        return f"{d[:3]}-{d[3:6]}-{d[6:]}"
    return str(phone or "")


def customer_label(c: dict) -> str:
    name = str(c.get("customer_name") or "").strip()
    phone = fmt_phone(str(c.get("phone") or "").strip())
    if name and phone:
        return f"{name}({phone})"
    return name or phone or "고객"


def property_label(p: dict) -> str:
    complex_name = str(p.get("complex_name") or "").strip()
    dong = str(p.get("dong") or "").strip()
    ho = str(p.get("ho") or "").strip()
    addr = str(p.get("address_detail") or "").strip()
    unit = str(p.get("unit_type") or "").strip()
    pyeong = str(p.get("pyeong") or "").strip()
    core = " ".join(x for x in [complex_name, f"{dong}동" if dong else "", f"{ho}호" if ho else ""] if x).strip()
    if not core:
        core = " ".join(x for x in [complex_name, addr] if x).strip() or "매물"
    tail = ""
    if unit and pyeong:
        tail = f" {unit}({pyeong}평)"
    elif unit:
        tail = f" {unit}"
    return (core + tail).strip()



@dataclass
class AppSettings:
    sync_dir: Path
    webhook_url: str
    auto_sync_sec: int = 60
    auto_sync_on: bool = True
    export_mode: str = "single"
    export_ics: bool = False


class LedgerDesktopApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("엄마 부동산 장부 (오프라인 우선)")
        self.root.geometry("1920x1080")
        try:
            self.root.state("zoomed")
        except Exception:
            pass
        self._apply_ui_scale(1.25)

        self.sort_state: dict[tuple[int, str], bool] = {}
        self.task_sort_col = "due_at"
        self.task_sort_desc = False
        self.enable_handoff = False
        self._sync_after_id = None
        self._auto_sync_after_id = None
        self._sync_inflight = False
        self.property_search_terms_by_tab: dict[str, str] = {tab: "" for tab in PROPERTY_TABS}
        self.property_deal_filter_vars: dict[str, tk.BooleanVar] = {}
        self.current_customer_var = tk.StringVar(value="현재 고객: 미선택")
        self.current_property_var = tk.StringVar(value="현재 매물: 미선택")
        self._last_undo_payload: dict | None = None
        self._preset_buttons: dict[int, ttk.Button] = {}

        init_db()
        self.settings = self._load_settings()

        status_bar = ttk.Frame(root)
        status_bar.pack(fill="x", padx=10, pady=(8, 2))
        ttk.Label(status_bar, textvariable=self.current_customer_var).pack(side="left", padx=(0, 16))
        ttk.Label(status_bar, textvariable=self.current_property_var).pack(side="left")

        self.main = ttk.Notebook(root)
        self.main.pack(fill="both", expand=True)

        self.dashboard_tab = ttk.Frame(self.main)
        self.property_tab = ttk.Frame(self.main)
        self.customer_tab = ttk.Frame(self.main)
        self.settings_tab = ttk.Frame(self.main)

        self.main.add(self.dashboard_tab, text="오늘")
        self.main.add(self.property_tab, text="물건")
        self.main.add(self.customer_tab, text="고객")
        self.main.add(self.settings_tab, text="설정")

        self._build_dashboard_ui()
        self._build_property_ui()
        self._build_customer_ui()
        self._build_settings_ui()

        self._bind_global_shortcuts()
        self.refresh_all()
        self._refresh_preset_buttons()
        self.start_auto_sync_loop()

    def _apply_ui_scale(self, scale: float = 1.25) -> None:
        try:
            self.root.tk.call("tk", "scaling", scale)
        except Exception:
            pass

        family = "Malgun Gothic" if platform.system().lower().startswith("win") else "TkDefaultFont"
        base_size = 11
        heading_size = base_size + 2

        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"):
            try:
                f = tkfont.nametofont(name)
                f.configure(family=family, size=base_size)
            except Exception:
                continue
        try:
            tkfont.nametofont("TkHeadingFont").configure(family=family, size=heading_size, weight="bold")
        except Exception:
            pass

        self.palette = {
            "bg": "#F7F9FC",
            "panel": "#FFFFFF",
            "text": "#1F2937",
            "muted": "#6B7280",
            "border": "#D7DFEA",
            "accent": "#2F6FEB",
            "danger": "#C24141",
        }

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(".", background=self.palette["bg"], foreground=self.palette["text"], font=(family, base_size))
        style.configure("TFrame", background=self.palette["bg"])
        style.configure("TLabelframe", background=self.palette["panel"], bordercolor=self.palette["border"], relief="solid")
        style.configure("TLabelframe.Label", background=self.palette["panel"], foreground=self.palette["text"], font=(family, heading_size, "bold"))
        style.configure("TLabel", background=self.palette["bg"], foreground=self.palette["text"], font=(family, base_size))
        style.configure("Muted.TLabel", background=self.palette["bg"], foreground=self.palette["muted"], font=(family, base_size))

        style.configure("TButton", font=(family, base_size), padding=(10, 6), borderwidth=1)
        style.configure("Secondary.TButton", font=(family, base_size), padding=(10, 6), background=self.palette["panel"], foreground=self.palette["text"], bordercolor=self.palette["border"])
        style.map("Secondary.TButton", background=[("active", "#F0F4FA")])
        style.configure("Primary.TButton", font=(family, base_size, "bold"), padding=(10, 6), background=self.palette["accent"], foreground="#FFFFFF", bordercolor=self.palette["accent"])
        style.map("Primary.TButton", background=[("active", "#285FD0")])
        style.configure("Danger.TButton", font=(family, base_size, "bold"), padding=(10, 6), background="#FBECEC", foreground=self.palette["danger"], bordercolor="#F0BDBD")
        style.map("Danger.TButton", background=[("active", "#F6DDDD")])

        style.configure("TEntry", fieldbackground=self.palette["panel"], bordercolor=self.palette["border"], padding=(6, 5), foreground=self.palette["text"])
        style.configure("TCombobox", fieldbackground=self.palette["panel"], bordercolor=self.palette["border"], padding=(6, 5), foreground=self.palette["text"])
        style.map("TCombobox", bordercolor=[("focus", self.palette["accent"])])
        style.configure("TCheckbutton", background=self.palette["bg"], foreground=self.palette["text"], font=(family, base_size))
        style.configure("TRadiobutton", background=self.palette["bg"], foreground=self.palette["text"], font=(family, base_size))

        style.configure("TNotebook", background=self.palette["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", font=(family, base_size), padding=(12, 7), background="#ECF1F8", foreground=self.palette["muted"])
        style.map("TNotebook.Tab", background=[("selected", self.palette["panel"])], foreground=[("selected", self.palette["text"])])

        style.configure("Treeview", font=(family, 10), rowheight=30, background=self.palette["panel"], fieldbackground=self.palette["panel"], foreground=self.palette["text"], bordercolor=self.palette["border"])
        style.configure("Treeview.Heading", font=(family, base_size, "bold"), background="#EEF3FA", foreground=self.palette["text"], relief="flat")
        style.map("Treeview", background=[("selected", self.palette["accent"])], foreground=[("selected", "#FFFFFF")])

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

    def _resolve_property_contact(self, row: dict) -> str:
        for key in ("owner_phone", "tenant_phone", "agent_phone", "phone"):
            value = str(row.get(key) or "").strip()
            if value:
                return value
        return "미입력"

    def _set_current_customer(self, row: dict | None) -> None:
        if not row:
            self.current_customer_var.set("현재 고객: 미선택")
            return
        self.current_customer_var.set(f"현재 고객: {customer_label(row)}")

    def _set_current_property(self, row: dict | None) -> None:
        if not row:
            self.current_property_var.set("현재 매물: 미선택")
            return
        self.current_property_var.set(f"현재 매물: {property_label(row)}")

    def _active_property_tab(self) -> str | None:
        if not hasattr(self, "main"):
            return None
        if self.main.tab(self.main.select(), "text") != "물건":
            return None
        if not hasattr(self, "inner_tabs"):
            return None
        return self.inner_tabs.tab(self.inner_tabs.select(), "text")

    def _active_search_entry(self):
        tab = self.main.tab(self.main.select(), "text") if hasattr(self, "main") else ""
        if tab == "물건" and hasattr(self, "property_search_entry"):
            return self.property_search_entry
        if tab == "고객" and hasattr(self, "customer_search_entry"):
            return self.customer_search_entry
        return None

    def _active_treeview(self) -> ttk.Treeview | None:
        tab = self.main.tab(self.main.select(), "text") if hasattr(self, "main") else ""
        if tab == "물건":
            active = self._active_property_tab()
            if active and active in getattr(self, "prop_trees", {}):
                return self.prop_trees[active]
        if tab == "고객":
            return self._current_customer_tree()
        return None

    def _setup_tree_style(self, tree: ttk.Treeview) -> None:
        try:
            tree.tag_configure("odd", background="#FFFFFF")
            tree.tag_configure("even", background="#F7FAFF")
        except Exception:
            pass

    def _snapshot_tree_state(self, tree: ttk.Treeview) -> tuple[tuple[str, ...], float]:
        selected = tuple(tree.selection())
        try:
            y = float(tree.yview()[0])
        except Exception:
            y = 0.0
        return selected, y

    def _restore_tree_state(self, tree: ttk.Treeview, selected: tuple[str, ...], y: float) -> None:
        if selected:
            existing = [iid for iid in selected if tree.exists(iid)]
            if existing:
                tree.selection_set(existing)
                tree.focus(existing[0])
                tree.see(existing[0])
            else:
                tree.selection_remove(tree.selection())
        try:
            tree.yview_moveto(max(0.0, min(1.0, y)))
        except Exception:
            pass

    def _wire_form_navigation(self, widgets: list[tk.Widget]) -> None:
        for w in widgets:
            if isinstance(w, tk.Text):
                continue
            def _next(_e, cur=w):
                try:
                    idx = widgets.index(cur)
                except Exception:
                    return None
                for nxt in widgets[idx + 1:]:
                    if isinstance(nxt, tk.Widget) and nxt.winfo_exists() and str(nxt.cget("state") if hasattr(nxt, "cget") else "") != "disabled":
                        nxt.focus_set()
                        return "break"
                return "break"
            w.bind("<Return>", _next, add="+")

    def _collect_form_widgets(self, container: tk.Widget) -> list[tk.Widget]:
        targets: list[tuple[int, int, tk.Widget]] = []
        for child in container.winfo_children():
            if isinstance(child, (ttk.Entry, ttk.Combobox, tk.Text)):
                gi = child.grid_info() if hasattr(child, "grid_info") else {}
                r = int(gi.get("row", 999)) if gi else 999
                c = int(gi.get("column", 999)) if gi else 999
                targets.append((r, c, child))
            if isinstance(child, (ttk.Frame, ttk.LabelFrame, tk.Frame)):
                for sub in child.winfo_children():
                    if isinstance(sub, (ttk.Entry, ttk.Combobox, tk.Text)):
                        gi = sub.grid_info() if hasattr(sub, "grid_info") else {}
                        r = int(gi.get("row", 999)) if gi else 999
                        c = int(gi.get("column", 999)) if gi else 999
                        targets.append((r, c, sub))
        targets.sort(key=lambda x: (x[0], x[1]))
        return [w for _, _, w in targets]

    def _set_initial_focus(self, win: tk.Toplevel, widget: tk.Widget | None) -> None:
        if not widget:
            return
        win.after(30, lambda: widget.focus_set())

    def _attach_tooltip(self, widget: tk.Widget, text_getter) -> None:
        state = {"tip": None, "job": None}

        def _show():
            try:
                txt = str(text_getter() or "").strip()
                if not txt:
                    return
                tip = tk.Toplevel(widget)
                tip.wm_overrideredirect(True)
                x = widget.winfo_rootx() + 8
                y = widget.winfo_rooty() + widget.winfo_height() + 6
                tip.wm_geometry(f"+{x}+{y}")
                lbl = tk.Label(tip, text=txt, justify="left", bg="#111827", fg="#FFFFFF", padx=8, pady=5, wraplength=480)
                lbl.pack()
                state["tip"] = tip
            except Exception:
                pass

        def _enter(_e=None):
            try:
                state["job"] = widget.after(500, _show)
            except Exception:
                pass

        def _leave(_e=None):
            job = state.get("job")
            if job:
                try:
                    widget.after_cancel(job)
                except Exception:
                    pass
                state["job"] = None
            tip = state.get("tip")
            if tip:
                try:
                    tip.destroy()
                except Exception:
                    pass
                state["tip"] = None

        widget.bind("<Enter>", _enter, add="+")
        widget.bind("<Leave>", _leave, add="+")

    def run_with_busy_ui(self, fn, *, busy_message: str = "처리 중…", success_message: str = "완료", fail_prefix: str = "실패"):
        prev = str(self.sync_status_var.get()) if hasattr(self, "sync_status_var") else ""
        changed: list[tk.Widget] = []
        try:
            self.root.config(cursor="watch")
            if hasattr(self, "sync_status_var"):
                self.sync_status_var.set(busy_message)
            for attr in ("undo_btn",):
                w = getattr(self, attr, None)
                if isinstance(w, tk.Widget):
                    try:
                        if str(w.cget("state")) != "disabled":
                            w.configure(state="disabled")
                            changed.append(w)
                    except Exception:
                        pass
            self.root.update_idletasks()
            result = fn()
            if hasattr(self, "sync_status_var"):
                self.sync_status_var.set(success_message)
            return result
        except Exception as exc:
            if hasattr(self, "sync_status_var"):
                self.sync_status_var.set(f"{fail_prefix}: {exc}")
            raise
        finally:
            self.root.config(cursor="")
            for w in changed:
                try:
                    w.configure(state="normal")
                except Exception:
                    pass
            self.root.update_idletasks()
            if prev and hasattr(self, "root"):
                self.root.after(2500, lambda p=prev: hasattr(self, "sync_status_var") and self.sync_status_var.set(p))

    def _remember_undo(self, payload: dict | None) -> None:
        self._last_undo_payload = payload
        if hasattr(self, "undo_btn"):
            self.undo_btn.configure(state=("normal" if payload else "disabled"))

    def _undo_last_action(self):
        payload = self._last_undo_payload or {}
        if not payload:
            return
        action = payload.get("action")
        try:
            if action == "PROPERTY_DELETE":
                restore_property(int(payload.get("id")))
            elif action == "CUSTOMER_DELETE":
                restore_customer(int(payload.get("id")))
            elif action == "PROPERTY_HIDE":
                update_property(int(payload.get("id")), {"hidden": payload.get("before_hidden", 0)})
            elif action == "CUSTOMER_HIDE":
                update_customer(int(payload.get("id")), {"hidden": payload.get("before_hidden", 0)})
            elif action == "PROPERTY_STATUS":
                update_property(int(payload.get("id")), {"status": payload.get("before_status", "신규등록")})
            elif action == "CUSTOMER_STATUS":
                update_customer(int(payload.get("id")), {"status": payload.get("before_status", "문의")})
            else:
                return
            self.request_sync(reason="UNDO")
            self.refresh_all()
        finally:
            self._remember_undo(None)

    def _bind_global_shortcuts(self):
        self.root.bind_all("<Control-f>", self._focus_active_search)
        self.root.bind_all("/", self._focus_active_search)
        self.root.bind_all("<Delete>", self._delete_active_selection)
        self.root.bind_all("<Return>", self._open_active_selection)

    def _focus_active_search(self, _event=None):
        entry = self._active_search_entry()
        if entry is not None:
            entry.focus_set()
            entry.icursor("end")
            return "break"
        return None

    def _delete_active_selection(self, _event=None):
        w = self.root.focus_get()
        if isinstance(w, (tk.Entry, tk.Text, ttk.Entry, ttk.Combobox, tk.Spinbox)):
            return None
        tab = self.main.tab(self.main.select(), "text") if hasattr(self, "main") else ""
        if tab == "물건":
            active = self._active_property_tab()
            if active:
                self.delete_selected_property(active)
                return "break"
        if tab == "고객":
            self.delete_selected_customer()
            return "break"
        return None

    def _open_active_selection(self, _event=None):
        w = self.root.focus_get()
        if isinstance(w, (tk.Entry, tk.Text, ttk.Entry, ttk.Combobox, tk.Spinbox)):
            return None
        tree = self._active_treeview()
        if tree is None:
            return None
        tab = self.main.tab(self.main.select(), "text") if hasattr(self, "main") else ""
        if tab == "물건":
            active = self._active_property_tab()
            if active:
                self.open_selected_property_detail(active)
                return "break"
        if tab == "고객":
            self.open_selected_customer_detail()
            return "break"
        return None

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

        auto_sync_on_raw = get_setting("AUTO_SYNC_ON", "1").strip()
        auto_sync_on = auto_sync_on_raw not in {"0", "false", "False", "off", "OFF"}
        try:
            auto_sync_sec = max(10, int(get_setting("AUTO_SYNC_SEC", "60").strip() or "60"))
        except Exception:
            auto_sync_sec = 60

        export_mode_raw = get_setting("EXPORT_MODE", "single").strip() or "single"
        export_mode = "advanced" if export_mode_raw in {"advanced", "고급"} else "single"
        export_ics_raw = get_setting("EXPORT_ICS", "0").strip()
        export_ics = export_ics_raw in {"1", "true", "True", "on", "ON"}

        return AppSettings(
            sync_dir=sync_dir_path,
            webhook_url=webhook_url,
            auto_sync_sec=auto_sync_sec,
            auto_sync_on=auto_sync_on,
            export_mode=export_mode,
            export_ics=export_ics,
        )

    def _persist_settings(self) -> None:
        set_setting("GOOGLE_DRIVE_SYNC_DIR", str(self.settings.sync_dir))
        set_setting("GOOGLE_SHEETS_WEBHOOK_URL", self.settings.webhook_url)
        set_setting("AUTO_SYNC_ON", "1" if self.settings.auto_sync_on else "0")
        set_setting("AUTO_SYNC_SEC", str(max(10, int(self.settings.auto_sync_sec or 60))))
        set_setting("EXPORT_MODE", self.settings.export_mode)
        set_setting("EXPORT_ICS", "1" if self.settings.export_ics else "0")

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

        tcols = ("due_at", "title", "entity", "status")
        self.tasks_tree = ttk.Treeview(task_frame, columns=tcols, show="headings", height=18)
        twidths = {"due_at": 170, "title": 470, "entity": 360, "status": 100}
        tlabels = {"due_at": "일정일시", "title": "할 일", "entity": "대상", "status": "상태"}
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
        if col not in {"due_at", "title", "entity", "status"}:
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
                return property_label(p)
        if et == "CUSTOMER":
            c = get_customer(eid_i)
            if c:
                return customer_label(c)
        if et == "VIEWING":
            v = get_viewing(eid_i)
            if v and v.get("property_id"):
                p = get_property(int(v.get("property_id")))
                if p:
                    return property_label(p)
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
                iid=str(r.get("id")),
                values=(
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
            return int(sel[0])
        except Exception:
            return None

    def mark_selected_task_done(self):
        tid = self._selected_task_id()
        if tid is None:
            return
        mark_task_done(tid)
        self.request_sync(reason="TASK_DONE")
        self.refresh_tasks()
        self.refresh_dashboard()

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
        self._fit_toplevel(win, 700, 460)

        create_next = tk.BooleanVar(value=True)
        next_type = tk.StringVar(value=options[0] if options else "기타")
        note_var = tk.StringVar(value="")

        now = datetime.now()
        y = tk.IntVar(value=now.year)
        m = tk.IntVar(value=now.month)
        d = tk.IntVar(value=now.day + 1 if now.day < 28 else now.day)
        hh = tk.IntVar(value=10)
        mm = tk.IntVar(value=0)

        selected_customer_id: dict[str, int | None] = {"value": None}
        selected_property_id: dict[str, int | None] = {"value": None}
        selected_customer_label = tk.StringVar(value="미선택")
        selected_property_label = tk.StringVar(value="미선택")

        et = str(task_row.get("entity_type") or "")
        eid = task_row.get("entity_id")
        if et == "CUSTOMER" and eid is not None:
            c = get_customer(int(eid))
            if c:
                selected_customer_id["value"] = int(eid)
                selected_customer_label.set(customer_label(c))
        if et == "PROPERTY" and eid is not None:
            p0 = get_property(int(eid))
            if p0:
                selected_property_id["value"] = int(eid)
                selected_property_label.set(property_label(p0))

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

        ttk.Label(frm, text="고객(선택)").grid(row=4, column=0, sticky="e", padx=6, pady=6)
        cwrap = ttk.Frame(frm)
        cwrap.grid(row=4, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(cwrap, textvariable=selected_customer_label, width=34).pack(side="left")

        def pick_customer():
            rows = [c for c in list_customers(include_deleted=False) if not c.get("hidden")]
            pop = tk.Toplevel(win)
            pop.title("고객 선택")
            self._fit_toplevel(pop, 760, 560)
            tree = ttk.Treeview(pop, columns=("phone", "name", "status"), show="headings", height=12)
            for key, label in (("phone", "전화번호"), ("name", "이름"), ("status", "상태")):
                tree.heading(key, text=label)
            tree.pack(fill="both", expand=True, padx=8, pady=8)
            for r in rows:
                tree.insert("", "end", iid=str(r.get("id")), values=(fmt_phone(r.get("phone")), r.get("customer_name") or "", r.get("status") or ""))

            def done():
                sel = tree.selection()
                if sel:
                    cid = int(sel[0])
                    c = get_customer(cid)
                    selected_customer_id["value"] = cid
                    selected_customer_label.set(customer_label(c or {}))
                pop.destroy()

            ttk.Button(pop, text="선택", command=done).pack(pady=6)

        ttk.Button(cwrap, text="고객선택", command=pick_customer).pack(side="left", padx=4)
        ttk.Button(cwrap, text="해제", command=lambda: (selected_customer_id.__setitem__("value", None), selected_customer_label.set("미선택"))).pack(side="left", padx=2)

        ttk.Label(frm, text="물건(선택)").grid(row=5, column=0, sticky="e", padx=6, pady=6)
        pwrap = ttk.Frame(frm)
        pwrap.grid(row=5, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(pwrap, textvariable=selected_property_label, width=34).pack(side="left")

        def pick_property():
            rows = [p for p in list_properties(include_deleted=False) if not p.get("hidden")]
            pop = tk.Toplevel(win)
            pop.title("물건 선택")
            self._fit_toplevel(pop, 860, 620)
            tree = ttk.Treeview(pop, columns=("complex", "dongho", "status"), show="headings", height=12)
            for key, label in (("complex", "단지"), ("dongho", "동호"), ("status", "상태")):
                tree.heading(key, text=label)
            tree.pack(fill="both", expand=True, padx=8, pady=8)
            for r in rows:
                tree.insert("", "end", iid=str(r.get("id")), values=(r.get("complex_name") or "", str(r.get("address_detail") or "").strip() or property_label(r), r.get("status") or ""))

            def done():
                sel = tree.selection()
                if sel:
                    pid = int(sel[0])
                    p1 = get_property(pid)
                    selected_property_id["value"] = pid
                    selected_property_label.set(property_label(p1 or {}))
                pop.destroy()

            ttk.Button(pop, text="선택", command=done).pack(pady=6)

        ttk.Button(pwrap, text="물건선택", command=pick_property).pack(side="left", padx=4)
        ttk.Button(pwrap, text="해제", command=lambda: (selected_property_id.__setitem__("value", None), selected_property_label.set("미선택"))).pack(side="left", padx=2)

        ttk.Label(frm, text="메모").grid(row=6, column=0, sticky="e", padx=6, pady=6)
        ttk.Entry(frm, textvariable=note_var, width=40).grid(row=6, column=1, sticky="w", padx=6, pady=6)
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
            note_text = note_var.get().strip()

            if mapped_title == "약속 어레인지":
                if selected_customer_id["value"] is None or selected_property_id["value"] is None:
                    messagebox.showwarning("확인", "약속 어레인지는 고객/물건을 모두 선택해주세요.")
                    return
                entity_type = "PROPERTY"
                entity_id = int(selected_property_id["value"])
                title_to_save = "약속 어레인지"
                customer_note = selected_customer_label.get().strip()
                if customer_note and customer_note != "미선택":
                    note_text = f"{note_text}\n고객: {customer_note}".strip()
            elif mapped_title in ("광고 등록", "후속(서류/정산/보관)") and selected_property_id["value"] is not None:
                entity_type = "PROPERTY"
                entity_id = int(selected_property_id["value"])

            try:
                add_task(title=title_to_save, due_at=due_at, entity_type=entity_type, entity_id=entity_id, note=note_text, kind="MANUAL", status="OPEN")
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
            self.request_sync(reason="TASK_DONE")
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
        if row and self.enable_handoff:
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
        target_type_var = tk.StringVar(value="매물")
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
        ttk.Combobox(frm, textvariable=target_type_var, state="readonly", width=24, values=["매물", "고객", "없음"]).grid(row=3, column=1, sticky="w", padx=6, pady=6)
        batch_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="여러 매물에 같은 할일 생성(고급)", variable=batch_var).grid(row=4, column=1, sticky="w", padx=6, pady=2)

        def render_summary():
            c_label = "-"
            if selected_customer:
                c = get_customer(int(selected_customer))
                c_label = customer_label(c or {})
            p_label = "-"
            if selected_properties:
                labels = []
                for pid in selected_properties:
                    prow = get_property(int(pid))
                    if prow:
                        labels.append(property_label(prow))
                p_label = ", ".join(labels) if labels else "-"
            selected_summary.set(f"고객: {c_label} / 매물: {p_label}")

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
            tree = ttk.Treeview(popup, columns=("pick", "complex", "label", "deal", "status"), show="headings")
            tree.heading("pick", text="선택")
            tree.heading("complex", text="단지")
            tree.heading("label", text="동/호")
            tree.heading("deal", text="가격")
            tree.heading("status", text="상태")
            tree.column("pick", width=56, anchor="center")
            tree.column("complex", width=170)
            tree.column("label", width=220)
            tree.column("deal", width=200)
            tree.column("status", width=90)
            tree.pack(fill="both", expand=True, padx=8, pady=8)

            item_to_id: dict[str, int] = {}

            def render():
                item_to_id.clear()
                for item in tree.get_children():
                    tree.delete(item)
                q = search_var.get().strip().lower()
                for r in rows:
                    hay = f"{r.get('complex_name','')} {r.get('dong','')} {r.get('ho','')} {r.get('address_detail','')} {r.get('owner_phone','')}".lower()
                    if q and q not in hay:
                        continue
                    pid = int(r.get("id") or 0)
                    mark = "☑" if pid in picked_ids else "☐"
                    iid = tree.insert("", "end", values=(mark, r.get("complex_name") or r.get("tab"), property_label(r), self._calc_price_summary(r), r.get("status")))
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
                if not batch_var.get() and len(selected_properties) > 1:
                    selected_properties = selected_properties[:1]
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
            tree = ttk.Treeview(popup, columns=("phone", "name", "status"), show="headings")
            tree.heading("phone", text="전화번호")
            tree.heading("name", text="이름")
            tree.heading("status", text="상태")
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
                    tree.insert("", "end", iid=str(r.get("id")), values=(fmt_phone(r.get("phone")), r.get("customer_name"), r.get("status") or ""))

            search_var.trace_add("write", lambda *_: render())
            render()

            def done():
                nonlocal selected_customer
                sel = tree.selection()
                if not sel:
                    return
                selected_customer = int(sel[0])
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

            target = target_type_var.get().strip()
            if target == "고객":
                if not selected_customer:
                    messagebox.showwarning("확인", "고객을 1명 선택해주세요.")
                    return
                add_task(title=title, due_at=due_at, entity_type="CUSTOMER", entity_id=selected_customer, note=note_var.get().strip(), kind="MANUAL", status="OPEN")
                self.request_sync(reason="TASK_ADD")
            elif target == "매물":
                if not selected_properties:
                    messagebox.showwarning("확인", "매물을 선택해주세요.")
                    return
                pids = selected_properties if batch_var.get() else selected_properties[:1]
                for pid in pids:
                    add_task(title=title, due_at=due_at, entity_type="PROPERTY", entity_id=pid, note=note_var.get().strip(), kind="MANUAL", status="OPEN")
                self.request_sync(reason="TASK_ADD")
            else:
                add_task(title=title, due_at=due_at, entity_type=None, entity_id=None, note=note_var.get().strip(), kind="MANUAL", status="OPEN")
                self.request_sync(reason="TASK_ADD")

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

        ttk.Button(top, text="+ 물건 등록", style="Secondary.TButton", command=self.open_property_wizard).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="내보내기/동기화", style="Secondary.TButton", command=self.export_sync).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="숨김함", style="Secondary.TButton", command=self.open_hidden_properties_window).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="PDF+사진 패킹", style="Primary.TButton", command=lambda: self._generate_proposal_from_property_tab()).pack(side="left", padx=4, pady=6)
        self.undo_btn = ttk.Button(top, text="Undo", style="Secondary.TButton", command=self._undo_last_action, state="disabled")
        self.undo_btn.pack(side="left", padx=4, pady=6)

        preset_wrap = ttk.Frame(top)
        preset_wrap.pack(side="left", padx=(10, 4))
        ttk.Label(preset_wrap, text="프리셋").pack(side="left", padx=(0, 4))
        for slot in range(1, 6):
            btn = ttk.Button(preset_wrap, text=f"P{slot}", command=lambda s=slot: self._apply_preset(s))
            btn.pack(side="left", padx=2)
            btn.bind("<Button-3>", lambda e, s=slot: self._open_preset_menu(e, s))
            self._preset_buttons[slot] = btn

        ttk.Label(top, text="거래유형").pack(side="left", padx=(16, 4))
        for deal in ["매매", "전세", "월세"]:
            var = tk.BooleanVar(value=True)
            self.property_deal_filter_vars[deal] = var
            ttk.Checkbutton(top, text=deal, variable=var, command=self.refresh_properties).pack(side="left", padx=2)

        ttk.Label(top, text="검색").pack(side="left", padx=(24, 4))
        self.property_search_var = tk.StringVar(value="")
        ent = ttk.Entry(top, textvariable=self.property_search_var, width=30)
        self.property_search_entry = ent
        ent.pack(side="left", padx=4)
        ent.bind("<KeyRelease>", lambda _e: self.refresh_properties())
        ent.bind("<Return>", lambda _e: self.refresh_properties())
        ttk.Button(top, text="🔍 검색", style="Secondary.TButton", command=self.refresh_properties).pack(side="left", padx=4)

        self.inner_tabs = ttk.Notebook(self.property_tab)
        self.inner_tabs.pack(fill="both", expand=True, padx=10, pady=8)
        self.inner_tabs.bind("<<NotebookTabChanged>>", self._on_property_tab_changed)
        self.prop_trees: dict[str, ttk.Treeview] = {}

        cols = ("status", "complex_name", "address_detail", "contact", "unit_type", "floor", "price_summary", "updated_at")
        col_defs = [
            ("status", 86),
            ("complex_name", 165),
            ("address_detail", 130),
            ("contact", 120),
            ("unit_type", 88),
            ("floor", 58),
            ("price_summary", 170),
            ("updated_at", 130),
        ]
        col_labels = {
            "status": "상태",
            "complex_name": "단지",
            "address_detail": "동/호",
            "contact": "연락처",
            "unit_type": "타입",
            "floor": "층",
            "price_summary": "가격요약",
            "updated_at": "업데이트",
        }

        for tab_name in PROPERTY_TABS:
            frame = ttk.Frame(self.inner_tabs)
            self.inner_tabs.add(frame, text=tab_name)

            tree = ttk.Treeview(frame, columns=cols, show="headings", height=17)
            for c, w in col_defs:
                tree.heading(c, text=col_labels.get(c, c), command=lambda col=c, t=tree: self.sort_tree(t, col))
                tree.column(c, width=w, stretch=True)
            tree.pack(fill="both", expand=True)
            self._setup_tree_style(tree)
            xsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=xsb.set)
            xsb.pack(fill="x")
            tree.bind("<Double-1>", lambda e, t=tab_name: self._on_double_click_property(e, t))
            tree.bind("<<TreeviewSelect>>", lambda _e, t=tree: self._set_current_property(get_property(self._selected_id_from_tree(t)) if self._selected_id_from_tree(t) else None))
            self.prop_trees[tab_name] = tree

            btns = ttk.Frame(frame)
            btns.pack(fill="x", pady=4)
            ttk.Button(btns, text="상세", command=lambda t=tab_name: self.open_selected_property_detail(t)).pack(side="left", padx=4)
            ttk.Button(btns, text="숨김/보임", command=lambda t=tab_name: self.toggle_selected_property(t)).pack(side="left", padx=4)
            ttk.Button(btns, text="🗑️ 삭제", style="Danger.TButton", command=lambda t=tab_name: self.delete_selected_property(t)).pack(side="left", padx=4)

    def sort_tree(self, tree: ttk.Treeview, col: str):
        key = (id(tree), col)
        asc = not self.sort_state.get(key, False)
        items = list(tree.get_children(""))

        def val(item: str):
            raw = str(tree.set(item, col) or "")
            if col in {"id", "floor"}:
                digits = "".join(ch for ch in raw if ch.isdigit())
                try:
                    return int(digits or 0)
                except Exception:
                    return 0
            if col == "price_summary":
                nums = re.findall(r"\d+", raw)
                try:
                    return tuple(int(n) for n in nums) if nums else (0,)
                except Exception:
                    return (0,)
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

    def _get_selected_property_id(self, tab_name: str) -> int | None:
        tree = self.prop_trees.get(tab_name)
        if not tree:
            return None
        sel = tree.selection()
        if not sel:
            return None
        if len(sel) > 1:
            messagebox.showwarning("확인", "현재는 1개씩만 제안서 생성 가능합니다")
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    def _generate_proposal_from_property_tab(self):
        tab = self.inner_tabs.tab(self.inner_tabs.select(), "text")
        pid = self._get_selected_property_id(tab)
        if pid is None:
            messagebox.showwarning("확인", "물건 1개를 선택하세요")
            return
        self.generate_proposal_for_property(pid)

    def refresh_properties(self):
        current_tab = self.inner_tabs.tab(self.inner_tabs.select(), "text") if hasattr(self, "inner_tabs") else PROPERTY_TABS[0]
        q_raw = self.property_search_var.get().strip() if hasattr(self, "property_search_var") else ""
        self.property_search_terms_by_tab[current_tab] = q_raw
        selected_deals = {deal for deal, var in self.property_deal_filter_vars.items() if var.get()}

        for tab in PROPERTY_TABS:
            tree = self.prop_trees[tab]
            prev_sel, prev_y = self._snapshot_tree_state(tree)
            for i in tree.get_children():
                tree.delete(i)

            tab_query = self.property_search_terms_by_tab.get(tab, "")
            q = tab_query.lower()
            q_digits = "".join(ch for ch in tab_query if ch.isdigit())
            rows = [r for r in list_properties(tab) if not r.get("hidden")]

            if selected_deals:
                filtered: list[dict] = []
                for r in rows:
                    deal_flags = {
                        deal
                        for deal, enabled in (("매매", r.get("deal_sale")), ("전세", r.get("deal_jeonse")), ("월세", r.get("deal_wolse")))
                        if enabled
                    }
                    if deal_flags & selected_deals:
                        filtered.append(r)
                rows = filtered

            def _score(row: dict) -> tuple[int, str]:
                if not q:
                    return (2, str(row.get("updated_at") or ""))
                complex_name = str(row.get("complex_name") or "").lower()
                dong = str(row.get("dong") or "").lower()
                ho = str(row.get("ho") or "").lower()
                addr_detail = str(row.get("address_detail") or "").lower()
                primary = " ".join([complex_name, dong, ho, addr_detail])
                secondary = " ".join([str(row.get("owner_phone") or "").lower(), str(row.get("special_notes") or "").lower()])
                if q in primary or (q_digits and q_digits in "".join(ch for ch in primary if ch.isdigit())):
                    return (0, str(row.get("updated_at") or ""))
                if q in secondary or (q_digits and q_digits in "".join(ch for ch in secondary if ch.isdigit())):
                    return (1, str(row.get("updated_at") or ""))
                return (9, str(row.get("updated_at") or ""))

            rows.sort(key=lambda r: str(r.get("updated_at") or ""), reverse=True)
            if q:
                rows = [r for r in rows if _score(r)[0] < 9]
                rows.sort(key=lambda r: _score(r)[0])

            for idx, row in enumerate(rows):
                tree.insert(
                    "",
                    "end",
                    iid=str(row.get("id")),
                    tags=(("even" if idx % 2 == 0 else "odd"),),
                    values=(
                        row.get("status"),
                        row.get("complex_name"),
                        row.get("address_detail"),
                        self._resolve_property_contact(row),
                        row.get("unit_type"),
                        row.get("floor"),
                        self._calc_price_summary(row),
                        row.get("updated_at"),
                    ),
                )

            self._restore_tree_state(tree, prev_sel, prev_y)

    def _on_property_tab_changed(self, _event=None):
        current_tab = self.inner_tabs.tab(self.inner_tabs.select(), "text") if hasattr(self, "inner_tabs") else PROPERTY_TABS[0]
        self.property_search_var.set(self.property_search_terms_by_tab.get(current_tab, ""))
        self.refresh_properties()

    def _preset_payload(self) -> dict:
        current_tab = self.inner_tabs.tab(self.inner_tabs.select(), "text") if hasattr(self, "inner_tabs") else PROPERTY_TABS[0]
        return {
            "tab": current_tab,
            "search_by_tab": dict(self.property_search_terms_by_tab),
            "deals": {k: bool(v.get()) for k, v in self.property_deal_filter_vars.items()},
        }

    def _refresh_preset_buttons(self):
        presets = {int(r.get("slot")): r for r in list_presets()}
        for slot, btn in self._preset_buttons.items():
            row = presets.get(slot)
            btn.configure(text=(str(row.get("name") or f"P{slot}") if row else f"P{slot}"))

    def _save_preset_slot(self, slot: int, *, rename_only: bool = False):
        existing = next((r for r in list_presets() if int(r.get("slot")) == int(slot)), None)
        default_name = str(existing.get("name") if existing else f"P{slot}")
        name = simpledialog.askstring("프리셋", "프리셋 이름", initialvalue=default_name, parent=self.root)
        if not name:
            return
        payload = json.loads(existing.get("payload_json") or "{}") if (rename_only and existing) else self._preset_payload()
        upsert_preset(slot, name.strip(), payload)
        self._refresh_preset_buttons()

    def _apply_preset(self, slot: int):
        row = next((r for r in list_presets() if int(r.get("slot")) == int(slot)), None)
        if not row:
            self._save_preset_slot(slot)
            return
        payload = json.loads(row.get("payload_json") or "{}")
        for k, v in payload.get("deals", {}).items():
            if k in self.property_deal_filter_vars:
                self.property_deal_filter_vars[k].set(bool(v))
        by_tab = payload.get("search_by_tab") or {}
        for tab in PROPERTY_TABS:
            self.property_search_terms_by_tab[tab] = str(by_tab.get(tab, ""))
        tab = payload.get("tab")
        if tab in PROPERTY_TABS and hasattr(self, "inner_tabs"):
            self.inner_tabs.select(PROPERTY_TABS.index(tab))
        self.property_search_var.set(self.property_search_terms_by_tab.get(self.inner_tabs.tab(self.inner_tabs.select(), "text"), ""))
        self.refresh_properties()

    def _open_preset_menu(self, event, slot: int):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="현재 상태 저장", command=lambda s=slot: self._save_preset_slot(s))
        menu.add_command(label="이름 변경", command=lambda s=slot: self._save_preset_slot(s, rename_only=True))
        menu.add_command(label="삭제", command=lambda s=slot: (delete_preset(s), self._refresh_preset_buttons()))
        menu.tk_popup(event.x_root, event.y_root)

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
            "price_sale_eok": tk.StringVar(value="0"),
            "price_sale_che": tk.StringVar(value="0"),
            "price_jeonse_eok": tk.StringVar(value="0"),
            "price_jeonse_che": tk.StringVar(value="0"),
            "wolse_deposit_eok": tk.StringVar(value="0"),
            "wolse_deposit_che": tk.StringVar(value="0"),
            "wolse_rent_man": tk.StringVar(value="0"),
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

        eok_options = [str(i) for i in range(0, 301)]
        che_options = [str(i) for i in range(0, 10)]
        man_options = [str(i * 10) for i in range(0, 501)]  # 자주 쓰는 값(직접 입력 가능)

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
        ttk.Combobox(step2, textvariable=vars_["price_sale_eok"], values=eok_options, state="normal", width=6).grid(row=0, column=1)
        ttk.Label(step2, text="억").grid(row=0, column=2)
        ttk.Combobox(step2, textvariable=vars_["price_sale_che"], values=che_options, state="normal", width=6).grid(row=0, column=3)
        ttk.Label(step2, text="천만원").grid(row=0, column=4)

        ttk.Checkbutton(step2, text="전세", variable=vars_["deal_jeonse"]).grid(row=1, column=0, padx=6, pady=6, sticky="w")
        ttk.Combobox(step2, textvariable=vars_["price_jeonse_eok"], values=eok_options, state="normal", width=6).grid(row=1, column=1)
        ttk.Label(step2, text="억").grid(row=1, column=2)
        ttk.Combobox(step2, textvariable=vars_["price_jeonse_che"], values=che_options, state="normal", width=6).grid(row=1, column=3)
        ttk.Label(step2, text="천만원").grid(row=1, column=4)

        ttk.Checkbutton(step2, text="월세", variable=vars_["deal_wolse"]).grid(row=2, column=0, padx=6, pady=6, sticky="w")
        ttk.Combobox(step2, textvariable=vars_["wolse_deposit_eok"], values=eok_options, state="normal", width=6).grid(row=2, column=1)
        ttk.Label(step2, text="억").grid(row=2, column=2)
        ttk.Combobox(step2, textvariable=vars_["wolse_deposit_che"], values=che_options, state="normal", width=6).grid(row=2, column=3)
        ttk.Label(step2, text="천만원").grid(row=2, column=4)
        ttk.Combobox(step2, textvariable=vars_["wolse_rent_man"], values=man_options, state="normal", width=8).grid(row=2, column=5)
        ttk.Label(step2, text="만원").grid(row=2, column=6)
        ttk.Label(step2, text="(만원 단위, 직접 입력 가능)", foreground="#666").grid(row=2, column=7, padx=4, sticky="w")

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
                    nums += [vars_["price_sale_eok"].get(), vars_["price_sale_che"].get()]
                if vars_["deal_jeonse"].get():
                    nums += [vars_["price_jeonse_eok"].get(), vars_["price_jeonse_che"].get()]
                if vars_["deal_wolse"].get():
                    nums += [vars_["wolse_deposit_eok"].get(), vars_["wolse_deposit_che"].get(), vars_["wolse_rent_man"].get()]
                if any(not _is_non_negative_int(n) for n in nums):
                    messagebox.showwarning("입력 확인", "가격(억/천만원, 만원)은 숫자만 입력해주세요.")
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

            data["price_sale_eok"] = int(vars_["price_sale_eok"].get() or 0)
            data["price_sale_che"] = int(vars_["price_sale_che"].get() or 0)
            data["price_jeonse_eok"] = int(vars_["price_jeonse_eok"].get() or 0)
            data["price_jeonse_che"] = int(vars_["price_jeonse_che"].get() or 0)
            data["wolse_deposit_eok"] = int(vars_["wolse_deposit_eok"].get() or 0)
            data["wolse_deposit_che"] = int(vars_["wolse_deposit_che"].get() or 0)
            data["wolse_rent_man"] = int(vars_["wolse_rent_man"].get() or 0)

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
                self.request_sync(reason="PROPERTY_ADD")
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
            return int(selected[0])
        except Exception:
            pass
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
        before = get_property(pid, include_deleted=True)
        toggle_property_hidden(pid)
        self._remember_undo({"action": "PROPERTY_HIDE", "id": pid, "before_hidden": int(before.get("hidden") or 0) if before else 0})
        self.request_sync(reason="PROPERTY_HIDE")
        self.refresh_all()
    
    def delete_selected_property(self, tab_name: str):
        tree = self.prop_trees[tab_name]
        pid = self._selected_id_from_tree(tree)
        if pid is None:
            return
        if messagebox.askyesno("확인", "해당 물건을 삭제 처리(복구 가능)할까요?"):
            soft_delete_property(pid)
            self._remember_undo({"action": "PROPERTY_DELETE", "id": pid})
            self.request_sync(reason="PROPERTY_DELETE")
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
        tree = ttk.Treeview(win, columns=("tab", "address"), show="headings", height=14)
        tree.heading("tab", text="구분")
        tree.heading("address", text="동/호")
        tree.pack(fill="both", expand=True, padx=8, pady=8)
        hidden_rows = [r for r in list_properties(include_deleted=False) if r.get("hidden")]
        for r in hidden_rows:
            tree.insert("", "end", iid=str(r.get("id")), values=(r.get("tab"), r.get("address_detail")))

        def _selected_id():
            sel = tree.selection()
            if not sel:
                return None
            try:
                return int(sel[0])
            except Exception:
                return None

        def restore():
            pid = _selected_id()
            if pid is None:
                return
            toggle_property_hidden(pid)
            self.request_sync(reason="PROPERTY_HIDE")
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
            self.request_sync(reason="PROPERTY_DELETE")
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

    def _parse_deal_types(self, text: str) -> set[str]:
        allowed = {"매매", "전세", "월세"}
        return {token.strip() for token in str(text or "").split(",") if token.strip() in allowed}

    def _join_deal_types(self, selected: list[str]) -> str:
        ordered = [d for d in ["매매", "전세", "월세"] if d in set(selected)]
        return ",".join(ordered)

    def _safe_file_component(self, text: str) -> str:
        cleaned = re.sub(r'[\\/:*?"<>|]+', "_", str(text or "").strip())
        cleaned = re.sub(r"\s+", "_", cleaned).strip("_")
        return cleaned or "미지정"

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

        ttk.Button(top, text="+ 고객 등록", style="Secondary.TButton", command=self.open_customer_wizard).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="숨김함", style="Secondary.TButton", command=self.open_hidden_customers_window).pack(side="left", padx=4, pady=6)
        ttk.Button(top, text="내보내기/동기화", style="Secondary.TButton", command=self.export_sync).pack(side="left", padx=4, pady=6)

        ttk.Label(top, text="거래유형").pack(side="left", padx=(18, 4))
        self.customer_deal_filter_var = tk.StringVar(value="전체")
        deal_cb = ttk.Combobox(top, textvariable=self.customer_deal_filter_var, values=["전체", "매매", "전세", "월세"], state="readonly", width=8)
        deal_cb.pack(side="left", padx=4)
        deal_cb.bind("<<ComboboxSelected>>", lambda _e: self.refresh_customers())

        ttk.Label(top, text="전화번호 검색").pack(side="left", padx=(18, 4))
        self.customer_phone_query = tk.StringVar(value="")
        ent = ttk.Entry(top, textvariable=self.customer_phone_query, width=24)
        self.customer_search_entry = ent
        ent.pack(side="left", padx=4)
        ent.bind("<KeyRelease>", lambda _e: self.refresh_customers())
        ent.bind("<Return>", lambda _e: self.refresh_customers())
        ttk.Button(top, text="🔍 검색", style="Secondary.TButton", command=self.refresh_customers).pack(side="left", padx=4)

        cols = ("customer_name", "phone", "preferred_tab", "deal_type", "budget", "size", "move_in", "floor_preference", "status", "updated_at")
        col_defs = [
            ("customer_name", 120),
            ("phone", 122),
            ("preferred_tab", 110),
            ("deal_type", 86),
            ("budget", 125),
            ("size", 85),
            ("move_in", 95),
            ("floor_preference", 78),
            ("status", 76),
            ("updated_at", 118),
        ]
        col_labels = {
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
                tree.column(c, width=w, stretch=True)
            tree.pack(fill="both", expand=True)
            self._setup_tree_style(tree)
            c_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=c_x.set)
            c_x.pack(fill="x", pady=(0, 6))
            tree.bind("<Double-1>", self._on_double_click_customer)
            tree.bind("<<TreeviewSelect>>", lambda _e, t=tree: self._set_current_customer(get_customer(self._selected_id_from_tree(t)) if self._selected_id_from_tree(t) else None))
            self.customer_trees[tab_name] = tree

        btns = ttk.Frame(self.customer_tab)
        btns.pack(fill="x", padx=10, pady=4)
        ttk.Button(btns, text="상세", command=self.open_selected_customer_detail).pack(side="left", padx=4)
        ttk.Button(btns, text="숨김/보임", command=self.toggle_selected_customer).pack(side="left", padx=4)
        ttk.Button(btns, text="🗑️ 삭제", style="Danger.TButton", command=self.delete_selected_customer).pack(side="left", padx=4)

    def _current_customer_tree(self) -> ttk.Treeview:
        current_tab = self.customer_nb.tab(self.customer_nb.select(), "text") if hasattr(self, "customer_nb") else "전체"
        return self.customer_trees.get(current_tab) or self.customer_trees.get("전체")

    def open_customer_wizard(self):
        win = tk.Toplevel(self.root)
        win.title("고객 등록")
        self._fit_toplevel(win, 820, 660)

        vars_ = {
            "customer_name": tk.StringVar(value=""),
            "phone": tk.StringVar(),
            "preferred_tab": tk.StringVar(value=""),
            "deal_type": tk.StringVar(value="매매"),
            "deal_sale": tk.BooleanVar(value=True),
            "deal_jeonse": tk.BooleanVar(value=False),
            "deal_wolse": tk.BooleanVar(value=False),
            "size_unit": tk.StringVar(value="㎡"),
            "size_value": tk.StringVar(),
            "budget_eok": tk.StringVar(value="0"),
            "budget_che": tk.StringVar(value="0"),
            "wolse_deposit_eok": tk.StringVar(value="0"),
            "wolse_deposit_che": tk.StringVar(value="0"),
            "wolse_rent_man": tk.StringVar(value="0"),
            "move_in_period": tk.StringVar(),
            "location_preference": tk.StringVar(),
            "view_preference": tk.StringVar(value="비중요"),
            "condition_preference": tk.StringVar(value="비중요"),
            "floor_preference": tk.StringVar(value="상관없음"),
            "has_pet": tk.StringVar(value="없음"),
            "extra_needs": tk.StringVar(),
            "status": tk.StringVar(value="문의"),
        }

        eok_options = [str(i) for i in range(0, 301)]
        che_options = [str(i) for i in range(0, 10)]
        man_options = [str(i * 10) for i in range(0, 501)]  # 자주 쓰는 값(직접 입력 가능)

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
        ttk.Label(s1, text="전화번호*").grid(row=0, column=0, padx=6, pady=8, sticky="e")
        ttk.Entry(s1, textvariable=vars_["phone"], width=32).grid(row=0, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s1, text="상태").grid(row=1, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s1, textvariable=vars_["status"], values=CUSTOMER_STATUS_VALUES, state="readonly", width=29).grid(row=1, column=1, padx=6, pady=8, sticky="w")

        # Step2
        ttk.Label(s2, text="희망 유형").grid(row=0, column=0, padx=6, pady=8, sticky="e")
        pref_tab_wrap = ttk.Frame(s2)
        pref_tab_wrap.grid(row=0, column=1, padx=6, pady=8, sticky="w")
        ttk.Entry(pref_tab_wrap, textvariable=vars_["preferred_tab"], width=24, state="readonly").pack(side="left")
        ttk.Button(pref_tab_wrap, text="선택", command=lambda: (lambda v: vars_["preferred_tab"].set(v) if v is not None else None)(self._open_tab_multi_select(win, vars_["preferred_tab"].get()))).pack(side="left", padx=4)

        ttk.Label(s2, text="거래유형").grid(row=1, column=0, padx=6, pady=8, sticky="e")
        deal_wrap = ttk.Frame(s2)
        deal_wrap.grid(row=1, column=1, padx=6, pady=8, sticky="w")
        ttk.Checkbutton(deal_wrap, text="매매", variable=vars_["deal_sale"]).pack(side="left", padx=2)
        ttk.Checkbutton(deal_wrap, text="전세", variable=vars_["deal_jeonse"]).pack(side="left", padx=2)
        ttk.Checkbutton(deal_wrap, text="월세", variable=vars_["deal_wolse"]).pack(side="left", padx=2)
        ttk.Label(s2, text="희망 크기").grid(row=2, column=0, padx=6, pady=8, sticky="e")
        size_wrap = ttk.Frame(s2)
        size_wrap.grid(row=2, column=1, padx=6, pady=8, sticky="w")
        ttk.Entry(size_wrap, textvariable=vars_["size_value"], width=18).pack(side="left")
        ttk.Combobox(size_wrap, textvariable=vars_["size_unit"], values=["㎡", "평"], state="readonly", width=8).pack(side="left", padx=4)
        ttk.Label(s2, text="매매/전세 예산").grid(row=3, column=0, padx=6, pady=8, sticky="e")
        b_wrap = ttk.Frame(s2)
        b_wrap.grid(row=3, column=1, padx=6, pady=8, sticky="w")
        ttk.Combobox(b_wrap, textvariable=vars_["budget_eok"], values=eok_options, state="normal", width=6).pack(side="left")
        ttk.Label(b_wrap, text="억").pack(side="left", padx=2)
        ttk.Combobox(b_wrap, textvariable=vars_["budget_che"], values=che_options, state="normal", width=6).pack(side="left")
        ttk.Label(b_wrap, text="천만원").pack(side="left", padx=2)

        ttk.Label(s2, text="월세 보증금").grid(row=4, column=0, padx=6, pady=8, sticky="e")
        d_wrap = ttk.Frame(s2)
        d_wrap.grid(row=4, column=1, padx=6, pady=8, sticky="w")
        ttk.Combobox(d_wrap, textvariable=vars_["wolse_deposit_eok"], values=eok_options, state="normal", width=6).pack(side="left")
        ttk.Label(d_wrap, text="억").pack(side="left", padx=2)
        ttk.Combobox(d_wrap, textvariable=vars_["wolse_deposit_che"], values=che_options, state="normal", width=6).pack(side="left")
        ttk.Label(d_wrap, text="천만원").pack(side="left", padx=2)

        ttk.Label(s2, text="월세액(만원)").grid(row=5, column=0, padx=6, pady=8, sticky="e")
        ttk.Combobox(s2, textvariable=vars_["wolse_rent_man"], values=man_options, state="normal", width=16).grid(row=5, column=1, padx=6, pady=8, sticky="w")
        ttk.Label(s2, text="만원 단위(직접 입력 가능)", foreground="#666").grid(row=5, column=2, padx=4, pady=8, sticky="w")
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
                if not vars_["phone"].get().strip():
                    messagebox.showwarning("입력 확인", "전화번호는 필수 입력입니다.")
                    return False
                return True
            if idx == 1:
                if not (vars_["deal_sale"].get() or vars_["deal_jeonse"].get() or vars_["deal_wolse"].get()):
                    messagebox.showwarning("입력 확인", "거래유형을 1개 이상 선택해주세요.")
                    return False
                nums = [
                    vars_["budget_eok"].get(), vars_["budget_che"].get(),
                    vars_["wolse_deposit_eok"].get(), vars_["wolse_deposit_che"].get(),
                    vars_["wolse_rent_man"].get(),
                ]
                if any(not _is_non_negative_int(n) for n in nums):
                    messagebox.showwarning("입력 확인", "예산(억/천만원, 만원)은 숫자만 입력해주세요.")
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
            payload: dict[str, object] = {}
            for key, tk_var in vars_.items():
                raw = tk_var.get()
                payload[key] = raw.strip() if isinstance(raw, str) else raw
            payload["customer_name"] = ""
            payload["move_in_period"] = f"{move_y.get():04d}-{move_m.get():02d}-{move_d.get():02d}" if move_mode.get() == "날짜선택" else move_mode.get().strip()
            selected_deals = [d for d, key in (("매매", "deal_sale"), ("전세", "deal_jeonse"), ("월세", "deal_wolse")) if bool(payload.get(key))]
            payload["deal_type"] = self._join_deal_types(selected_deals)
            deal = payload.get("deal_type", "")
            budget_10m = int(payload.get("budget_eok", "0") or 0) * 10 + int(payload.get("budget_che", "0") or 0)
            wolse_deposit_10m = int(payload.get("wolse_deposit_eok", "0") or 0) * 10 + int(payload.get("wolse_deposit_che", "0") or 0)
            wolse_rent_man = int(payload.get("wolse_rent_man", "0") or 0)
            if selected_deals == ["월세"]:
                payload["budget"] = f"월세 {payload.get('wolse_deposit_eok','0')}억 {payload.get('wolse_deposit_che','0')}천만원 / {wolse_rent_man}만원"
            elif selected_deals == ["전세"]:
                payload["budget"] = f"전세 {payload.get('budget_eok','0')}억 {payload.get('budget_che','0')}천만원"
            else:
                payload["budget"] = f"{deal or '미지정'} {payload.get('budget_eok','0')}억 {payload.get('budget_che','0')}천만원"
            if payload["size_unit"] == "㎡":
                payload["preferred_area"] = payload["size_value"]
                payload["preferred_pyeong"] = ""
            else:
                payload["preferred_pyeong"] = payload["size_value"]
                payload["preferred_area"] = ""

            payload["budget_10m"] = budget_10m
            payload["wolse_deposit_10m"] = wolse_deposit_10m
            payload["wolse_rent_10man"] = money_utils.man_to_ten_man(wolse_rent_man)

            try:
                add_customer(payload)
                self.request_sync(reason="CUSTOMER_ADD")
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

        tree_states: dict[str, tuple[tuple[str, ...], float]] = {}
        for name, tree in getattr(self, "customer_trees", {}).items():
            tree_states[name] = self._snapshot_tree_state(tree)
            for i in tree.get_children():
                tree.delete(i)

        for row in rows:
            if row.get("hidden"):
                continue
            row_deals = self._parse_deal_types(str(row.get("deal_type") or ""))
            if deal_filter != "전체" and deal_filter not in row_deals:
                continue

            size = ""
            if row.get("size_value"):
                size = f"{row.get('size_value')} {row.get('size_unit') or ''}".strip()
            elif row.get("preferred_area"):
                size = f"{row.get('preferred_area')} ㎡"
            elif row.get("preferred_pyeong"):
                size = f"{row.get('preferred_pyeong')} 평"

            shown_tab_text = row.get("preferred_tab") or ""
            values = (
                row.get("customer_name"),
                fmt_phone(row.get("phone")),
                shown_tab_text,
                row.get("deal_type") or "",
                row.get("budget") or "",
                size,
                row.get("move_in_period"),
                row.get("floor_preference"),
                row.get("status"),
                row.get("updated_at"),
            )

            if "전체" in self.customer_trees:
                idx_all = len(self.customer_trees["전체"].get_children())
                self.customer_trees["전체"].insert("", "end", iid=str(row.get("id")), tags=(("even" if idx_all % 2 == 0 else "odd"),), values=values)

            pref_tabs = {self._normalize_property_tab(t) for t in self._parse_preferred_tabs(str(row.get("preferred_tab") or ""))}
            for tab_name in PROPERTY_TABS:
                norm_tab = self._normalize_property_tab(tab_name)
                if norm_tab in pref_tabs and tab_name in self.customer_trees:
                    tab_values = list(values)
                    tab_values[2] = tab_name
                    idx_tab = len(self.customer_trees[tab_name].get_children())
                    self.customer_trees[tab_name].insert("", "end", iid=str(row.get("id")), tags=(("even" if idx_tab % 2 == 0 else "odd"),), values=tuple(tab_values))


        for name, tree in self.customer_trees.items():
            sel, y = tree_states.get(name, ((), 0.0))
            self._restore_tree_state(tree, sel, y)

    def toggle_selected_customer(self):
        cid = self._selected_id_from_tree(self._current_customer_tree())
        if cid is None:
            return
        before = get_customer(cid, include_deleted=True)
        toggle_customer_hidden(cid)
        self._remember_undo({"action": "CUSTOMER_HIDE", "id": cid, "before_hidden": int(before.get("hidden") or 0) if before else 0})
        self.refresh_all()
    
    def delete_selected_customer(self):
        cid = self._selected_id_from_tree(self._current_customer_tree())
        if cid is None:
            return
        if messagebox.askyesno("확인", "해당 고객 요청을 삭제 처리(복구 가능)할까요?"):
            soft_delete_customer(cid)
            self._remember_undo({"action": "CUSTOMER_DELETE", "id": cid})
            self.request_sync(reason="CUSTOMER_DELETE")
            self.refresh_all()
    
    def open_selected_customer_detail(self):
        cid = self._selected_id_from_tree(self._current_customer_tree())
        if cid is None:
            return
        self.open_customer_detail(cid)
    
    def _on_double_click_customer(self, event):
        tree = event.widget if isinstance(event.widget, ttk.Treeview) else self._current_customer_tree()
        item = tree.identify_row(event.y)
        if item:
            try:
                cid = int(item)
                self.open_customer_detail(cid)
                return
            except Exception:
                pass
        self.open_selected_customer_detail()

    def open_hidden_customers_window(self):
        win = tk.Toplevel(self.root)
        win.title("고객 숨김함")
        self._fit_toplevel(win, 760, 560)
        tree = ttk.Treeview(win, columns=("phone", "name", "status"), show="headings", height=14)
        tree.heading("phone", text="전화번호")
        tree.heading("name", text="이름")
        tree.heading("status", text="상태")
        tree.column("phone", width=180)
        tree.column("name", width=180)
        tree.column("status", width=120)
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        hidden_rows = [r for r in list_customers(include_deleted=False) if r.get("hidden")]
        for r in hidden_rows:
            tree.insert(
                "",
                "end",
                iid=str(r.get("id")),
                values=(fmt_phone(r.get("phone")), r.get("customer_name") or "", r.get("status") or ""),
            )

        def _selected_id():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("확인", "고객을 먼저 선택해주세요.", parent=win)
                return None
            try:
                return int(sel[0])
            except Exception:
                messagebox.showerror("오류", "선택된 고객 ID를 읽을 수 없습니다.", parent=win)
                return None

        def restore():
            cid = _selected_id()
            if cid is None:
                return
            toggle_customer_hidden(cid)
            self.request_sync(reason="CUSTOMER_HIDE")
            self.refresh_all()
            for item in tree.selection():
                tree.delete(item)

        def open_detail():
            cid = _selected_id()
            if cid is None:
                return
            self.open_customer_detail(cid)

        def delete_it():
            cid = _selected_id()
            if cid is None:
                return
            if not messagebox.askyesno("확인", "해당 고객을 삭제 처리(복구 가능)할까요?", parent=win):
                return
            soft_delete_customer(cid)
            self.request_sync(reason="CUSTOMER_DELETE")
            self.refresh_all()
            for item in tree.selection():
                tree.delete(item)

        btns = ttk.Frame(win)
        btns.pack(pady=6)
        ttk.Button(btns, text="복구", command=restore).pack(side="left", padx=4)
        ttk.Button(btns, text="정보보기", command=open_detail).pack(side="left", padx=4)
        ttk.Button(btns, text="삭제", command=delete_it).pack(side="left", padx=4)

    def _build_settings_ui(self):
        box = ttk.LabelFrame(self.settings_tab, text="동기화 설정(관리자)")
        box.pack(fill="x", padx=10, pady=10)

        self.set_sync_dir_var = tk.StringVar(value=str(self.settings.sync_dir))
        self.set_webhook_var = tk.StringVar(value=self.settings.webhook_url)
        self.set_auto_sync_on_var = tk.BooleanVar(value=bool(self.settings.auto_sync_on))
        self.set_auto_sync_sec_var = tk.IntVar(value=max(10, int(self.settings.auto_sync_sec or 60)))
        self.set_export_mode_var = tk.StringVar(value="단일파일 추천" if self.settings.export_mode == "single" else "고급(CSV/ICS 포함)")
        self.set_export_ics_var = tk.BooleanVar(value=bool(self.settings.export_ics))

        ttk.Label(box, text="Drive 동기화 폴더").grid(row=0, column=0, padx=6, pady=6, sticky="e")
        ttk.Entry(box, textvariable=self.set_sync_dir_var, width=80).grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ttk.Button(box, text="폴더 선택", command=self.browse_sync_dir).grid(row=0, column=2, padx=6, pady=6)

        ttk.Label(box, text="(선택) Sheets 웹훅 URL").grid(row=1, column=0, padx=6, pady=6, sticky="e")
        ttk.Entry(box, textvariable=self.set_webhook_var, width=80).grid(row=1, column=1, padx=6, pady=6, sticky="w")

        ttk.Checkbutton(box, text="자동 동기화 사용", variable=self.set_auto_sync_on_var).grid(row=2, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(box, text="자동 동기화 주기(초)").grid(row=3, column=0, padx=6, pady=6, sticky="e")
        ttk.Spinbox(box, from_=10, to=3600, increment=10, textvariable=self.set_auto_sync_sec_var, width=10).grid(row=3, column=1, padx=6, pady=6, sticky="w")

        ttk.Label(box, text="내보내기 모드").grid(row=4, column=0, padx=6, pady=6, sticky="e")
        ttk.Combobox(
            box,
            textvariable=self.set_export_mode_var,
            values=["단일파일 추천", "고급(CSV/ICS 포함)"],
            state="readonly",
            width=24,
        ).grid(row=4, column=1, padx=6, pady=6, sticky="w")
        ttk.Checkbutton(box, text="캘린더(ICS) 내보내기", variable=self.set_export_ics_var).grid(row=5, column=1, padx=6, pady=6, sticky="w")

        btns = ttk.Frame(box)
        btns.grid(row=6, column=0, columnspan=3, sticky="w", padx=6, pady=6)
        ttk.Button(btns, text="저장", command=self.save_settings).pack(side="left", padx=4)
        ttk.Button(btns, text="내보내기/동기화 실행", command=self.export_sync).pack(side="left", padx=4)
        ttk.Button(btns, text="폴더 열기", command=lambda: _open_folder(Path(self.set_sync_dir_var.get()).expanduser())).pack(side="left", padx=4)

        hint = ttk.Label(
            self.settings_tab,
            text=(
                "TIP: 자동 동기화(기본 60초) + 변경 후 즉시(5초 디바운스)로 버튼 없이도 동작합니다.\n"
                "     단일파일 모드에서는 exports/ledger_snapshot.json 1개만 생성됩니다."
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
        export_mode = "single" if self.set_export_mode_var.get() == "단일파일 추천" else "advanced"
        try:
            auto_sync_sec = max(10, int(self.set_auto_sync_sec_var.get()))
        except Exception:
            auto_sync_sec = 60

        self.settings = AppSettings(
            sync_dir=sync_dir,
            webhook_url=webhook,
            auto_sync_sec=auto_sync_sec,
            auto_sync_on=bool(self.set_auto_sync_on_var.get()),
            export_mode=export_mode,
            export_ics=bool(self.set_export_ics_var.get()),
        )
        self._persist_settings()
        self.start_auto_sync_loop()
        messagebox.showinfo("완료", "설정이 저장되었습니다.")

    def _export_sync_core(self) -> tuple[bool, str]:
        self.settings = self._load_settings()
        props = list_properties(include_deleted=False)
        custs = list_customers(include_deleted=False)
        photos = list_photos_all()
        viewings = list_viewings()
        tasks = list_tasks(include_done=False)

        return upload_visible_data(
            props,
            custs,
            photos=photos,
            viewings=viewings,
            tasks=tasks,
            settings=SyncSettings(
                webhook_url=self.settings.webhook_url,
                sync_dir=self.settings.sync_dir,
                export_mode=self.settings.export_mode,
                export_ics=self.settings.export_ics,
            ),
        )

    def _run_sync_in_thread(self, *, reason: str = "AUTO") -> None:
        if self._sync_inflight:
            return
        self._sync_inflight = True

        def worker():
            try:
                ok, msg = self._export_sync_core()
            except Exception as exc:
                ok, msg = False, str(exc)

            def ui_update():
                self._sync_inflight = False
                stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if hasattr(self, "sync_status_var"):
                    self.sync_status_var.set(f"마지막 동기화: {stamp} / {'성공' if ok else '실패'} ({reason})")

            self.root.after(0, ui_update)

        threading.Thread(target=worker, daemon=True).start()

    def request_sync(self, *, delay_ms: int = 5000, reason: str = "CHANGE") -> None:
        if self._sync_after_id is not None:
            try:
                self.root.after_cancel(self._sync_after_id)
            except Exception:
                pass
            self._sync_after_id = None

        def _fire():
            self._sync_after_id = None
            self._run_sync_in_thread(reason=reason)

        self._sync_after_id = self.root.after(max(100, int(delay_ms)), _fire)

    # -----------------
    # Sync / Export
    # -----------------
    def export_sync(self, *, silent: bool = False):
        ok, msg = self._export_sync_core()
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
        self.settings = self._load_settings()
        if self._auto_sync_after_id is not None:
            try:
                self.root.after_cancel(self._auto_sync_after_id)
            except Exception:
                pass
            self._auto_sync_after_id = None

        if not self.settings.auto_sync_on:
            return

        def _run():
            self.settings = self._load_settings()
            if not self.settings.auto_sync_on:
                self._auto_sync_after_id = None
                return
            self._run_sync_in_thread(reason="AUTO")
            interval = max(10, int(self.settings.auto_sync_sec or 60)) * 1000
            self._auto_sync_after_id = self.root.after(interval, _run)

        self._auto_sync_after_id = self.root.after(10_000, _run)

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

        self._set_current_property(row)
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
            "contact_display": tk.StringVar(value=self._resolve_property_contact(row)),
        }

        tab_basic.columnconfigure(0, weight=1)
        tab_basic.rowconfigure(0, weight=1)

        split = ttk.Panedwindow(tab_basic, orient="horizontal")
        split.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.Frame(split, style="TFrame")
        right = ttk.Frame(split, style="TFrame")
        split.add(left, weight=5)
        split.add(right, weight=2)

        form = ttk.LabelFrame(left, text="물건 정보")
        form.pack(fill="both", expand=True)

        def add_row(r, c, label, widget):
            ttk.Label(form, text=label).grid(row=r, column=c, padx=8, pady=6, sticky="e")
            widget.grid(row=r, column=c + 1, padx=8, pady=6, sticky="ew")

        # row 0
        add_row(0, 0, "탭", ttk.Entry(form, textvariable=vars_["tab"], width=24, state="readonly"))
        add_row(0, 2, "단지명", ttk.Entry(form, textvariable=vars_["complex_name"], width=36))
        add_row(0, 4, "상태", ttk.Combobox(form, textvariable=vars_["status"], width=18, state="readonly", values=PROPERTY_STATUS_VALUES))

        for c in (1, 3, 5):
            form.columnconfigure(c, weight=1)

        add_row(1, 0, "연락처", ttk.Entry(form, textvariable=vars_["contact_display"], width=24, state="readonly"))
        add_row(1, 2, "동/호(상세)", ttk.Entry(form, textvariable=vars_["address_detail"], width=24))
        add_row(1, 4, "면적타입", ttk.Entry(form, textvariable=vars_["unit_type"], width=18))

        add_row(2, 0, "층/총층", ttk.Entry(form, textvariable=vars_["floor"], width=10))
        add_row(2, 2, "면적(㎡)", ttk.Entry(form, textvariable=vars_["area"], width=18))
        add_row(2, 4, "평형", ttk.Entry(form, textvariable=vars_["pyeong"], width=18))

        add_row(3, 0, "컨디션", ttk.Combobox(form, textvariable=vars_["condition"], width=18, state="readonly", values=["상", "중", "하"]))
        add_row(3, 2, "조망/뷰", ttk.Entry(form, textvariable=vars_["view"], width=24))
        add_row(3, 4, "향", ttk.Entry(form, textvariable=vars_["orientation"], width=18))
        ttk.Checkbutton(form, text="수리필요", variable=vars_["repair_needed"]).grid(row=4, column=0, padx=8, pady=6, sticky="w")

        add_row(5, 0, "세입자정보", ttk.Entry(form, textvariable=vars_["tenant_info"], width=64))
        link_entry = ttk.Entry(form, textvariable=vars_["naver_link"], width=64)
        add_row(6, 0, "네이버링크", link_entry)

        notes_expanded = {"v": True}
        notes_frame = ttk.Frame(form)
        notes_frame.grid(row=7, column=0, columnspan=6, sticky="nsew")
        form.rowconfigure(7, weight=1)
        notes_frame.columnconfigure(1, weight=1)

        ttk.Label(notes_frame, text="특이사항").grid(row=0, column=0, padx=8, pady=6, sticky="ne")
        special_txt = tk.Text(notes_frame, height=5, wrap="word")
        special_txt.insert("1.0", str(vars_["special_notes"].get() or ""))
        special_txt.grid(row=0, column=1, padx=8, pady=6, sticky="nsew")

        ttk.Label(notes_frame, text="별도기재").grid(row=1, column=0, padx=8, pady=6, sticky="ne")
        note_txt = tk.Text(notes_frame, height=5, wrap="word")
        note_txt.insert("1.0", str(vars_["note"].get() or ""))
        note_txt.grid(row=1, column=1, padx=8, pady=6, sticky="nsew")
        notes_frame.rowconfigure(0, weight=1)
        notes_frame.rowconfigure(1, weight=1)

        side = ttk.LabelFrame(right, text="핵심 액션")
        side.pack(fill="y", expand=False, anchor="n")

        def save_changes():
            vars_["special_notes"].set(special_txt.get("1.0", "end").strip())
            vars_["note"].set(note_txt.get("1.0", "end").strip())
            data = {k: (v.get() if not isinstance(v, tk.BooleanVar) else bool(v.get())) for k, v in vars_.items()}
            before = get_property(property_id, include_deleted=True)
            update_property(property_id, data)
            if before and str(before.get("status") or "") != str(data.get("status") or ""):
                self._remember_undo({"action": "PROPERTY_STATUS", "id": property_id, "before_status": before.get("status")})
            self.request_sync(reason="PROPERTY_UPDATE")
            self.refresh_all()
            messagebox.showinfo("완료", "저장되었습니다.")

        def toggle_hide():
            before = get_property(property_id, include_deleted=True)
            toggle_property_hidden(property_id)
            self._remember_undo({"action": "PROPERTY_HIDE", "id": property_id, "before_hidden": int(before.get("hidden") or 0) if before else 0})
            self.request_sync(reason="PROPERTY_HIDE")
            messagebox.showinfo("완료", "상태가 변경되었습니다(숨김/보임).")
            self.refresh_all()

        def soft_delete():
            if messagebox.askyesno("확인", "삭제 처리(복구 가능) 하시겠습니까?"):
                soft_delete_property(property_id)
                self._remember_undo({"action": "PROPERTY_DELETE", "id": property_id})
                self.request_sync(reason="PROPERTY_DELETE")
                self.refresh_all()
                win.destroy()

        def open_link():
            url = str(vars_["naver_link"].get()).strip()
            if not url:
                return
            webbrowser.open(url)

        ttk.Button(side, text="💾 저장", style="Primary.TButton", command=save_changes).pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(side, text="📷 사진 탭 열기", style="Secondary.TButton", command=lambda: nb.select(tab_photos)).pack(fill="x", padx=8, pady=4)
        ttk.Button(side, text="PDF+사진 패킹", style="Secondary.TButton", command=lambda: self.generate_proposal_for_property(property_id)).pack(fill="x", padx=8, pady=4)
        ttk.Button(side, text="링크 열기", style="Secondary.TButton", command=open_link).pack(fill="x", padx=8, pady=4)
        ttk.Button(side, text="숨김/보임", style="Secondary.TButton", command=toggle_hide).pack(fill="x", padx=8, pady=4)
        ttk.Button(side, text="🗑️ 삭제", style="Danger.TButton", command=soft_delete).pack(fill="x", padx=8, pady=(4, 8))
        def toggle_notes():
            notes_expanded["v"] = not notes_expanded["v"]
            if notes_expanded["v"]:
                notes_frame.grid()
                notes_btn.configure(text="메모 접기")
            else:
                notes_frame.grid_remove()
                notes_btn.configure(text="메모 펼치기")

        notes_btn = ttk.Button(side, text="메모 접기", style="Secondary.TButton", command=toggle_notes)
        notes_btn.pack(fill="x", padx=8, pady=(0, 6))
        ttk.Button(side, text="할 일 추가", style="Secondary.TButton", command=lambda: self.open_add_task_window(default_entity_type="PROPERTY", default_entity_id=property_id)).pack(fill="x", padx=8, pady=(4, 8))

        self._attach_tooltip(link_entry, lambda: vars_["naver_link"].get())
        self._attach_tooltip(special_txt, lambda: special_txt.get("1.0", "end").strip())
        self._attach_tooltip(note_txt, lambda: note_txt.get("1.0", "end").strip())
        form_widgets = self._collect_form_widgets(form)
        self._wire_form_navigation(form_widgets)
        self._set_initial_focus(win, form_widgets[0] if form_widgets else None)

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
        photo_tags = ["거실", "주방", "화장실", "방", "현관", "외관", "조망", "커뮤니티", "기타"]
        ph_tag_var = tk.StringVar(value="")
        ttk.Label(ph_controls, text="사진 구분").pack(side="left", padx=4)
        ttk.Combobox(ph_controls, textvariable=ph_tag_var, values=photo_tags, state="readonly", width=14).pack(side="left", padx=4)

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
                messagebox.showwarning("확인", "사진 분류(태그)를 먼저 선택해주세요.")
                return
            try:
                dst = self._copy_photo_to_library(Path(src), property_id, tag=tag)
                add_photo(property_id, str(dst), tag=tag)
                current = get_property(property_id)
                if current and str(current.get("status") or "") == "신규등록":
                    update_property(property_id, {"status": "검수완료(사진등록)"})
                    self.request_sync(reason="PROPERTY_UPDATE")
                ph_tag_var.set("")
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

        ttk.Button(ph_controls, text="📷 사진등록", style="Primary.TButton", command=add_photo_ui).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="사진 보기", style="Secondary.TButton", command=open_photo).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="폴더 열기", command=lambda: _open_folder(Path(ph_tree.item(ph_tree.selection()[0], "values")[2]).parent) if ph_tree.selection() else None).pack(side="left", padx=4)
        ttk.Button(ph_controls, text="🗑️ 기록 삭제", style="Danger.TButton", command=remove_photo).pack(side="left", padx=4)
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
            self.request_sync(reason="VIEWING_UPDATE")
            refresh_viewings()
            self.refresh_all()

        def remove_viewing():
            sel = vw_tree.selection()
            if not sel:
                return
            vid = int(vw_tree.item(sel[0], "values")[0])
            if messagebox.askyesno("확인", "일정 기록을 삭제할까요?"):
                delete_viewing(vid)
                self.request_sync(reason="VIEWING_DELETE")
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
            "title": tk.StringVar(value="현장/상담"),
            "memo": tk.StringVar(value=""),
        }
        selected_customer_id: dict[str, int | None] = {"id": None}
        selected_customer_label = tk.StringVar(value="미선택")

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        def row(r, label, key, width=42):
            ttk.Label(frm, text=label).grid(row=r, column=0, padx=6, pady=6, sticky="e")
            ttk.Entry(frm, textvariable=vars_[key], width=width).grid(row=r, column=1, padx=6, pady=6, sticky="w")

        row(0, "시작(YYYY-MM-DD HH:MM)", "start")
        row(1, "종료(YYYY-MM-DD HH:MM)", "end")

        ttk.Label(frm, text="고객(선택)").grid(row=2, column=0, padx=6, pady=6, sticky="e")
        pick_wrap = ttk.Frame(frm)
        pick_wrap.grid(row=2, column=1, padx=6, pady=6, sticky="w")
        ttk.Label(pick_wrap, textvariable=selected_customer_label, width=28).pack(side="left")

        def pick_customer():
            pop = tk.Toplevel(win)
            pop.title("고객 선택")
            self._fit_toplevel(pop, 760, 560)

            search_var = tk.StringVar(value="")
            top = ttk.Frame(pop)
            top.pack(fill="x", padx=8, pady=(8, 4))
            ttk.Label(top, text="검색(전화/이름)").pack(side="left")
            ent = ttk.Entry(top, textvariable=search_var, width=24)
            ent.pack(side="left", padx=6)

            tree = ttk.Treeview(pop, columns=("phone", "name", "status"), show="headings", height=14)
            tree.heading("phone", text="전화번호")
            tree.heading("name", text="이름")
            tree.heading("status", text="상태")
            tree.column("phone", width=170, anchor="center")
            tree.column("name", width=170, anchor="w")
            tree.column("status", width=120, anchor="center")
            tree.pack(fill="both", expand=True, padx=8, pady=8)

            rows = [c for c in list_customers(include_deleted=False) if not c.get("hidden")]

            def refresh(*_):
                q = search_var.get().strip().lower()
                qd = "".join(ch for ch in q if ch.isdigit())
                for iid in tree.get_children():
                    tree.delete(iid)
                for c in rows:
                    phone_digits = "".join(ch for ch in str(c.get("phone") or "") if ch.isdigit())
                    hay = f"{c.get('customer_name','')} {c.get('phone','')}".lower()
                    if q and q not in hay and (not qd or qd not in phone_digits):
                        continue
                    cid = c.get("id")
                    if not cid:
                        continue
                    tree.insert("", "end", iid=str(cid), values=(fmt_phone(c.get("phone")), c.get("customer_name") or "", c.get("status") or ""))

            def done():
                sel = tree.selection()
                if not sel:
                    messagebox.showwarning("선택", "고객을 선택해주세요.")
                    return
                cid = int(sel[0])
                c = get_customer(cid) or {}
                selected_customer_id["id"] = cid
                selected_customer_label.set(customer_label(c))
                pop.destroy()

            ent.bind("<KeyRelease>", refresh)
            btns = ttk.Frame(pop)
            btns.pack(fill="x", padx=8, pady=(0, 8))
            ttk.Button(btns, text="선택", command=done).pack(side="left")
            ttk.Button(btns, text="취소", command=pop.destroy).pack(side="left", padx=6)
            tree.bind("<Double-1>", lambda _e: done())
            refresh()

        def clear_customer():
            selected_customer_id["id"] = None
            selected_customer_label.set("미선택")

        ttk.Button(pick_wrap, text="고객 선택", command=pick_customer).pack(side="left", padx=4)
        ttk.Button(pick_wrap, text="해제", command=clear_customer).pack(side="left")

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

            cid = selected_customer_id["id"]
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
            self.request_sync(reason="VIEWING_ADD")
            self.refresh_all()
            win.destroy()

        ttk.Button(frm, text="저장", command=save).grid(row=5, column=0, padx=6, pady=12)
        ttk.Button(frm, text="취소", command=win.destroy).grid(row=5, column=1, padx=6, pady=12, sticky="w")
    
    def open_customer_detail(self, customer_id: int):
        row = get_customer(customer_id)
        if not row:
            messagebox.showerror("오류", "고객 정보를 찾을 수 없습니다.")
            return

        self._set_current_customer(row)
        win = tk.Toplevel(self.root)
        win.title(f"고객 상세 - ID {customer_id}")
        self._fit_toplevel(win, 980, 700)

        vars_ = {k: tk.StringVar(value=str(row.get(k, "") or "")) for k in row.keys()}

        wrapper = ttk.Frame(win)
        wrapper.pack(fill="both", expand=True, padx=10, pady=10)
        wrapper.columnconfigure(0, weight=5)
        wrapper.columnconfigure(1, weight=2)
        wrapper.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(wrapper, text="고객 정보")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right = ttk.LabelFrame(wrapper, text="핵심 액션")
        right.grid(row=0, column=1, sticky="ns")

        fields = [
            ("고객명", "customer_name"), ("전화", "phone"), ("희망탭", "preferred_tab"),
            ("희망면적", "preferred_area"), ("희망평형", "preferred_pyeong"), ("예산", "budget"),
            ("기간", "move_in_period"), ("뷰", "view_preference"), ("위치", "location_preference"),
            ("층수선호", "floor_preference"), ("상태", "status"),
        ]

        for col in (1, 3, 5):
            left.columnconfigure(col, weight=1)

        for i, (label, key) in enumerate(fields):
            ttk.Label(left, text=label).grid(row=i // 3, column=(i % 3) * 2, padx=8, pady=6, sticky="e")
            if key == "preferred_tab":
                wrap = ttk.Frame(left)
                ttk.Entry(wrap, textvariable=vars_[key], width=20, state="readonly").pack(side="left")
                ttk.Button(wrap, text="선택", style="Secondary.TButton", command=lambda k=key: (lambda v: vars_[k].set(v) if v is not None else None)(self._open_tab_multi_select(win, vars_[k].get()))).pack(side="left", padx=2)
                w = wrap
            elif key == "status":
                w = ttk.Combobox(left, textvariable=vars_[key], values=CUSTOMER_STATUS_VALUES, width=22, state="readonly")
            else:
                w = ttk.Entry(left, textvariable=vars_[key], width=26)
            w.grid(row=i // 3, column=(i % 3) * 2 + 1, padx=8, pady=6, sticky="ew")

        detail_expanded = {"v": True}
        extra_wrap = ttk.Frame(left)
        extra_wrap.grid(row=4, column=0, columnspan=6, sticky="nsew")
        extra_wrap.columnconfigure(1, weight=1)
        ttk.Label(extra_wrap, text="기타요청").grid(row=0, column=0, padx=8, pady=6, sticky="ne")
        extra_txt = tk.Text(extra_wrap, height=8, wrap="word")
        extra_txt.insert("1.0", str(vars_.get("extra_needs", tk.StringVar(value="")).get()))
        extra_txt.grid(row=0, column=1, padx=8, pady=6, sticky="nsew")
        left.rowconfigure(4, weight=1)

        def save_changes():
            vars_["extra_needs"].set(extra_txt.get("1.0", "end").strip())
            data = {k: v.get() for k, v in vars_.items()}
            before = get_customer(customer_id, include_deleted=True)
            update_customer(customer_id, data)
            if before and str(before.get("status") or "") != str(data.get("status") or ""):
                self._remember_undo({"action": "CUSTOMER_STATUS", "id": customer_id, "before_status": before.get("status")})
            self.refresh_all()
            messagebox.showinfo("완료", "저장되었습니다.")

        def toggle_hide():
            before = get_customer(customer_id, include_deleted=True)
            toggle_customer_hidden(customer_id)
            self._remember_undo({"action": "CUSTOMER_HIDE", "id": customer_id, "before_hidden": int(before.get("hidden") or 0) if before else 0})
            self.request_sync(reason="CUSTOMER_HIDE")
            messagebox.showinfo("완료", "숨김/보임이 변경되었습니다.")
            self.refresh_all()

        def soft_delete():
            if messagebox.askyesno("확인", "삭제 처리(복구 가능) 하시겠습니까?"):
                soft_delete_customer(customer_id)
                self._remember_undo({"action": "CUSTOMER_DELETE", "id": customer_id})
                self.request_sync(reason="CUSTOMER_DELETE")
                self.refresh_all()
                win.destroy()

        ttk.Button(right, text="💾 저장", style="Primary.TButton", command=save_changes).pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(right, text="할 일 추가", style="Secondary.TButton", command=lambda: self.open_add_task_window(default_entity_type="CUSTOMER", default_entity_id=customer_id)).pack(fill="x", padx=8, pady=4)
        ttk.Button(right, text="숨김/보임", style="Secondary.TButton", command=toggle_hide).pack(fill="x", padx=8, pady=4)
        ttk.Button(right, text="🗑️ 삭제", style="Danger.TButton", command=soft_delete).pack(fill="x", padx=8, pady=(4, 8))

        def toggle_extra():
            detail_expanded["v"] = not detail_expanded["v"]
            if detail_expanded["v"]:
                extra_wrap.grid()
                extra_btn.configure(text="기타요청 접기")
            else:
                extra_wrap.grid_remove()
                extra_btn.configure(text="기타요청 펼치기")

        extra_btn = ttk.Button(right, text="기타요청 접기", style="Secondary.TButton", command=toggle_extra)
        extra_btn.pack(fill="x", padx=8, pady=(0, 8))

        self._attach_tooltip(extra_txt, lambda: extra_txt.get("1.0", "end").strip())
        form_widgets = self._collect_form_widgets(left)
        self._wire_form_navigation(form_widgets)
        self._set_initial_focus(win, form_widgets[0] if form_widgets else None)

    def _open_related_navigation_popup(self, *, title: str, customers: list[int] | None = None, properties: list[int] | None = None):
        customers = customers or []
        properties = properties or []
        if not customers and not properties:
            return

        win = tk.Toplevel(self.root)
        win.title(title)
        self._fit_toplevel(win, 620, 460)

        tree = ttk.Treeview(win, columns=("type", "name"), show="headings")
        tree.heading("type", text="유형")
        tree.heading("name", text="대상")
        tree.column("type", width=100, anchor="center")
        tree.column("name", width=380)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for cid in customers:
            c = get_customer(cid)
            tree.insert("", "end", iid=f"C:{cid}", values=("고객", customer_label(c or {})))
        for pid in properties:
            p = get_property(pid)
            tree.insert("", "end", iid=f"P:{pid}", values=("물건", property_label(p or {})))

        def open_selected():
            sel = tree.selection()
            if not sel:
                return
            iid = str(sel[0])
            if iid.startswith("C:"):
                self.open_customer_detail(int(iid.split(":", 1)[1]))
            elif iid.startswith("P:"):
                self.open_property_detail(int(iid.split(":", 1)[1]))

        tree.bind("<Double-1>", lambda _e: open_selected())
        ttk.Button(win, text="열기", command=open_selected).pack(pady=6)
    
    def _pick_one_property_popup(self, parent) -> int | None:
        rows = [p for p in list_properties(include_deleted=False) if not p.get("hidden")]
        if not rows:
            messagebox.showwarning("안내", "선택할 물건이 없습니다.", parent=parent)
            return None

        popup = tk.Toplevel(parent)
        popup.title("물건 선택")
        self._fit_toplevel(popup, 900, 620)

        search_var = tk.StringVar(value="")
        ttk.Entry(popup, textvariable=search_var, width=44).pack(padx=10, pady=6, anchor="w")

        tree = ttk.Treeview(popup, columns=("complex", "dongho", "type", "price"), show="headings", height=18)
        tree.heading("complex", text="단지")
        tree.heading("dongho", text="동/호")
        tree.heading("type", text="타입/평형")
        tree.heading("price", text="가격요약")
        tree.column("complex", width=220)
        tree.column("dongho", width=220)
        tree.column("type", width=170)
        tree.column("price", width=220)
        tree.pack(fill="both", expand=True, padx=10, pady=8)

        def render():
            q = search_var.get().strip().lower()
            qd = "".join(ch for ch in q if ch.isdigit())
            for item in tree.get_children():
                tree.delete(item)
            for r in rows:
                hay = " ".join([
                    str(r.get("complex_name") or ""),
                    str(r.get("dong") or ""),
                    str(r.get("ho") or ""),
                    str(r.get("address_detail") or ""),
                    str(r.get("owner_phone") or ""),
                    str(r.get("special_notes") or ""),
                ]).lower()
                hnum = "".join(ch for ch in str(r.get("owner_phone") or "") if ch.isdigit())
                if q and q not in hay and (not qd or qd not in hnum):
                    continue
                unit = str(r.get("unit_type") or "").strip()
                pyeong = str(r.get("pyeong") or "").strip()
                type_text = f"{unit} {pyeong}평".strip()
                tree.insert(
                    "",
                    "end",
                    iid=str(r.get("id")),
                    values=(
                        r.get("complex_name") or "",
                        str(r.get("address_detail") or "").strip() or property_label(r),
                        type_text,
                        self._calc_price_summary(r),
                    ),
                )

        selected: dict[str, int | None] = {"pid": None}

        def done():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("확인", "물건 1개를 선택하세요", parent=popup)
                return
            try:
                selected["pid"] = int(sel[0])
            except Exception:
                selected["pid"] = None
            popup.destroy()

        tree.bind("<Double-1>", lambda _e: done())
        search_var.trace_add("write", lambda *_: render())
        render()

        btns = ttk.Frame(popup)
        btns.pack(fill="x", padx=10, pady=8)
        ttk.Button(btns, text="선택", command=done).pack(side="right")
        ttk.Button(btns, text="취소", command=popup.destroy).pack(side="right", padx=4)
        popup.grab_set()
        popup.wait_window()
        return selected["pid"]

    def _ranked_photos_for_property(self, property_id: int) -> list[dict]:
        photos = list_photos(property_id)
        order = {tag: idx for idx, tag in enumerate(PHOTO_TAG_VALUES)}
        return sorted(photos, key=lambda p: (order.get(str(p.get("tag") or "기타"), 999), int(p.get("id") or 0)))

    def _pick_customer_popup(self, parent=None) -> dict | None:
        popup = tk.Toplevel(parent or self.root)
        popup.title("고객 선택 (선택사항)")
        self._fit_toplevel(popup, 760, 560)
        rows = [c for c in list_customers(include_deleted=False) if not c.get("hidden")]

        tree = ttk.Treeview(popup, columns=("phone", "name", "status"), show="headings", height=14)
        for col, text, width in (("phone", "전화번호", 160), ("name", "이름", 180), ("status", "상태", 120)):
            tree.heading(col, text=text)
            tree.column(col, width=width)
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        for r in rows:
            tree.insert("", "end", iid=str(r.get("id")), values=(fmt_phone(r.get("phone")), r.get("customer_name") or "", r.get("status") or ""))

        selected = {"customer": None}

        def done():
            sel = tree.selection()
            if sel:
                selected["customer"] = get_customer(int(sel[0]))
            popup.destroy()

        btns = ttk.Frame(popup)
        btns.pack(fill="x", padx=10, pady=8)
        ttk.Button(btns, text="선택", command=done).pack(side="right")
        ttk.Button(btns, text="건너뛰기", command=popup.destroy).pack(side="right", padx=4)
        popup.grab_set()
        popup.wait_window()
        return selected["customer"]

    def generate_proposal_for_property(self, property_id: int, customer: dict | None = None) -> None:
        row = get_property(property_id)
        if not row:
            messagebox.showerror("오류", "물건 정보를 찾을 수 없습니다.")
            return

        customer = customer or self._pick_customer_popup(self.root) or {
            "customer_name": "",
            "phone": "",
            "deal_type": "",
            "preferred_tab": "",
            "budget": "",
        }

        def _work():
            photos = self._ranked_photos_for_property(property_id)
            self.settings = self._load_settings()
            out_dir = self.settings.sync_dir / "exports" / "proposals"
            out = generate_proposal_pdf(
                customer=customer,
                properties=[row],
                photos_by_property={property_id: photos},
                output_dir=out_dir,
                title="매물 제안서",
            )

            ymd = datetime.now().strftime("%Y%m%d")
            complex_part = self._safe_file_component(row.get("complex_name") or "매물")
            dongho_text = str(row.get("address_detail") or "").strip() or f"{str(row.get('dong') or '').strip()}_{str(row.get('ho') or '').strip()}"
            dongho_part = self._safe_file_component(dongho_text)
            package_dir = self.settings.sync_dir / "exports" / "proposal_packages" / f"{complex_part}_{dongho_part}_{ymd}"
            package_dir.mkdir(parents=True, exist_ok=True)

            if out.pdf_path.exists():
                shutil.copy2(out.pdf_path, package_dir / self._safe_file_component(out.pdf_path.name))

            for ph in photos:
                src = Path(str(ph.get("file_path") or "").strip())
                if not src.exists() or not src.is_file():
                    continue
                tag = str(ph.get("tag") or "기타").strip() or "기타"
                tag_dir = package_dir / self._safe_file_component(tag)
                tag_dir.mkdir(parents=True, exist_ok=True)
                dst = tag_dir / src.name
                if not dst.exists():
                    shutil.copy2(src, dst)
            return package_dir

        try:
            package_dir = self.run_with_busy_ui(_work, busy_message="PDF/사진 패킹 처리 중…", success_message="PDF/사진 패킹 완료")
        except Exception as exc:
            messagebox.showerror("오류", f"제안서 생성 실패: {exc}")
            return

        messagebox.showinfo("완료", f"PDF/사진 패킹 완료\n{package_dir.name}")
        _open_folder(package_dir)

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


def run_desktop_app():
    root = tk.Tk()
    LedgerDesktopApp(root)
    root.mainloop()
