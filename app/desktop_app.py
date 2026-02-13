import tkinter as tk
from tkinter import messagebox, ttk

from .sheet_sync import upload_visible_data
from .storage import (
    PROPERTY_TABS,
    add_customer,
    add_property,
    delete_customer,
    delete_property,
    init_db,
    list_customers,
    list_properties,
    toggle_customer_hidden,
    toggle_property_hidden,
)


class LedgerDesktopApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("엄마 부동산 장부 (데스크탑)")
        self.root.geometry("1400x900")

        init_db()

        self.main = ttk.Notebook(root)
        self.main.pack(fill="both", expand=True)

        self.property_tab = ttk.Frame(self.main)
        self.customer_tab = ttk.Frame(self.main)
        self.main.add(self.property_tab, text="물건 장부")
        self.main.add(self.customer_tab, text="고객 요구사항")

        self._build_property_ui()
        self._build_customer_ui()
        self.refresh_all()

    def _build_property_ui(self):
        top = ttk.LabelFrame(self.property_tab, text="물건 등록")
        top.pack(fill="x", padx=10, pady=8)

        self.pvars = {
            "tab": tk.StringVar(value=PROPERTY_TABS[0]),
            "complex_name": tk.StringVar(),
            "unit_type": tk.StringVar(),
            "area": tk.StringVar(),
            "pyeong": tk.StringVar(),
            "floor": tk.StringVar(),
            "condition": tk.StringVar(value="중"),
            "repair_needed": tk.BooleanVar(value=False),
            "tenant_info": tk.StringVar(),
            "naver_link": tk.StringVar(),
            "special_notes": tk.StringVar(),
            "note": tk.StringVar(),
        }

        fields = [
            ("탭", "tab", "combo", PROPERTY_TABS),
            ("단지명", "complex_name", "entry", None),
            ("면적타입", "unit_type", "entry", None),
            ("면적(㎡)", "area", "entry", None),
            ("평형", "pyeong", "entry", None),
            ("층수", "floor", "entry", None),
            ("컨디션", "condition", "combo", ["상", "중", "하"]),
            ("세입자정보", "tenant_info", "entry", None),
            ("네이버링크", "naver_link", "entry", None),
            ("특이사항", "special_notes", "entry", None),
            ("별도기재", "note", "entry", None),
        ]

        for i, (label, key, kind, options) in enumerate(fields):
            ttk.Label(top, text=label).grid(row=i // 4, column=(i % 4) * 2, padx=6, pady=4, sticky="e")
            if kind == "combo":
                w = ttk.Combobox(top, textvariable=self.pvars[key], values=options, width=20, state="readonly")
            else:
                w = ttk.Entry(top, textvariable=self.pvars[key], width=24)
            w.grid(row=i // 4, column=(i % 4) * 2 + 1, padx=6, pady=4, sticky="w")

        ttk.Checkbutton(top, text="수리필요", variable=self.pvars["repair_needed"]).grid(row=3, column=0, padx=6, pady=6, sticky="w")
        ttk.Button(top, text="물건 등록", command=self.create_property).grid(row=3, column=1, padx=6, pady=6, sticky="w")
        ttk.Button(top, text="숨김 제외 데이터 구글시트 업로드", command=self.upload_sheet).grid(row=3, column=2, columnspan=2, padx=6, pady=6, sticky="w")

        self.inner_tabs = ttk.Notebook(self.property_tab)
        self.inner_tabs.pack(fill="both", expand=True, padx=10, pady=8)
        self.prop_trees = {}

        cols = ("id", "hidden", "complex_name", "unit_type", "floor", "condition", "special_notes")
        for tab_name in PROPERTY_TABS:
            frame = ttk.Frame(self.inner_tabs)
            self.inner_tabs.add(frame, text=tab_name)
            tree = ttk.Treeview(frame, columns=cols, show="headings", height=15)
            for c, w in [("id", 55), ("hidden", 70), ("complex_name", 220), ("unit_type", 100), ("floor", 80), ("condition", 80), ("special_notes", 450)]:
                tree.heading(c, text=c)
                tree.column(c, width=w)
            tree.pack(fill="both", expand=True)
            self.prop_trees[tab_name] = tree

            btns = ttk.Frame(frame)
            btns.pack(fill="x", pady=4)
            ttk.Button(btns, text="숨김/보임 전환", command=lambda t=tab_name: self.toggle_selected_property(t)).pack(side="left", padx=4)
            ttk.Button(btns, text="삭제", command=lambda t=tab_name: self.delete_selected_property(t)).pack(side="left", padx=4)

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
        ]

        for i, (label, key, kind, options) in enumerate(fields):
            ttk.Label(top, text=label).grid(row=i // 4, column=(i % 4) * 2, padx=6, pady=4, sticky="e")
            if kind == "combo":
                w = ttk.Combobox(top, textvariable=self.cvars[key], values=options, width=20, state="readonly")
            else:
                w = ttk.Entry(top, textvariable=self.cvars[key], width=24)
            w.grid(row=i // 4, column=(i % 4) * 2 + 1, padx=6, pady=4, sticky="w")

        ttk.Button(top, text="고객 등록", command=self.create_customer).grid(row=3, column=0, padx=6, pady=6, sticky="w")

        cols = ("id", "hidden", "customer_name", "phone", "preferred_tab", "budget", "floor_preference", "extra_needs")
        self.customer_tree = ttk.Treeview(self.customer_tab, columns=cols, show="headings", height=20)
        for c, w in [("id", 55), ("hidden", 70), ("customer_name", 120), ("phone", 120), ("preferred_tab", 110), ("budget", 120), ("floor_preference", 120), ("extra_needs", 500)]:
            self.customer_tree.heading(c, text=c)
            self.customer_tree.column(c, width=w)
        self.customer_tree.pack(fill="both", expand=True, padx=10, pady=8)

        btns = ttk.Frame(self.customer_tab)
        btns.pack(fill="x", padx=10, pady=4)
        ttk.Button(btns, text="숨김/보임 전환", command=self.toggle_selected_customer).pack(side="left", padx=4)
        ttk.Button(btns, text="삭제", command=self.delete_selected_customer).pack(side="left", padx=4)

    def create_property(self):
        add_property({k: v.get() if not isinstance(v, tk.BooleanVar) else v.get() for k, v in self.pvars.items()})
        self.refresh_all()

    def create_customer(self):
        if not self.cvars["customer_name"].get().strip():
            messagebox.showwarning("확인", "고객명은 필수입니다.")
            return
        add_customer({k: v.get() for k, v in self.cvars.items()})
        self.refresh_all()

    def refresh_all(self):
        for tab in PROPERTY_TABS:
            tree = self.prop_trees[tab]
            for i in tree.get_children():
                tree.delete(i)
            for row in list_properties(tab):
                tree.insert("", "end", values=(row["id"], "숨김" if row["hidden"] else "보임", row["complex_name"], row["unit_type"], row["floor"], row["condition"], row["special_notes"]))

        for i in self.customer_tree.get_children():
            self.customer_tree.delete(i)
        for row in list_customers():
            self.customer_tree.insert("", "end", values=(row["id"], "숨김" if row["hidden"] else "보임", row["customer_name"], row["phone"], row["preferred_tab"], row["budget"], row["floor_preference"], row["extra_needs"]))

    def toggle_selected_property(self, tab_name: str):
        tree = self.prop_trees[tab_name]
        selected = tree.selection()
        if not selected:
            return
        pid = int(tree.item(selected[0], "values")[0])
        toggle_property_hidden(pid)
        self.refresh_all()

    def delete_selected_property(self, tab_name: str):
        tree = self.prop_trees[tab_name]
        selected = tree.selection()
        if not selected:
            return
        pid = int(tree.item(selected[0], "values")[0])
        if messagebox.askyesno("확인", "해당 물건을 삭제할까요?"):
            delete_property(pid)
            self.refresh_all()

    def toggle_selected_customer(self):
        selected = self.customer_tree.selection()
        if not selected:
            return
        cid = int(self.customer_tree.item(selected[0], "values")[0])
        toggle_customer_hidden(cid)
        self.refresh_all()

    def delete_selected_customer(self):
        selected = self.customer_tree.selection()
        if not selected:
            return
        cid = int(self.customer_tree.item(selected[0], "values")[0])
        if messagebox.askyesno("확인", "해당 고객 요청을 삭제할까요?"):
            delete_customer(cid)
            self.refresh_all()

    def upload_sheet(self):
        ok, msg = upload_visible_data(list_properties(), list_customers())
        if ok:
            messagebox.showinfo("업로드", msg)
        else:
            messagebox.showerror("업로드 실패", msg)


def run_desktop_app():
    root = tk.Tk()
    LedgerDesktopApp(root)
    root.mainloop()
