# -*- coding: utf-8 -*-
"""
账单转换可视化界面（tkinter，无需额外安装）

启动:
  python bill_convert_ui.py
  或双击 启动转换工具.bat
"""
from __future__ import annotations

import importlib
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from convert_profiles import PROFILES, get_profile, list_customers, list_suppliers
from fx_rate import fetch_usd_rates, get_china_pn_fx_rate, get_hk_pn_fx_rate, get_tw_pn_fx_rate
from pn_meta import PnMeta, build_invoice_number, parse_date

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "输出"
OUTPUT_DIR.mkdir(exist_ok=True)


class BillConvertApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("账单转换工具")
        self.geometry("760x860")
        self.minsize(680, 760)
        self.configure(padx=12, pady=12)

        self.file_path = tk.StringVar()
        self.supplier_var = tk.StringVar()
        self.customer_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择供应商、客户和原始账单文件")
        self.fx_china_var = tk.StringVar(value="—")
        self.fx_hk_var = tk.StringVar(value="—")
        self.last_output: Path | None = None

        self.fx_tw_var = tk.StringVar(value="—")

        self.pn_customer_name_var = tk.StringVar()
        self.pn_customer_id_var = tk.StringVar()
        self.pn_billing_address_var = tk.StringVar()
        self.pn_invoice_number_var = tk.StringVar()
        self.pn_invoice_date_var = tk.StringVar(value=date.today().isoformat())
        self.pn_due_date_var = tk.StringVar()

        self._build_ui()
        self._init_suppliers()
        self._refresh_fx_rates()

    def _build_ui(self) -> None:
        title = ttk.Label(self, text="账单转换工具", font=("Microsoft YaHei UI", 16, "bold"))
        title.pack(anchor="w")
        ttk.Label(
            self,
            text="选择供应商与客户后，系统自动匹配 China / Hong Kong / Taiwan 转换流程与汇率",
            foreground="#555",
        ).pack(anchor="w", pady=(4, 12))

        # ---------- 文件选择（顶部）----------
        file_frame = ttk.LabelFrame(self, text="1. 选择转换文件", padding=10)
        file_frame.pack(fill="x", pady=(0, 10))

        row = ttk.Frame(file_frame)
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.file_path).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="浏览…", command=self._browse_file).pack(side="right")

        # ---------- 供应商 / 客户 ----------
        sel_frame = ttk.LabelFrame(self, text="2. 选择供应商与客户", padding=10)
        sel_frame.pack(fill="x", pady=(0, 10))

        grid = ttk.Frame(sel_frame)
        grid.pack(fill="x")
        ttk.Label(grid, text="供应商").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.supplier_cb = ttk.Combobox(grid, textvariable=self.supplier_var, state="readonly", width=36)
        self.supplier_cb.grid(row=0, column=1, sticky="ew", pady=4)
        self.supplier_cb.bind("<<ComboboxSelected>>", self._on_supplier_change)

        ttk.Label(grid, text="客户").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.customer_cb = ttk.Combobox(grid, textvariable=self.customer_var, state="readonly", width=36)
        self.customer_cb.grid(row=1, column=1, sticky="ew", pady=4)
        self.customer_cb.bind("<<ComboboxSelected>>", self._on_customer_change)
        grid.columnconfigure(1, weight=1)

        # ---------- PN 外部信息 ----------
        pn_frame = ttk.LabelFrame(self, text="3. PN 账单信息（外部提供，写入母版 PN 页）", padding=10)
        pn_frame.pack(fill="x", pady=(0, 10))

        pn_grid = ttk.Frame(pn_frame)
        pn_grid.pack(fill="x")
        pn_fields = [
            ("客户名称 (B8)", self.pn_customer_name_var),
            ("客户 ID (B9)", self.pn_customer_id_var),
            ("账单地址 (B10)", self.pn_billing_address_var),
        ]
        for row, (label, var) in enumerate(pn_fields):
            ttk.Label(pn_grid, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
            entry = ttk.Entry(pn_grid, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", pady=3)
            if row == 0:
                entry.bind("<KeyRelease>", lambda _e: self._refresh_invoice_preview())
            if row == 1:
                entry.bind("<KeyRelease>", lambda _e: self._refresh_invoice_preview())

        ttk.Label(pn_grid, text="账单编号 (F9)").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=3)
        inv_row = ttk.Frame(pn_grid)
        inv_row.grid(row=3, column=1, sticky="ew", pady=3)
        self.pn_invoice_entry = ttk.Entry(inv_row, textvariable=self.pn_invoice_number_var, state="readonly")
        self.pn_invoice_entry.pack(side="left", fill="x", expand=True)
        ttk.Label(inv_row, text="自动生成", foreground="#666").pack(side="left", padx=(8, 0))

        ttk.Label(pn_grid, text="发账单日期 (F10)").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=3)
        inv_date_entry = ttk.Entry(pn_grid, textvariable=self.pn_invoice_date_var, width=16)
        inv_date_entry.grid(row=4, column=1, sticky="w", pady=3)
        inv_date_entry.bind("<KeyRelease>", lambda _e: self._refresh_invoice_preview())
        ttk.Label(pn_grid, text="格式 YYYY-MM-DD，默认今天", foreground="#666").grid(
            row=4, column=2, sticky="w", padx=(8, 0), pady=3
        )

        ttk.Label(pn_grid, text="Due date (F11)").grid(row=5, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(pn_grid, textvariable=self.pn_due_date_var, width=16).grid(row=5, column=1, sticky="w", pady=3)
        ttk.Label(pn_grid, text="必填，格式 YYYY-MM-DD", foreground="#666").grid(
            row=5, column=2, sticky="w", padx=(8, 0), pady=3
        )

        ttk.Label(
            pn_frame,
            text="账单编号规则: PN-{客户ID}-{MMDDYYYY}{当天第几单}，例 PN-CUS15253-031820261",
            foreground="#555",
            wraplength=700,
        ).pack(anchor="w", pady=(8, 0))
        pn_grid.columnconfigure(1, weight=1)

        # ---------- 匹配信息 ----------
        info_frame = ttk.LabelFrame(self, text="4. 当前匹配规则", padding=10)
        info_frame.pack(fill="x", pady=(0, 10))

        self.info_region = ttk.Label(info_frame, text="地区: —")
        self.info_region.pack(anchor="w")
        self.info_engine = ttk.Label(info_frame, text="转换引擎: —")
        self.info_engine.pack(anchor="w")
        self.info_fx = ttk.Label(info_frame, text="汇率: —")
        self.info_fx.pack(anchor="w")
        self.info_template = ttk.Label(info_frame, text="母版: —")
        self.info_template.pack(anchor="w")

        # ---------- 实时汇率 ----------
        fx_frame = ttk.LabelFrame(self, text="实时汇率（exchangerate-api）", padding=10)
        fx_frame.pack(fill="x", pady=(0, 10))
        fx_row = ttk.Frame(fx_frame)
        fx_row.pack(fill="x")
        ttk.Label(fx_row, text="China PN!B29 (CNY):").pack(side="left")
        ttk.Label(fx_row, textvariable=self.fx_china_var, font=("Consolas", 10, "bold")).pack(side="left", padx=8)
        ttk.Label(fx_row, text="HK PN!B28 (HKD×0.97):").pack(side="left", padx=(20, 0))
        ttk.Label(fx_row, textvariable=self.fx_hk_var, font=("Consolas", 10, "bold")).pack(side="left", padx=8)
        ttk.Label(fx_row, text="TW PN!B31 (TWD):").pack(side="left", padx=(20, 0))
        ttk.Label(fx_row, textvariable=self.fx_tw_var, font=("Consolas", 10, "bold")).pack(side="left", padx=8)
        ttk.Button(fx_frame, text="刷新汇率", command=self._refresh_fx_rates).pack(anchor="e", pady=(6, 0))

        # ---------- 一键转换 ----------
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", pady=(0, 10))
        self.convert_btn = ttk.Button(
            btn_frame,
            text="一键转换",
            command=self._on_convert_click,
            style="Accent.TButton",
        )
        self.convert_btn.pack(fill="x", ipady=8)

        try:
            style = ttk.Style(self)
            style.configure("Accent.TButton", font=("Microsoft YaHei UI", 11, "bold"))
        except tk.TclError:
            pass

        aux = ttk.Frame(self)
        aux.pack(fill="x", pady=(0, 8))
        ttk.Button(aux, text="打开输出文件夹", command=self._open_output_dir).pack(side="left")
        ttk.Button(aux, text="打开最近生成的文件", command=self._open_last_output).pack(side="left", padx=8)

        # ---------- 日志 ----------
        log_frame = ttk.LabelFrame(self, text="转换日志", padding=8)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, height=10, wrap="word", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, side="left")
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        scroll.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=scroll.set, state="disabled")

        ttk.Label(self, textvariable=self.status_var, foreground="#0066cc").pack(anchor="w", pady=(6, 0))

    def _log(self, msg: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _init_suppliers(self) -> None:
        suppliers = list_suppliers()
        self.supplier_cb["values"] = suppliers
        if suppliers:
            self.supplier_var.set(suppliers[0])
            self._on_supplier_change()

    def _on_supplier_change(self, _event=None) -> None:
        supplier = self.supplier_var.get()
        customers = list_customers(supplier)
        self.customer_cb["values"] = customers
        if customers:
            self.customer_var.set(customers[0])
        else:
            self.customer_var.set("")
        self._fill_pn_defaults_from_profile()
        self._update_profile_info()

    def _on_customer_change(self, _event=None) -> None:
        self._fill_pn_defaults_from_profile()
        self._update_profile_info()

    def _fill_pn_defaults_from_profile(self) -> None:
        profile = get_profile(self.supplier_var.get(), self.customer_var.get())
        if not profile:
            return
        self.pn_customer_name_var.set(profile.pn_customer_name)
        self.pn_customer_id_var.set(profile.pn_customer_id)
        self.pn_billing_address_var.set(profile.pn_billing_address)
        if not self.pn_invoice_date_var.get().strip():
            self.pn_invoice_date_var.set(date.today().isoformat())
        self._refresh_invoice_preview()

    def _refresh_invoice_preview(self) -> None:
        customer_id = self.pn_customer_id_var.get().strip()
        if not customer_id:
            self.pn_invoice_number_var.set("")
            return
        try:
            invoice_date = parse_date(self.pn_invoice_date_var.get().strip()) or date.today()
            preview = build_invoice_number(
                customer_id,
                invoice_date,
                registry_dir=OUTPUT_DIR,
                reserve=False,
            )
            self.pn_invoice_number_var.set(preview)
        except ValueError:
            self.pn_invoice_number_var.set("（日期或客户 ID 无效）")

    def _collect_pn_meta(self) -> PnMeta:
        try:
            invoice_date = parse_date(self.pn_invoice_date_var.get().strip()) or date.today()
            due_date = parse_date(self.pn_due_date_var.get().strip())
        except ValueError as exc:
            raise ValueError(str(exc)) from exc
        if due_date is None:
            raise ValueError("请填写 Due date (PN!F11)，格式 YYYY-MM-DD")
        customer_name = self.pn_customer_name_var.get().strip()
        customer_id = self.pn_customer_id_var.get().strip()
        billing_address = self.pn_billing_address_var.get().strip()
        if not customer_name:
            raise ValueError("请填写客户名称 (PN!B8)")
        if not customer_id:
            raise ValueError("请填写客户 ID (PN!B9)")
        if not billing_address:
            raise ValueError("请填写账单地址 (PN!B10)")
        return PnMeta(
            customer_name=customer_name,
            customer_id=customer_id,
            billing_address=billing_address,
            invoice_date=invoice_date,
            due_date=due_date,
        )

    def _update_profile_info(self) -> None:
        profile = get_profile(self.supplier_var.get(), self.customer_var.get())
        if not profile:
            self.info_region.config(text="地区: 未匹配")
            self.info_engine.config(text="转换引擎: —")
            self.info_fx.config(text="汇率: —")
            self.info_template.config(text="母版: —")
            return
        engine = profile.convert_engine
        self.info_region.config(text=f"地区: {profile.region}")
        self.info_engine.config(text=f"转换引擎: {engine.label}（{engine.description}）")
        self.info_fx.config(text=f"汇率: {profile.fx_cell}  ({profile.fx_description})")
        rel = profile.template.relative_to(BASE_DIR) if profile.template.is_relative_to(BASE_DIR) else profile.template
        self.info_template.config(text=f"母版: {rel}")

    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择原始账单",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir=str(BASE_DIR / "账单"),
        )
        if path:
            self.file_path.set(path)
            self.status_var.set(f"已选择: {Path(path).name}")

    def _refresh_fx_rates(self) -> None:
        def task() -> None:
            try:
                rates = fetch_usd_rates()
                cny = get_china_pn_fx_rate(rates)
                hk = get_hk_pn_fx_rate(rates)
                tw = get_tw_pn_fx_rate(rates)
                self.after(0, lambda: self.fx_china_var.set(f"{cny:.4f}"))
                self.after(0, lambda: self.fx_hk_var.set(f"{hk:.4f}"))
                self.after(0, lambda: self.fx_tw_var.set(f"{tw:.4f}"))
            except Exception as exc:
                self.after(0, lambda: self.status_var.set(f"汇率获取失败: {exc}"))

        threading.Thread(target=task, daemon=True).start()

    def _open_output_dir(self) -> None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        import os
        os.startfile(OUTPUT_DIR)

    def _open_last_output(self) -> None:
        if self.last_output and self.last_output.is_file():
            import os
            os.startfile(self.last_output)
        else:
            messagebox.showinfo("提示", "还没有生成过文件，请先一键转换。")

    def _on_convert_click(self) -> None:
        profile = get_profile(self.supplier_var.get(), self.customer_var.get())
        if not profile:
            messagebox.showerror("错误", "请选择有效的供应商与客户组合")
            return

        src = self.file_path.get().strip()
        if not src:
            messagebox.showerror("错误", "请先选择要转换的原始账单文件")
            return
        source_path = Path(src)
        if not source_path.is_file():
            messagebox.showerror("错误", f"文件不存在:\n{source_path}")
            return
        if not profile.template.is_file():
            messagebox.showerror("错误", f"母版不存在:\n{profile.template}")
            return

        try:
            pn_meta = self._collect_pn_meta()
        except ValueError as exc:
            messagebox.showerror("PN 信息不完整", str(exc))
            return

        self.convert_btn.config(state="disabled")
        self.status_var.set("正在转换…")
        threading.Thread(
            target=self._run_convert,
            args=(profile, source_path, pn_meta),
            daemon=True,
        ).start()

    def _run_convert(self, profile, source_path: Path, pn_meta: PnMeta) -> None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"{profile.output_prefix}_N-C_{ts}.xlsx"
        output_path = OUTPUT_DIR / out_name

        try:
            mod = importlib.import_module(profile.module)
            result = mod.convert(
                source_path,
                output_path,
                profile.template,
                pn_meta=pn_meta,
                registry_dir=OUTPUT_DIR,
            )

            self.last_output = output_path
            pn_result = result.get("pn_meta") or {}
            lines = [
                f"✓ 转换成功: {out_name}",
                f"  账单编号: {pn_result.get('invoice_number', '—')}",
                f"  员工: {result.get('employee_count')} 人",
                f"  {profile.fx_cell}: {result.get('fx_rate')}",
            ]
            names = result.get("employee_names") or []
            if names:
                lines.append(f"  姓名: {', '.join(str(n) for n in names if n)}")
            if profile.engine == "china_hrone":
                lines.append(f"  Other: {result.get('other_amount')}  报销笔数: {result.get('expense_count')}")
            else:
                lines.append(f"  公司: {result.get('company_name')}  账期: {result.get('period')}")

            msg = "\n".join(lines)

            def done_ok() -> None:
                self._log(msg)
                self.status_var.set(f"完成 → {output_path}")
                self.convert_btn.config(state="normal")
                self._refresh_invoice_preview()
                self._refresh_fx_rates()
                if messagebox.askyesno("转换完成", f"{msg}\n\n是否打开输出文件？"):
                    import os
                    os.startfile(output_path)

            self.after(0, done_ok)

        except Exception as exc:
            def done_err() -> None:
                self._log(f"✗ 失败: {exc}")
                self.status_var.set("转换失败")
                self.convert_btn.config(state="normal")
                messagebox.showerror("转换失败", str(exc))

            self.after(0, done_err)


def main() -> None:
    try:
        app = BillConvertApp()
        # 窗口置顶并置前，避免被 IDE 终端挡住
        app.update_idletasks()
        app.lift()
        app.attributes("-topmost", True)
        app.after(300, lambda: app.attributes("-topmost", False))
        app.focus_force()
        print("账单转换工具窗口已打开（请查看任务栏）", flush=True)
        app.mainloop()
    except Exception as exc:
        import traceback
        traceback.print_exc()
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("启动失败", str(exc))
            root.destroy()
        except Exception:
            pass
        input("按回车键退出…")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
