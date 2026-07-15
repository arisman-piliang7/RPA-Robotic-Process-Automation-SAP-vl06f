"""
SAP VL06F Automation Bot v2.3
- Plant selector pakai Listbox (bukan chip/widget per item) → instant untuk 109+ plant
- Scheduler fix: wait_visibility() sebelum grab_set() → tidak blank lagi
- SAP popup scroll tiap 8 baris
- Auto-save & buka Excel setelah export
- Windows Task Scheduler integration
Python 3 + tkinter (stdlib only)

FIXES v2.3:
- PlantListbox: native tk.Listbox + Scrollbar, multi-select, tanpa overhead widget per item
- SchedulerDialog: grab_set() dipindah setelah wait_visibility() → tidak blank/freeze
- _gen_plant_lines: scroll SAP popup tiap SAP_POPUP_PAGE (8) plant
- Threading worker: set_running(False) selalu di finally
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess, tempfile, os, calendar, re, threading
from datetime import datetime, timedelta
from pathlib import Path

# ── colours ──────────────────────────────────────────────────────────
BG = "#0F172A"
SURF = "#1E293B"
SURF2 = "#243044"
ACCENT = "#38BDF8"
INDIGO = "#818CF8"
GREEN = "#34D399"
AMBER = "#FBBF24"
RED = "#F87171"
TEXT = "#F1F5F9"
MUTED = "#94A3B8"
BORDER = "#2D3F55"

FN = ("Segoe UI", 9)
FNB = ("Segoe UI", 9, "bold")
FNS = ("Segoe UI", 8)
FH = ("Segoe UI", 11, "bold")
FM = ("Consolas", 9)

DEFAULT_PLANTS = ["2150", "2151", "2152", "2153", "2154", "2155", "2162", "2163"]
DEFAULT_LAYOUTS = ["/ARIPILI_OS3", "/ARIPILI_OS4", "/ARIPILI_OS2"]
DEFAULT_EXPORT_DIR = str(Path.home() / "Documents" / "SAP_Export")

MAX_PLANTS = 200  # raised — Listbox handles it fine
SAP_POPUP_PAGE = 8  # baris visible SAP popup sebelum scroll


# ── small helpers ─────────────────────────────────────────────────────
def mkbtn(parent, text, cmd, bg=SURF2, fg=ACCENT, px=10, py=5, **kw):
    return tk.Button(
        parent,
        text=text,
        command=cmd,
        bg=bg,
        fg=fg,
        font=FNB,
        relief="flat",
        bd=0,
        cursor="hand2",
        activebackground=ACCENT,
        activeforeground=BG,
        padx=px,
        pady=py,
        **kw,
    )


def mkentry(parent, var=None, w=10, mono=False):
    return tk.Entry(
        parent,
        textvariable=var,
        bg=SURF2,
        fg=TEXT,
        insertbackground=TEXT,
        font=FM if mono else FN,
        relief="flat",
        bd=0,
        width=w,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
    )


def divider(parent):
    tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(6, 5))


def sec_label(parent, icon, title):
    tk.Label(parent, text=f"{icon}  {title}", bg=SURF, fg=TEXT, font=FNB).pack(
        anchor="w"
    )
    divider(parent)


# ── Calendar popup ────────────────────────────────────────────────────
class CalPicker(tk.Toplevel):
    MO = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "Mei",
        "Jun",
        "Jul",
        "Agu",
        "Sep",
        "Okt",
        "Nov",
        "Des",
    ]

    def __init__(self, parent, cb, init=None):
        super().__init__(parent)
        self.cb = cb
        self.configure(bg=SURF, padx=10, pady=10)
        self.title("Pilih Tanggal")
        self.resizable(False, False)
        self.wait_visibility()
        self.grab_set()
        d = init or datetime.today()
        self.yr, self.mo, self.day = d.year, d.month, d.day
        self._ui()
        self._draw()

    def _ui(self):
        nav = tk.Frame(self, bg=SURF)
        nav.pack(fill="x", pady=(0, 5))

        def nb(t, c):
            return tk.Button(
                nav,
                text=t,
                command=c,
                bg=SURF2,
                fg=ACCENT,
                font=FNB,
                relief="flat",
                bd=0,
                padx=8,
                pady=2,
                cursor="hand2",
            )

        nb("‹", self._prev).pack(side="left")
        self.ml = tk.Label(nav, text="", bg=SURF, fg=TEXT, font=FNB, width=15)
        self.ml.pack(side="left", expand=True)
        nb("›", self._next).pack(side="right")
        hdr = tk.Frame(self, bg=SURF)
        hdr.pack()
        for d in ["S", "S", "R", "K", "J", "S", "M"]:
            tk.Label(hdr, text=d, bg=SURF, fg=MUTED, font=FNS, width=3).pack(
                side="left"
            )
        self.gf = tk.Frame(self, bg=SURF)
        self.gf.pack(pady=2)
        mkbtn(self, "✓  Pilih", self._ok, bg=ACCENT, fg=BG).pack(fill="x", pady=(5, 0))

    def _draw(self):
        for w in self.gf.winfo_children():
            w.destroy()
        self.ml.config(text=f"{self.MO[self.mo-1]} {self.yr}")
        for wk in calendar.monthcalendar(self.yr, self.mo):
            row = tk.Frame(self.gf, bg=SURF)
            row.pack()
            for d in wk:
                if d == 0:
                    tk.Label(row, text="", bg=SURF, width=3).pack(side="left")
                else:
                    s = d == self.day
                    tk.Button(
                        row,
                        text=str(d),
                        bg=ACCENT if s else SURF2,
                        fg=BG if s else TEXT,
                        font=FNS,
                        relief="flat",
                        bd=0,
                        width=2,
                        pady=3,
                        cursor="hand2",
                        command=lambda x=d: self._pick(x),
                    ).pack(side="left", padx=1, pady=1)

    def _pick(self, d):
        self.day = d
        self._draw()

    def _prev(self):
        self.mo -= 1
        if self.mo == 0:
            self.mo = 12
            self.yr -= 1
        self._draw()

    def _next(self):
        self.mo += 1
        if self.mo == 13:
            self.mo = 1
            self.yr += 1
        self._draw()

    def _ok(self):
        try:
            self.cb(datetime(self.yr, self.mo, self.day))
            self.destroy()
        except ValueError:
            pass


# ── DateEntry ─────────────────────────────────────────────────────────
class DateEntry(tk.Frame):
    def __init__(self, parent, init=None, **kw):
        super().__init__(parent, bg=kw.pop("bg", SURF))
        self.var = tk.StringVar(value=(init or datetime.today()).strftime("%d.%m.%Y"))
        mkentry(self, var=self.var, w=10).pack(side="left", ipady=4, padx=(0, 3))
        tk.Button(
            self,
            text="📅",
            bg=SURF2,
            fg=ACCENT,
            font=("Segoe UI", 10),
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=4,
            command=self._open,
        ).pack(side="left")

    def _open(self):
        try:
            d = datetime.strptime(self.var.get(), "%d.%m.%Y")
        except Exception:
            d = datetime.today()
        CalPicker(self, lambda dt: self.var.set(dt.strftime("%d.%m.%Y")), d)

    def get(self):
        return self.var.get()


# ── PlantListbox ──────────────────────────────────────────────────────
class PlantListbox(tk.Frame):
    """
    Pengganti chip area. Pakai native tk.Listbox (MULTIPLE select).
    Tidak ada widget per item — langsung insert text → instant untuk 100+ plant.
    Item selected (highlight biru) = plant yang AKAN diproses.
    """

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=kw.pop("bg", SURF), **kw)
        self._plants: list[str] = []  # ordered, no duplicates
        self._build()

    def _build(self):
        # listbox + scrollbar
        lf = tk.Frame(self, bg=SURF)
        lf.pack(fill="both", expand=True)
        self.lb = tk.Listbox(
            lf,
            selectmode="extended",
            height=7,
            bg=SURF2,
            fg=TEXT,
            selectbackground=ACCENT,
            selectforeground=BG,
            font=FM,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            activestyle="none",
            exportselection=False,
        )
        vsb = tk.Scrollbar(lf, orient="vertical", command=self.lb.yview)
        self.lb.config(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.lb.pack(side="left", fill="both", expand=True)

        # counter label
        self.count_var = tk.StringVar(value="")
        tk.Label(
            self, textvariable=self.count_var, bg=SURF, fg=ACCENT, font=FNS, anchor="w"
        ).pack(fill="x", pady=(2, 0))

        self.lb.bind("<<ListboxSelect>>", lambda e: self._update_count())

    def _update_count(self):
        total = len(self._plants)
        checked = len(self.lb.curselection())
        if total == 0:
            self.count_var.set("")
        else:
            self.count_var.set(
                f"✔  {checked} dari {total} plant dipilih  —  klik+Ctrl/Shift untuk multi-select"
            )

    def set_plants(self, plants: list[str]):
        """Replace seluruh daftar, select semua."""
        self._plants = list(
            dict.fromkeys(p.upper().strip() for p in plants if p.strip())
        )
        self._reload(select_all=True)

    def add_plants(self, codes: list[str]):
        """Tambahkan kode baru, abaikan duplikat."""
        existing = set(self._plants)
        added = False
        for c in codes:
            c = c.upper().strip()
            if c and c not in existing:
                self._plants.append(c)
                existing.add(c)
                added = True
        if added:
            self._reload(select_all=False)

    def _reload(self, select_all=True):
        self.lb.delete(0, "end")
        for p in self._plants:
            self.lb.insert("end", p)
        if select_all:
            self.lb.select_set(0, "end")
        self._update_count()

    def select_all(self):
        self.lb.select_set(0, "end")
        self._update_count()

    def deselect_all(self):
        self.lb.select_clear(0, "end")
        self._update_count()

    def clear_all(self):
        self._plants.clear()
        self.lb.delete(0, "end")
        self._update_count()

    def remove_selected(self):
        sel = sorted(self.lb.curselection(), reverse=True)
        for i in sel:
            self._plants.pop(i)
            self.lb.delete(i)
        self._update_count()

    def get_selected(self) -> list[str]:
        return [self._plants[i] for i in self.lb.curselection()]

    def get_all(self) -> list[str]:
        return list(self._plants)


# ── VBS generator ─────────────────────────────────────────────────────
def _gen_plant_lines(plants):
    PAGE = SAP_POPUP_PAGE
    BASE = (
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
        "ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
    )
    lines = []
    for i, code in enumerate(plants):
        if i > 0 and i % PAGE == 0:
            lines.append(
                f'session.findById("{BASE}").verticalScrollbar.position = {i}\n'
            )
        row_in_page = i % PAGE
        lines.append(
            f'session.findById("{BASE}/ctxtRSCSEL_255-SLOW_I[1,{row_in_page}]")'
            f'.text = "{code}"\n'
        )
    return "".join(lines)


def gen_vbs(
    plants, st_lo, st_hi, dt_lo, dt_hi, layout, ck_item, ck_anzpo, ck_spd_a, export_dir
):
    pl = _gen_plant_lines(plants)
    b = lambda x: "true" if x else "false"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"VL06F_{ts}.xlsx"
    export_path = export_dir.replace("/", "\\").rstrip("\\") + "\\" + fname

    return f"""' SAP VL06F Bot v2.3 | {datetime.now():%Y-%m-%d %H:%M}
' Plants  : {", ".join(plants)}
' Status  : {st_lo}-{st_hi} | Dates: {dt_lo}-{dt_hi} | Layout: {layout}
' Export  : {export_path}

Dim exportPath : exportPath = "{export_path}"
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(fso.GetParentFolderName(exportPath)) Then
    fso.CreateFolder fso.GetParentFolderName(exportPath)
End If

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "vl06f"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtIT_VKORG-LOW").text  = "1005"
session.findById("wnd[0]/usr/ctxtIT_VKORG-HIGH").text = "2205"
session.findById("wnd[0]/usr/ctxtIT_VTWEG-LOW").text  = "10"
session.findById("wnd[0]/usr/ctxtIT_VTWEG-HIGH").text = "20"
session.findById("wnd[0]/usr/ctxtIT_SPART-LOW").text  = "00"
session.findById("wnd[0]/usr/ctxtIT_SPART-HIGH").text = "07"
session.findById("wnd[0]/usr/ctxtIT_SPART-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtIT_SPART-HIGH").caretPosition = 2
session.findById("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH").press
{pl}session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/chkIF_ITEM").selected  = {b(ck_item)}
session.findById("wnd[0]/usr/chkIF_ANZPO").selected = {b(ck_anzpo)}
session.findById("wnd[0]/usr/chkIF_SPD_A").selected = {b(ck_spd_a)}
session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").text  = "{dt_lo}"
session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").text = "{dt_hi}"
session.findById("wnd[0]/usr/ctxtIT_WBSTK-LOW").text  = "{st_lo}"
session.findById("wnd[0]/usr/ctxtIT_WBSTK-HIGH").text = "{st_hi}"
session.findById("wnd[0]/usr/chkIF_SPD_A").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 33
session.findById("wnd[1]/usr").verticalScrollbar.position = 1073
session.findById("wnd[1]/usr").verticalScrollbar.position = 1147
session.findById("wnd[1]/usr").verticalScrollbar.position = 1098
session.findById("wnd[1]/usr/lbl[1,23]").setFocus
session.findById("wnd[1]/usr/lbl[1,23]").caretPosition = 10
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]").sendVKey 43

WScript.Sleep 1500
On Error Resume Next
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 500
session.findById("wnd[1]/usr/ctxtDY_PATH").text = fso.GetParentFolderName(exportPath)
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fso.GetFileName(exportPath)
session.findById("wnd[1]/tbar[0]/btn[11]").press
On Error GoTo 0

WScript.Sleep 2000
If fso.FileExists(exportPath) Then
    Dim shell : Set shell = CreateObject("WScript.Shell")
    shell.Run Chr(34) & exportPath & Chr(34), 1, False
    MsgBox "Export selesai!" & vbCrLf & "File: " & exportPath, vbInformation, "SAP VL06F Bot"
Else
    MsgBox "Export selesai, file tidak ditemukan:" & vbCrLf & exportPath, vbExclamation, "SAP VL06F Bot"
End If
"""


# ── Scheduler helpers ─────────────────────────────────────────────────
def register_task(task_name, vbs_path, run_dt):
    date_str = run_dt.strftime("%m/%d/%Y")
    time_str = run_dt.strftime("%H:%M")
    cmd = [
        "schtasks",
        "/Create",
        "/F",
        "/TN",
        task_name,
        "/TR",
        f'cscript //nologo "{vbs_path}"',
        "/SC",
        "ONCE",
        "/SD",
        date_str,
        "/ST",
        time_str,
        "/RL",
        "HIGHEST",
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        if r.returncode == 0:
            return (
                True,
                f"Task '{task_name}' terdaftar.\nJadwal: {run_dt:%d %b %Y %H:%M}",
            )
        return False, f"schtasks error:\n{r.stderr or r.stdout}"
    except FileNotFoundError:
        return False, "schtasks tidak ditemukan. Pastikan Windows."
    except Exception as e:
        return False, str(e)


def delete_task(task_name):
    try:
        r = subprocess.run(
            ["schtasks", "/Delete", "/F", "/TN", task_name],
            capture_output=True,
            text=True,
            timeout=10,
        )
        return r.returncode == 0
    except Exception:
        return False


# ── Scheduler Dialog ──────────────────────────────────────────────────
class SchedulerDialog(tk.Toplevel):
    def __init__(self, parent, get_script_fn, get_params_fn):
        super().__init__(parent)
        self.get_script = get_script_fn
        self.get_params = get_params_fn
        self.title("⏰  Jadwal Otomatis")
        self.configure(bg=SURF, padx=16, pady=14)
        self.resizable(False, False)
        self.geometry("500x320")

        self._task_name = tk.StringVar(value="SAP_VL06F_AutoRun")
        self._hour = tk.StringVar(value="00")
        self._min = tk.StringVar(value="01")
        self._vbs_path_var = tk.StringVar(
            value=str(Path.home() / "Documents" / "SAP_VL06F_Scheduler.vbs")
        )
        self._log_var = tk.StringVar(value="")

        self._build()

        # FIX v2.3: wait_visibility() sebelum grab_set() → tidak blank
        self.wait_visibility()
        self.grab_set()

    def _build(self):
        tk.Label(
            self, text="⏰  Jadwal Windows Task Scheduler", bg=SURF, fg=TEXT, font=FH
        ).pack(anchor="w", pady=(0, 4))
        tk.Label(
            self,
            text="Script VBS disimpan permanen dan didaftarkan ke\n"
            "Windows Task Scheduler. App boleh ditutup setelahnya.",
            bg=SURF,
            fg=MUTED,
            font=FNS,
            justify="left",
        ).pack(anchor="w")
        divider(self)

        r = tk.Frame(self, bg=SURF)
        r.pack(fill="x", pady=3)
        tk.Label(
            r, text="Nama Task :", bg=SURF, fg=MUTED, font=FN, width=12, anchor="e"
        ).pack(side="left")
        mkentry(r, var=self._task_name, w=28).pack(side="left", ipady=4, padx=(4, 0))

        r2 = tk.Frame(self, bg=SURF)
        r2.pack(fill="x", pady=3)
        tk.Label(
            r2, text="Tanggal   :", bg=SURF, fg=MUTED, font=FN, width=12, anchor="e"
        ).pack(side="left")
        self.sched_date = DateEntry(
            r2, init=datetime.today() + timedelta(days=1), bg=SURF
        )
        self.sched_date.pack(side="left", padx=(4, 12))
        tk.Label(r2, text="Pukul :", bg=SURF, fg=MUTED, font=FN).pack(side="left")
        mkentry(r2, var=self._hour, w=3).pack(side="left", ipady=4, padx=(4, 2))
        tk.Label(r2, text=":", bg=SURF, fg=MUTED, font=FNB).pack(side="left")
        mkentry(r2, var=self._min, w=3).pack(side="left", ipady=4, padx=(2, 0))

        tk.Label(self, text="Lokasi file VBS:", bg=SURF, fg=MUTED, font=FNS).pack(
            anchor="w", pady=(8, 2)
        )
        pr = tk.Frame(self, bg=SURF)
        pr.pack(fill="x")
        mkentry(pr, var=self._vbs_path_var, w=38, mono=True).pack(
            side="left", ipady=4, padx=(0, 4)
        )
        mkbtn(pr, "...", self._browse_vbs, px=6, py=4).pack(side="left")

        divider(self)

        bf = tk.Frame(self, bg=SURF)
        bf.pack(fill="x")
        mkbtn(bf, "✅  Daftar Jadwal", self._register, bg=GREEN, fg=BG).pack(
            side="left", padx=(0, 6)
        )
        mkbtn(bf, "🗑  Hapus Task", self._delete, bg=RED, fg=BG).pack(
            side="left", padx=(0, 6)
        )
        mkbtn(bf, "✕  Tutup", self.destroy, bg=SURF2, fg=MUTED).pack(side="right")

        tk.Label(
            self,
            textvariable=self._log_var,
            bg=SURF2,
            fg=ACCENT,
            font=FM,
            wraplength=440,
            justify="left",
            padx=8,
            pady=6,
        ).pack(fill="x", pady=(8, 0))

    def _browse_vbs(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".vbs",
            filetypes=[("VBScript", "*.vbs")],
            initialfile="SAP_VL06F_Scheduler.vbs",
            title="Lokasi simpan VBS",
            parent=self,
        )
        if p:
            self._vbs_path_var.set(p)

    def _get_run_dt(self):
        try:
            d = datetime.strptime(self.sched_date.get(), "%d.%m.%Y")
            return d.replace(
                hour=int(self._hour.get()), minute=int(self._min.get()), second=0
            )
        except Exception:
            messagebox.showerror(
                "Format Salah",
                "Tanggal: DD.MM.YYYY | Jam: 0-23 | Menit: 0-59",
                parent=self,
            )
            return None

    def _register(self):
        script = self.get_script()
        if not script:
            return
        run_dt = self._get_run_dt()
        if not run_dt:
            return
        if run_dt <= datetime.now():
            messagebox.showwarning(
                "Waktu Lampau", "Jadwal harus di masa depan.", parent=self
            )
            return
        vbs_path = self._vbs_path_var.get().strip()
        if not vbs_path:
            messagebox.showerror("Path Kosong", "Tentukan lokasi VBS.", parent=self)
            return
        try:
            os.makedirs(os.path.dirname(vbs_path), exist_ok=True)
            with open(vbs_path, "w", encoding="utf-8") as f:
                f.write(script)
        except Exception as e:
            messagebox.showerror("Gagal Simpan", str(e), parent=self)
            return
        task_name = self._task_name.get().strip() or "SAP_VL06F_AutoRun"
        ok, msg = register_task(task_name, vbs_path, run_dt)
        self._log_var.set(
            ("✅ " if ok else "❌ ") + msg + (f"\nVBS: {vbs_path}" if ok else "")
        )
        if not ok:
            messagebox.showerror("Gagal Daftar", msg, parent=self)

    def _delete(self):
        task_name = self._task_name.get().strip() or "SAP_VL06F_AutoRun"
        if delete_task(task_name):
            self._log_var.set(f"🗑 Task '{task_name}' dihapus.")
        else:
            self._log_var.set(f"⚠ Gagal hapus '{task_name}' (mungkin tidak ada).")


# ── MAIN APP ──────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SAP VL06F Bot v2.3")
        self.configure(bg=SURF)
        self.resizable(False, False)

        self.status_lo = tk.StringVar(value="A")
        self.status_hi = tk.StringVar(value="A")
        self.layout_var = tk.StringVar(value=DEFAULT_LAYOUTS[0])
        self.ck_item = tk.BooleanVar(value=True)
        self.ck_anzpo = tk.BooleanVar(value=True)
        self.ck_spd_a = tk.BooleanVar(value=True)
        self.log_var = tk.StringVar(value="Siap.")
        self.new_plant = tk.StringVar()
        self.custom_lo = tk.StringVar()
        self.export_dir = tk.StringVar(value=DEFAULT_EXPORT_DIR)

        self._build()
        # seed default plants
        self.plant_lb.set_plants(DEFAULT_PLANTS)

    # ── build UI ──────────────────────────────────────────────────────
    def _build(self):
        hdr = tk.Frame(self, bg=BG, padx=14, pady=8)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="⚙  SAP VL06F  Automation Bot v2.3", bg=BG, fg=TEXT, font=FH
        ).pack(side="left")
        tk.Label(
            hdr, text="VL06F · Scheduler · Auto-Open Excel", bg=BG, fg=MUTED, font=FNS
        ).pack(side="right")

        cols = tk.Frame(self, bg=SURF)
        cols.pack(fill="both")
        left = tk.Frame(cols, bg=SURF, padx=12, pady=8)
        right = tk.Frame(cols, bg=SURF, padx=12, pady=8)
        left.pack(side="left", fill="both", expand=True)
        tk.Frame(cols, bg=BORDER, width=1).pack(side="left", fill="y", pady=6)
        right.pack(side="left", fill="both", expand=True)

        self._panel_left(left)
        self._panel_right(right)

        bot = tk.Frame(self, bg=BG, padx=14, pady=8)
        bot.pack(fill="x")
        tk.Label(
            bot, textvariable=self.log_var, bg=BG, fg=MUTED, font=FNS, anchor="w"
        ).pack(fill="x", pady=(0, 6))
        acts = tk.Frame(bot, bg=BG)
        acts.pack(fill="x")
        mkbtn(acts, "👁 Preview", self._preview, bg=SURF2, fg=ACCENT).pack(
            side="left", padx=(0, 5)
        )
        mkbtn(acts, "💾 Simpan .vbs", self._save, bg=AMBER, fg=BG).pack(
            side="left", padx=(0, 5)
        )
        mkbtn(acts, "⏰ Jadwalkan", self._schedule, bg=INDIGO, fg=BG).pack(
            side="left", padx=(0, 5)
        )
        self._btn_run = mkbtn(acts, "▶  Jalankan SAP", self._run, bg=ACCENT, fg=BG)
        self._btn_run.pack(side="left")

    # ── LEFT PANEL ────────────────────────────────────────────────────
    def _panel_left(self, p):
        sec_label(p, "🏭", "Plant / Shipping Point")

        # PlantListbox — instant untuk 100+ plant
        self.plant_lb = PlantListbox(p, bg=SURF)
        self.plant_lb.pack(fill="x", pady=(0, 4))

        # quick actions
        qr = tk.Frame(p, bg=SURF)
        qr.pack(fill="x", pady=(3, 3))
        mkbtn(
            qr, "✔ Semua", self.plant_lb.select_all, bg=SURF2, fg=GREEN, px=8, py=3
        ).pack(side="left", padx=(0, 4))
        mkbtn(
            qr,
            "☐ Batal Pilih",
            self.plant_lb.deselect_all,
            bg=SURF2,
            fg=MUTED,
            px=8,
            py=3,
        ).pack(side="left", padx=(0, 4))
        mkbtn(
            qr,
            "🗑 Hapus Dipilih",
            self.plant_lb.remove_selected,
            bg=SURF2,
            fg=RED,
            px=8,
            py=3,
        ).pack(side="left", padx=(0, 4))
        mkbtn(qr, "🗑 Hapus Semua", self._clear_all, bg=SURF2, fg=RED, px=8, py=3).pack(
            side="left"
        )

        # paste row
        pr = tk.Frame(p, bg=SURF)
        pr.pack(fill="x", pady=(2, 0))
        mkbtn(pr, "📋 Paste", self._paste, bg=SURF2, fg=INDIGO, px=8, py=3).pack(
            side="left", padx=(0, 8)
        )

        # manual add
        ar = tk.Frame(p, bg=SURF)
        ar.pack(fill="x", pady=(4, 0))
        e = mkentry(ar, var=self.new_plant, w=10)
        e.pack(side="left", ipady=4, padx=(0, 4))
        e.bind("<Return>", lambda _: self._add_manual())
        mkbtn(ar, "+ Tambah", self._add_manual, px=8, py=3).pack(side="left")
        tk.Label(ar, text="  koma/spasi OK", bg=SURF, fg=MUTED, font=FNS).pack(
            side="left"
        )

        divider(p)

        sec_label(p, "📋", "Status Delivery")
        for var, prefix in [(self.status_lo, "Dari"), (self.status_hi, "S/d ")]:
            fr = tk.Frame(p, bg=SURF)
            fr.pack(fill="x", pady=1)
            tk.Label(
                fr, text=f"{prefix}:", bg=SURF, fg=MUTED, font=FN, width=5, anchor="e"
            ).pack(side="left")
            for val in ("A", "B", "C"):
                tk.Radiobutton(
                    fr,
                    text=val,
                    variable=var,
                    value=val,
                    bg=SURF,
                    fg=TEXT,
                    selectcolor=SURF2,
                    activebackground=SURF,
                    font=FNB,
                    cursor="hand2",
                ).pack(side="left", padx=4)

        divider(p)

        sec_label(p, "☑", "Filter SAP")
        for var, label, tip in [
            (self.ck_item, "Item Detail", "Tampilkan data per baris item."),
            (self.ck_anzpo, "Jumlah Posisi", "Tampilkan kolom total posisi."),
            (self.ck_spd_a, "Speed A", "Saring hanya delivery Speed-A."),
        ]:
            row = tk.Frame(p, bg=SURF)
            row.pack(fill="x", pady=2)
            tk.Checkbutton(
                row,
                text=label,
                variable=var,
                bg=SURF,
                fg=TEXT,
                selectcolor=SURF2,
                activebackground=SURF,
                font=FNB,
                cursor="hand2",
            ).pack(side="left")
            tk.Label(
                row,
                text=f"  ← {tip}",
                bg=SURF,
                fg=MUTED,
                font=FNS,
                wraplength=200,
                justify="left",
            ).pack(side="left")

    # ── RIGHT PANEL ───────────────────────────────────────────────────
    def _panel_right(self, p):
        sec_label(p, "📅", "Tanggal Goods Issue")
        dg = tk.Frame(p, bg=SURF)
        dg.pack(fill="x")
        for label, attr, init in [
            ("Dari  :", "date_lo", datetime(2025, 1, 1)),
            ("S/d   :", "date_hi", datetime.today()),
        ]:
            r = tk.Frame(dg, bg=SURF)
            r.pack(fill="x", pady=2)
            tk.Label(
                r, text=label, bg=SURF, fg=MUTED, font=FN, width=7, anchor="e"
            ).pack(side="left")
            de = DateEntry(r, init=init, bg=SURF)
            de.pack(side="left")
            setattr(self, attr, de)

        qd = tk.Frame(p, bg=SURF)
        qd.pack(fill="x", pady=(5, 0))
        tk.Label(qd, text="Cepat:", bg=SURF, fg=MUTED, font=FNS).pack(
            side="left", padx=(0, 4)
        )
        for t, d in [("7h", 7), ("30h", 30), ("90h", 90), ("YTD", 0)]:
            mkbtn(qd, t, lambda x=d: self._quick(x), px=7, py=2).pack(
                side="left", padx=2
            )

        divider(p)

        sec_label(p, "📂", "Folder Export Excel")
        ef = tk.Frame(p, bg=SURF)
        ef.pack(fill="x")
        mkentry(ef, var=self.export_dir, w=24, mono=True).pack(
            side="left", ipady=4, padx=(0, 4)
        )
        mkbtn(ef, "...", self._browse_export, px=6, py=4).pack(side="left")
        tk.Label(
            p,
            text="File Excel dibuka otomatis setelah export.",
            bg=SURF,
            fg=MUTED,
            font=FNS,
        ).pack(anchor="w", pady=(2, 0))

        divider(p)

        sec_label(p, "🗂", "Layout SAP")
        for lo in DEFAULT_LAYOUTS:
            r = tk.Frame(p, bg=SURF)
            r.pack(fill="x", pady=1)
            tk.Radiobutton(
                r,
                text=lo,
                variable=self.layout_var,
                value=lo,
                bg=SURF,
                fg=TEXT,
                selectcolor=SURF2,
                activebackground=SURF,
                font=FM,
                cursor="hand2",
            ).pack(side="left")

        cr = tk.Frame(p, bg=SURF)
        cr.pack(fill="x", pady=(6, 0))
        tk.Label(cr, text="Custom:", bg=SURF, fg=MUTED, font=FNS).pack(
            side="left", padx=(0, 4)
        )
        mkentry(cr, var=self.custom_lo, w=16, mono=True).pack(
            side="left", ipady=4, padx=(0, 4)
        )
        mkbtn(cr, "Pakai", self._use_custom, bg=SURF2, fg=INDIGO, px=7, py=3).pack(
            side="left"
        )

        divider(p)

        sec_label(p, "📊", "Ringkasan")
        self.summary_var = tk.StringVar(value="Tekan ↻ untuk lihat ringkasan")
        tk.Label(
            p,
            textvariable=self.summary_var,
            bg=SURF2,
            fg=ACCENT,
            font=FM,
            justify="left",
            wraplength=270,
            padx=8,
            pady=6,
            anchor="w",
        ).pack(fill="x")
        mkbtn(
            p, "↻  Refresh Ringkasan", self._refresh, bg=SURF2, fg=MUTED, px=8, py=3
        ).pack(fill="x", pady=(5, 0))

    # ── actions ───────────────────────────────────────────────────────
    def _clear_all(self):
        if not self.plant_lb.get_all():
            return
        if messagebox.askyesno(
            "Hapus Semua Plant",
            f"Hapus semua {len(self.plant_lb.get_all())} plant?\n\n"
            "Gunakan 📋 Paste untuk isi ulang.",
            icon="warning",
        ):
            self.plant_lb.clear_all()
            self._log("🗑 Semua plant dihapus.", AMBER)

    def _paste(self):
        try:
            txt = self.clipboard_get()
            codes = [c for c in re.split(r"[\s,;|\t]+", txt.strip()) if c]
            before = len(self.plant_lb.get_all())
            self.plant_lb.add_plants(codes)
            added = len(self.plant_lb.get_all()) - before
            self._log(f"📋 Paste: {added} baru ({len(codes)} diparsing).", ACCENT)
        except Exception:
            messagebox.showinfo("Clipboard", "Clipboard kosong / tidak bisa dibaca.")

    def _add_manual(self):
        txt = self.new_plant.get().strip()
        if txt:
            codes = [c for c in re.split(r"[\s,;|\t]+", txt) if c]
            self.plant_lb.add_plants(codes)
            self.new_plant.set("")

    def _quick(self, days):
        today = datetime.today()
        lo = datetime(today.year, 1, 1) if days == 0 else today - timedelta(days=days)
        self.date_lo.var.set(lo.strftime("%d.%m.%Y"))
        self.date_hi.var.set(today.strftime("%d.%m.%Y"))

    def _use_custom(self):
        v = self.custom_lo.get().strip()
        if v:
            self.layout_var.set(v)
            self._log(f"Layout → {v}", ACCENT)
        else:
            messagebox.showwarning("Kosong", "Isi nama layout custom dulu.")

    def _browse_export(self):
        d = filedialog.askdirectory(
            title="Pilih Folder Export", initialdir=self.export_dir.get()
        )
        if d:
            self.export_dir.set(d)

    def _log(self, msg, color=MUTED):
        self.log_var.set(f"[{datetime.now():%H:%M:%S}]  {msg}")

    def _params(self):
        plants = self.plant_lb.get_selected()
        if not plants:
            messagebox.showerror(
                "Plant Kosong", "Pilih (klik/highlight) minimal 1 plant di daftar."
            )
            return None
        for dv in (self.date_lo.get(), self.date_hi.get()):
            try:
                datetime.strptime(dv, "%d.%m.%Y")
            except Exception:
                messagebox.showerror("Tanggal Salah", "Format: DD.MM.YYYY")
                return None
        exp_dir = self.export_dir.get().strip() or DEFAULT_EXPORT_DIR
        return dict(
            plants=plants,
            st_lo=self.status_lo.get(),
            st_hi=self.status_hi.get(),
            dt_lo=self.date_lo.get(),
            dt_hi=self.date_hi.get(),
            layout=self.layout_var.get(),
            ck_item=self.ck_item.get(),
            ck_anzpo=self.ck_anzpo.get(),
            ck_spd_a=self.ck_spd_a.get(),
            export_dir=exp_dir,
        )

    def _script(self):
        p = self._params()
        return gen_vbs(**p) if p else None

    def _refresh(self):
        p = self._params()
        if not p:
            return
        self.summary_var.set(
            f"Plants   : {', '.join(p['plants'])} ({len(p['plants'])}/{len(self.plant_lb.get_all())})\n"
            f"Status   : {p['st_lo']} – {p['st_hi']}\n"
            f"Tanggal  : {p['dt_lo']}  s/d  {p['dt_hi']}\n"
            f"Layout   : {p['layout']}\n"
            f"Export → : {p['export_dir']}"
        )

    def _preview(self):
        s = self._script()
        if not s:
            return
        w = tk.Toplevel(self)
        w.title("Preview VBS")
        w.configure(bg=BG)
        w.geometry("700x520")
        t = tk.Text(
            w,
            bg=SURF,
            fg=MUTED,
            font=FM,
            relief="flat",
            bd=0,
            highlightthickness=0,
            wrap="none",
        )
        sb = ttk.Scrollbar(w, command=t.yview)
        t.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        t.pack(fill="both", expand=True, padx=8, pady=8)
        t.insert("end", s)
        t.config(state="disabled")

    def _save(self):
        s = self._script()
        if not s:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".vbs",
            filetypes=[("VBScript", "*.vbs"), ("All", "*.*")],
            initialfile="sap_vl06f.vbs",
            title="Simpan Script",
        )
        if path:
            open(path, "w", encoding="utf-8").write(s)
            self._log(f"Disimpan: {path}", GREEN)
            messagebox.showinfo("Tersimpan", f"File disimpan:\n{path}")

    def _schedule(self):
        SchedulerDialog(self, self._script, self._params)

    def _set_running(self, running: bool):
        if running:
            self._btn_run.config(
                text="⏳ Sedang Berjalan…", state="disabled", bg=SURF2, fg=MUTED
            )
        else:
            self._btn_run.config(
                text="▶  Jalankan SAP", state="normal", bg=ACCENT, fg=BG
            )

    def _run(self):
        s = self._script()
        if not s:
            return
        p = self._params()
        self._log(
            f"Menjalankan … {len(p['plants'])} plant | layout={p['layout']}", ACCENT
        )
        self._set_running(True)

        tmp = tempfile.NamedTemporaryFile(
            mode="w", suffix=".vbs", delete=False, encoding="utf-8"
        )
        tmp.write(s)
        tmp.close()
        tmp_path = tmp.name

        def _worker():
            try:
                r = subprocess.run(
                    ["cscript", "//nologo", tmp_path],
                    capture_output=True,
                    text=True,
                    timeout=600,
                )
                if r.returncode == 0:
                    self.after(
                        0,
                        lambda: self._log(
                            "✅ Selesai! Excel akan terbuka otomatis.", GREEN
                        ),
                    )
                else:
                    msg = r.stderr[:120] or r.stdout[:120]
                    self.after(
                        0, lambda: self._log(f"⚠ Exit {r.returncode}: {msg}", AMBER)
                    )
            except FileNotFoundError:
                self.after(0, lambda: self._log("❌ cscript tidak ditemukan.", RED))
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        "Error",
                        "cscript tidak ditemukan di sistem ini.\n\n"
                        "Gunakan 'Simpan .vbs' lalu double-click manual\n"
                        "di komputer Windows yang terhubung SAP.",
                    ),
                )
            except subprocess.TimeoutExpired:
                self.after(0, lambda: self._log("⏱ Timeout (>10 menit).", RED))
            except Exception as ex:
                self.after(0, lambda: self._log(f"❌ {ex}", RED))
            finally:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass
                self.after(0, lambda: self._set_running(False))

        threading.Thread(target=_worker, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
