"""Quote Generator — one-click desktop app for non-technical staff.

Flow:
  1. App opens to two large buttons: Personal Auto / Homeowners (HO3)
  2. Choosing one finds the newest matching JotForm export automatically
     (or opens a file picker if none is found)
  3. A clean list of customer names; click one
  4. Optionally add a cover photo, then one big "Generate Quote" button
  5. "Quote Created Successfully" -> "Open File"

Engineering notes:
  - Only tkinter is imported at startup so the window appears fast even as
    a PyInstaller onefile exe; pandas/python-pptx load in worker threads.
  - Every action is guarded by a busy flag, every worker catches all
    exceptions, and report_callback_exception is overridden — no raw
    Python error can ever reach the screen. Details are logged to
    Desktop/Quotes/error_log.txt for the administrator.
"""

from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

APP_TITLE = "Quote Generator"

# ---------------------------------------------------------------------------
# Professional palette: deep navy brand bar, blue accent, soft grey canvas
# ---------------------------------------------------------------------------
NAVY = "#16335B"          # header bar / Auto button
NAVY_HOVER = "#1D4276"
ACCENT = "#2D6CDF"        # primary action / Homeowners button
ACCENT_HOVER = "#3D7BEE"
ACCENT_DISABLED = "#AECBF5"
GREEN = "#1A7F37"
GREEN_HOVER = "#229944"
RED = "#C0392B"
BG = "#EEF1F6"            # window canvas
CARD = "#FFFFFF"          # card surfaces
TEXT = "#1A2433"          # near-black text
GREY = "#6B7280"          # secondary text
ROW_ALT = "#F4F7FB"       # alternating list rows
BADGE_BG = "#E4EDFC"      # quote-type pill

FONT = "Segoe UI"


def _hover(widget: tk.Widget, normal: str, hovered: str) -> None:
    """Simple flat-button hover effect."""
    widget.bind("<Enter>", lambda _e: widget.configure(bg=hovered))
    widget.bind("<Leave>", lambda _e: widget.configure(bg=normal))


class QuoteApp(tk.Tk):
    """Main window. Screens: 'home' (choose quote type) and 'list'."""

    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("560x700")
        self.minsize(480, 560)
        self.configure(bg=BG)

        self.rows: list = []                 # loaded submissions
        self.profile_key: str = "AUTO"       # which quote type is loaded
        self.generated_indices: set = set()  # rows generated this session
        self.last_output_path = None         # most recent quote file
        self.photo_path = None               # optional cover photo
        self._busy = False

        self._build_header()
        self._build_home_screen()
        self._build_list_screen()
        self._show_home()

    # ------------------------------------------------------------------
    # Safety net: never show a raw Python traceback
    # ------------------------------------------------------------------
    def report_callback_exception(self, exc_type, exc, tb) -> None:  # noqa: N802
        self._log(exc)
        messagebox.showinfo(
            APP_TITLE,
            "Something didn't work just now.\n\nPlease try again.",
        )

    @staticmethod
    def _log(exc) -> None:
        """Technical details go to Desktop/Quotes/error_log.txt, never the UI."""
        try:
            import proposal_generator as gen
            gen._log_error(exc)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Brand header (always visible)
    # ------------------------------------------------------------------
    def _build_header(self) -> None:
        header = tk.Frame(self, bg=NAVY)
        header.pack(fill="x", side="top")
        tk.Label(header, text="Quote Generator", bg=NAVY, fg="white",
                 font=(FONT, 16, "bold")).pack(side="left", padx=22, pady=14)
        tk.Label(header, text="Proposal builder", bg=NAVY, fg="#9FB6D9",
                 font=(FONT, 10)).pack(side="left", pady=(20, 14))
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x", side="top")

    # ------------------------------------------------------------------
    # Screen 1: choose the quote type
    # ------------------------------------------------------------------
    def _build_home_screen(self) -> None:
        self.home = tk.Frame(self, bg=BG)

        inner = tk.Frame(self.home, bg=BG)
        inner.place(relx=0.5, rely=0.40, anchor="center")

        tk.Label(inner, text="What would you like to quote?", bg=BG, fg=TEXT,
                 font=(FONT, 17, "bold")).pack(pady=(0, 26))

        self.btn_auto = tk.Button(
            inner, text="Personal Auto", command=lambda: self.choose_type("AUTO"),
            font=(FONT, 15, "bold"), bg=NAVY, fg="white",
            activebackground=NAVY_HOVER, activeforeground="white",
            relief="flat", bd=0, width=22, pady=18, cursor="hand2",
        )
        self.btn_auto.pack(pady=(0, 14))
        _hover(self.btn_auto, NAVY, NAVY_HOVER)

        self.btn_home_q = tk.Button(
            inner, text="Homeowners (HO3)", command=lambda: self.choose_type("HO3"),
            font=(FONT, 15, "bold"), bg=ACCENT, fg="white",
            activebackground=ACCENT_HOVER, activeforeground="white",
            relief="flat", bd=0, width=22, pady=18, cursor="hand2",
        )
        self.btn_home_q.pack()
        _hover(self.btn_home_q, ACCENT, ACCENT_HOVER)

        self.home_status = tk.Label(inner, text="", bg=BG, fg=GREY,
                                    font=(FONT, 11))
        self.home_status.pack(pady=(26, 0))

        manual = tk.Label(inner, text="open a file manually", bg=BG, fg=GREY,
                          font=(FONT, 9, "underline"), cursor="hand2")
        manual.pack(pady=(10, 0))
        manual.bind("<Button-1>", lambda _e: self.browse_excel(None))

    # ------------------------------------------------------------------
    # Screen 2: customer list + one big Generate button
    # ------------------------------------------------------------------
    def _build_list_screen(self) -> None:
        self.listscreen = tk.Frame(self, bg=BG)

        # Breadcrumb row: back link + quote-type badge
        nav = tk.Frame(self.listscreen, bg=BG)
        nav.pack(fill="x", padx=22, pady=(14, 4))
        back = tk.Label(nav, text="‹  Back", bg=BG, fg=GREY,
                        font=(FONT, 11), cursor="hand2")
        back.pack(side="left")
        back.bind("<Button-1>", lambda _e: self._go_home())
        self.badge = tk.Label(nav, text="", bg=BADGE_BG, fg=ACCENT,
                              font=(FONT, 10, "bold"), padx=12, pady=3)
        self.badge.pack(side="right")

        tk.Label(self.listscreen, text="Select a customer", bg=BG, fg=TEXT,
                 font=(FONT, 15, "bold")).pack(pady=(4, 10))

        # Customer list inside a white card
        card = tk.Frame(self.listscreen, bg=CARD, highlightthickness=1,
                        highlightbackground="#D8DEE9")
        card.pack(fill="both", expand=True, padx=22)

        self.listbox = tk.Listbox(
            card, font=(FONT, 13), activestyle="none",
            selectbackground=ACCENT, selectforeground="white",
            highlightthickness=0, relief="flat", bg=CARD, fg=TEXT,
        )
        sb = tk.Scrollbar(card, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        sb.pack(side="right", fill="y")
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        # Optional cover photo — one small line
        photo_row = tk.Frame(self.listscreen, bg=BG)
        photo_row.pack(pady=(10, 0))
        self.photo_link = tk.Label(
            photo_row, text="add a cover photo (optional)", bg=BG,
            font=(FONT, 10, "underline"), fg=GREY, cursor="hand2",
        )
        self.photo_link.pack(side="left")
        self.photo_link.bind("<Button-1>", lambda _e: self.choose_photo())
        self.photo_clear = tk.Label(
            photo_row, text="  remove", bg=BG,
            font=(FONT, 10, "underline"), fg=RED, cursor="hand2",
        )
        self.photo_clear.bind("<Button-1>", lambda _e: self.clear_photo())

        # The one action button
        self.btn_generate = tk.Button(
            self.listscreen, text="Generate Quote", command=self.generate_selected,
            font=(FONT, 15, "bold"), bg=ACCENT_DISABLED, fg="white",
            activebackground=ACCENT_HOVER, activeforeground="white",
            relief="flat", bd=0, padx=44, pady=14, cursor="hand2",
            state="disabled",
        )
        self.btn_generate.pack(pady=(14, 4))

        # Result area (hidden until a quote is created)
        self.result_frame = tk.Frame(self.listscreen, bg=BG)
        self.result_label = tk.Label(self.result_frame, text="", bg=BG,
                                     font=(FONT, 13, "bold"))
        self.result_label.pack(pady=(4, 8))
        self.btn_open_file = tk.Button(
            self.result_frame, text="Open File", command=self.open_last_file,
            font=(FONT, 12, "bold"), bg=GREEN, fg="white",
            activebackground=GREEN_HOVER, activeforeground="white",
            relief="flat", bd=0, padx=30, pady=9, cursor="hand2",
        )
        self.btn_open_file.pack()
        _hover(self.btn_open_file, GREEN, GREEN_HOVER)

        # Footer: source file name
        self.source_label = tk.Label(self.listscreen, text="", bg=BG, fg=GREY,
                                     font=(FONT, 9))
        self.source_label.pack(side="bottom", pady=(2, 10))

    # ------------------------------------------------------------------
    # Screen switching
    # ------------------------------------------------------------------
    def _show_home(self) -> None:
        self.listscreen.pack_forget()
        self.home.pack(fill="both", expand=True)

    def _show_list(self) -> None:
        self.home.pack_forget()
        self.listscreen.pack(fill="both", expand=True)

    def _go_home(self) -> None:
        if self._busy:
            return
        self.clear_photo()
        self._hide_result()
        self.home_status.configure(text="")
        self._show_home()

    def _show_result(self, message: str, ok: bool) -> None:
        self.result_label.configure(text=message, fg=GREEN if ok else RED)
        if ok:
            self.btn_open_file.pack()
        else:
            self.btn_open_file.pack_forget()
        self.result_frame.pack(pady=(0, 6))

    def _hide_result(self) -> None:
        self.result_frame.pack_forget()

    def _set_generate_enabled(self, enabled: bool) -> None:
        if enabled:
            self.btn_generate.configure(bg=ACCENT, state="normal")
        else:
            self.btn_generate.configure(bg=ACCENT_DISABLED, state="disabled")

    # ------------------------------------------------------------------
    # Choosing a quote type -> auto-find the newest matching export
    # ------------------------------------------------------------------
    def choose_type(self, profile_key: str) -> None:
        if self._busy:
            return
        self._busy = True
        label = "Auto" if profile_key == "AUTO" else "Homeowners"
        self.home_status.configure(text=f"Looking for your newest {label} quote list…")
        threading.Thread(target=self._autodetect_worker,
                         args=(profile_key,), daemon=True).start()

    def _autodetect_worker(self, profile_key: str) -> None:
        try:
            import data_loader
            for candidate in data_loader.find_export_candidates(limit=8):
                try:
                    rows, pk = data_loader.load_submissions(
                        candidate, expected_profile=profile_key
                    )
                except Exception:
                    continue  # wrong type or not a valid export — try next
                self.after(0, self._on_load_done, rows, pk, candidate.name)
                return
        except Exception as exc:
            self._log(exc)
        # Nothing matched: let the user pick the file themselves
        self.after(0, self._autodetect_failed, profile_key)

    def _autodetect_failed(self, profile_key: str) -> None:
        self._busy = False
        self.home_status.configure(text="")
        self.browse_excel(profile_key)

    # ------------------------------------------------------------------
    # Manual file selection
    # ------------------------------------------------------------------
    def browse_excel(self, expected_profile) -> None:
        if self._busy:
            return
        label = {"AUTO": "Auto", "HO3": "Homeowners"}.get(expected_profile, "quote")
        path = filedialog.askopenfilename(
            title=f"Choose the {label} Excel file exported from JotForm",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path:
            return
        self._busy = True
        self.home_status.configure(text="Loading…")
        threading.Thread(target=self._load_worker,
                         args=(path, expected_profile), daemon=True).start()

    def _load_worker(self, path: str, expected_profile) -> None:
        try:
            import data_loader
            rows, profile_key = data_loader.load_submissions(
                path, expected_profile=expected_profile
            )
        except Exception as exc:
            friendly = (
                str(exc)
                if exc.__class__.__name__ == "ExcelLoadError"
                else "Wrong file selected — please export from JotForm again."
            )
            self._log(exc)
            self.after(0, self._on_load_failed, friendly)
            return
        import os.path
        self.after(0, self._on_load_done, rows, profile_key, os.path.basename(path))

    def _on_load_failed(self, friendly_message: str) -> None:
        self._busy = False
        self.home_status.configure(text="")
        messagebox.showinfo(APP_TITLE, friendly_message)

    def _on_load_done(self, rows: list, profile_key: str, source_name: str) -> None:
        """Fill the list with customer names only."""
        import data_loader
        import quote_profiles
        self._busy = False
        self.home_status.configure(text="")
        self.rows = rows
        self.profile_key = profile_key
        self.generated_indices.clear()
        self.last_output_path = None
        self.clear_photo()
        self._hide_result()

        self.listbox.delete(0, "end")
        for i, row in enumerate(rows):
            name = data_loader.display_value(row.get("Client Name")) or f"Customer {i + 1}"
            company = data_loader.get_optional(row, "Company Name") or ""
            label = f"  {name}" + (f"   ({company})" if company else "")
            self.listbox.insert("end", label)
            if i % 2 == 1:
                self.listbox.itemconfigure(i, background=ROW_ALT)

        self.badge.configure(text=quote_profiles.PROFILES[profile_key]["label"])
        self.source_label.configure(text=source_name)
        self._set_generate_enabled(False)
        self._show_list()

    # ------------------------------------------------------------------
    # Selection
    # ------------------------------------------------------------------
    def _selected_index(self):
        sel = self.listbox.curselection()
        return sel[0] if sel else None

    def _on_select(self, _event=None) -> None:
        self._hide_result()
        self.clear_photo()  # a chosen photo belongs to one customer only
        self._set_generate_enabled(not self._busy and self._selected_index() is not None)

    # ------------------------------------------------------------------
    # Optional cover photo
    # ------------------------------------------------------------------
    def choose_photo(self) -> None:
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="Choose a photo for the front page",
            filetypes=[("Pictures", "*.jpg *.jpeg *.png"), ("All files", "*.*")],
        )
        if not path:
            return
        self.photo_path = path
        import os.path
        self.photo_link.configure(
            text="photo: " + os.path.basename(path), fg=GREEN
        )
        self.photo_clear.pack(side="left")

    def clear_photo(self) -> None:
        self.photo_path = None
        self.photo_link.configure(text="add a cover photo (optional)", fg=GREY)
        self.photo_clear.pack_forget()

    # ------------------------------------------------------------------
    # Quote generation
    # ------------------------------------------------------------------
    def generate_selected(self) -> None:
        idx = self._selected_index()
        if idx is None or self._busy:
            return
        row = self.rows[idx]

        import data_loader
        import proposal_generator as gen

        # Rows missing essential data are explained, never crash
        if data_loader.missing_required_fields(row):
            messagebox.showinfo(
                APP_TITLE,
                "This customer's information is incomplete, so a quote "
                "can't be created for them.",
            )
            return

        # Duplicate check — simple yes/no in plain language
        name = data_loader.display_value(row.get("Client Name"))
        try:
            existing = gen.existing_quotes_for_row(row, gen.default_output_dir())
        except Exception:
            existing = []
        if idx in self.generated_indices or existing:
            if not messagebox.askyesno(
                APP_TITLE,
                f"A quote for {name} was already created.\n\nMake another one?",
            ):
                return

        self._busy = True
        self._hide_result()
        self.btn_generate.configure(text="Creating…")
        self._set_generate_enabled(False)
        threading.Thread(
            target=self._generate_worker,
            args=(idx, row, self.photo_path), daemon=True,
        ).start()

    def _generate_worker(self, idx: int, row, photo_path) -> None:
        try:
            import proposal_generator as gen
            output_path = gen.generate_quote(
                row, profile_key=self.profile_key, photo_path=photo_path
            )
        except Exception as exc:
            friendly = (
                str(exc)
                if exc.__class__.__name__ == "QuoteGenerationError"
                else "The quote could not be created. Please try again."
            )
            self._log(exc)
            self.after(0, self._on_generate_done, idx, None, friendly)
            return
        self.after(0, self._on_generate_done, idx, output_path, None)

    def _on_generate_done(self, idx: int, output_path, error) -> None:
        self._busy = False
        self.btn_generate.configure(text="Generate Quote")
        self._set_generate_enabled(self._selected_index() is not None)

        if error:
            self._show_result("Quote could not be created", ok=False)
            messagebox.showinfo(APP_TITLE, error)
            return

        self.generated_indices.add(idx)
        self.last_output_path = output_path
        self.clear_photo()  # photo was for this quote; don't reuse by accident
        self._show_result("✓  Quote Created Successfully", ok=True)

    # ------------------------------------------------------------------
    # Opening the result
    # ------------------------------------------------------------------
    def open_last_file(self) -> None:
        if not self.last_output_path:
            return
        try:
            if sys.platform == "win32":
                os.startfile(self.last_output_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", str(self.last_output_path)])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(self.last_output_path)])
        except Exception as exc:
            self._log(exc)
            messagebox.showinfo(
                APP_TITLE,
                "The file couldn't open automatically.\n\n"
                "You can find it in the 'Quotes' folder on your Desktop.",
            )


def main() -> None:
    app = QuoteApp()
    app.mainloop()


if __name__ == "__main__":
    main()
