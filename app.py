"""
AutoPresent - PowerPoint Auto Presenter with Text-to-Speech
Reads slide notes aloud and automatically advances to the next slide.
"""

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import time
import os

from pptx import Presentation
from lxml import etree
import win32com.client
import pythoncom

# XML namespaces used in PPTX notes slides
_PPTX_NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


# ---------------------------------------------------------------------------
# TTS Engine wrapper (lives on a dedicated thread)
# ---------------------------------------------------------------------------

class TTSEngine:
    """
    Text-to-speech using Windows SAPI (SpVoice) directly via win32com.

    Threading model
    ---------------
    - _init_engine() and speak() run on the presenter background thread.
      The SpVoice COM object is NEVER touched from the GUI thread.
    - pause() / resume() / stop() only set threading.Event flags.
      The speak() polling loop reads those flags and calls voice.Pause()
      / voice.Resume() / voice.Skip() all from the presenter thread.
    - get_voices_sync() / set_*() are safe to call from any thread.
    """

    # SAPI RunningState values
    _RS_DONE    = 0
    _RS_PAUSED  = 1
    _RS_SPEAKING = 2

    # SAPI Speak flags
    _ASYNC = 1           # start speaking without blocking
    _PURGE = 2           # stop & clear the queue

    def __init__(self):
        self._voice = None

        # Settings queued from GUI thread, flushed in _apply_pending()
        self._lock = threading.Lock()
        self._pending_rate:        int | None = None
        self._pending_volume:      int | None = None
        self._pending_voice_name:  str | None = None

        # Control flags set by GUI thread, polled by presenter thread
        self._stop_flag  = threading.Event()
        self._pause_flag = threading.Event()   # set = paused

    # ------------------------------------------------------------------
    # GUI-thread helpers (no COM access)
    # ------------------------------------------------------------------

    def get_voices_sync(self):
        """Enumerate SAPI voices. Returns objects with .name / .id."""
        pythoncom.CoInitialize()
        try:
            v = win32com.client.Dispatch("SAPI.SpVoice")
            tokens = v.GetVoices()
            voices = []
            for i in range(tokens.Count):
                tok = tokens.Item(i)
                class _V: pass
                vobj = _V()
                vobj.name = tok.GetDescription()
                vobj.id   = tok.GetDescription()
                voices.append(vobj)
            del v
            return voices
        finally:
            pythoncom.CoUninitialize()

    def set_rate(self, rate_wpm: int):
        """Map wpm (80-300) → SAPI rate (-10..+10)."""
        sapi_rate = int((rate_wpm - 175) / 12.5)
        sapi_rate = max(-10, min(10, sapi_rate))
        with self._lock:
            self._pending_rate = sapi_rate

    def set_volume(self, volume: float):
        """Map 0.0-1.0 → SAPI volume 0-100."""
        with self._lock:
            self._pending_volume = int(volume * 100)

    def set_voice(self, voice_name: str):
        with self._lock:
            self._pending_voice_name = voice_name

    def pause(self):
        """Signal speak() loop to pause. GUI-thread safe."""
        self._pause_flag.set()

    def resume(self):
        """Signal speak() loop to resume. GUI-thread safe."""
        self._pause_flag.clear()

    def stop(self):
        """Signal speak() loop to abort. GUI-thread safe."""
        self._stop_flag.set()
        self._pause_flag.clear()   # unblock if paused

    # ------------------------------------------------------------------
    # Presenter-thread methods
    # ------------------------------------------------------------------

    def _init_engine(self):
        """Create SpVoice on the presenter thread. Called once from _run()."""
        self._voice = win32com.client.Dispatch("SAPI.SpVoice")
        self._stop_flag.clear()
        self._pause_flag.clear()

    def _apply_pending(self):
        """Flush queued property changes into SpVoice."""
        with self._lock:
            if self._pending_rate is not None:
                self._voice.Rate = self._pending_rate
                self._pending_rate = None
            if self._pending_volume is not None:
                self._voice.Volume = self._pending_volume
                self._pending_volume = None
            if self._pending_voice_name is not None:
                name = self._pending_voice_name
                self._pending_voice_name = None
                tokens = self._voice.GetVoices()
                for i in range(tokens.Count):
                    tok = tokens.Item(i)
                    if tok.GetDescription() == name:
                        self._voice.Voice = tok
                        break

    def speak(self, text: str) -> bool:
        """
        Speak text and block until done, paused-then-resumed, or stopped.
        Must be called from the presenter thread.
        Returns True if speech completed normally, False if interrupted.
        """
        # Clear stop flag; pause_flag is intentionally NOT cleared here —
        # the caller (Presenter.resume) already cleared it before calling speak().
        self._stop_flag.clear()
        self._apply_pending()

        # Start speaking asynchronously so this thread stays unblocked
        self._voice.Speak(text, self._ASYNC)

        currently_paused = False

        while True:
            # --- stop requested ---
            if self._stop_flag.is_set():
                self._voice.Speak("", self._ASYNC | self._PURGE)
                return False

            # --- pause requested ---
            if self._pause_flag.is_set():
                if not currently_paused:
                    self._voice.Pause()
                    currently_paused = True
                time.sleep(0.05)
                continue

            # --- resume if we were paused ---
            if currently_paused:
                self._voice.Resume()
                currently_paused = False

            state = self._voice.Status.RunningState
            if state != self._RS_SPEAKING:
                # Speech finished (or was already done)
                break

            time.sleep(0.05)

        return True


# ---------------------------------------------------------------------------
# PowerPoint controller
# ---------------------------------------------------------------------------

class PPTController:
    """
    Manages PowerPoint note extraction (python-pptx) and COM slideshow control.

    IMPORTANT – COM threading rule:
      open_slideshow(), goto_slide(), and close() must ALL be called from the
      SAME thread (the presenter background thread).  COM objects are
      apartment-threaded and cannot be shared across threads.
    """

    def __init__(self, filepath: str):
        self.filepath = os.path.abspath(filepath)
        self.prs = Presentation(filepath)   # python-pptx – thread-safe, for notes
        self._app = None
        self._presentation = None
        self._slideshow = None

    # ---- note extraction (safe to call from any thread) -----------------

    def get_notes(self, slide_index: int) -> str:
        """
        Return the speaker notes for the given 0-based slide index.

        Uses direct XPath extraction instead of notes_text_frame so that it
        works regardless of placeholder index (which varies across PPTX files
        and PowerPoint versions).
        """
        slide = self.prs.slides[slide_index]
        notes_slide = slide.notes_slide
        if notes_slide is None:
            return ""
        texts = []
        for sp in notes_slide._element.findall('.//p:sp', _PPTX_NSMAP):
            ph = sp.find('.//p:ph', _PPTX_NSMAP)
            if ph is None:
                continue
            # Skip non-body placeholders (slide image, page number, date, etc.)
            if ph.get('type', 'body') in ('sldImg', 'sldNum', 'dt', 'hdr', 'ftr'):
                continue
            tx_body = sp.find('p:txBody', _PPTX_NSMAP)
            if tx_body is None:
                continue
            for t_elem in tx_body.findall('.//a:t', _PPTX_NSMAP):
                if t_elem.text:
                    texts.append(t_elem.text)
        return ' '.join(texts).strip()

    @property
    def slide_count(self) -> int:
        return len(self.prs.slides)

    # ---- COM slideshow control (must stay on ONE thread) ----------------

    def open_slideshow(self):
        """
        Launch PowerPoint and start the slideshow.
        Must be called on the presenter thread AFTER CoInitialize().
        """
        self._app = win32com.client.Dispatch("PowerPoint.Application")
        self._app.Visible = True
        self._presentation = self._app.Presentations.Open(
            self.filepath, ReadOnly=True, Untitled=False, WithWindow=True
        )
        self._presentation.SlideShowSettings.Run()
        # Wait for the SlideShowWindow to become available
        for _ in range(20):
            time.sleep(0.3)
            try:
                if self._presentation.SlideShowWindow is not None:
                    self._slideshow = self._presentation.SlideShowWindow.View
                    break
            except Exception:
                pass
        # Give the window a moment to settle, then force slide 1
        time.sleep(0.3)
        if self._slideshow is not None:
            try:
                self._slideshow.GotoSlide(1)
            except Exception:
                pass

    def goto_slide(self, slide_number: int):
        """Jump to a 1-based slide number. Must be called on the same thread as open_slideshow()."""
        if self._slideshow is None:
            return
        try:
            self._slideshow.GotoSlide(slide_number)
        except Exception:
            pass

    def close(self):
        """Close the presentation. Must be called on the same thread as open_slideshow()."""
        try:
            if self._presentation:
                self._presentation.Close()
        except Exception:
            pass
        try:
            if self._app:
                self._app.Quit()
        except Exception:
            pass
        finally:
            self._slideshow = None
            self._presentation = None
            self._app = None


# ---------------------------------------------------------------------------
# Presenter (orchestrates TTS + slide advancing)
# ---------------------------------------------------------------------------

class Presenter:
    """Runs the auto-presentation loop on a background thread."""

    def __init__(self, ppt: PPTController, tts: TTSEngine,
                 on_slide_change, on_status_change, on_finished):
        self._ppt = ppt
        self._tts = tts
        self._on_slide_change = on_slide_change   # callback(slide_index)
        self._on_status_change = on_status_change # callback(message)
        self._on_finished = on_finished           # callback()
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()  # not paused initially
        self.current_slide = 0

    def start(self, from_slide: int = 0):
        self._stop_event.clear()
        self._pause_event.set()
        self.current_slide = from_slide
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def pause(self):
        # First tell TTS to pause (sets _pause_flag → speak() calls voice.Pause())
        self._tts.pause()
        # Then block the between-slide loop
        self._pause_event.clear()

    def resume(self):
        # First clear TTS pause flag so speak() calls voice.Resume()
        self._tts.resume()
        # Then unblock the between-slide loop
        self._pause_event.set()

    def stop(self):
        self._tts.stop()             # purge speech immediately
        self._stop_event.set()
        self._pause_event.set()      # unblock between-slide wait if paused

    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _run(self):
        """
        Main presentation loop.
        All COM calls (open_slideshow / goto_slide / close) happen here so
        they share the same COM apartment.
        """
        pythoncom.CoInitialize()
        try:
            # Init SpVoice TTS on this thread
            self._tts._init_engine()

            # Open the slideshow on this thread
            self._on_status_change("Opening PowerPoint slideshow…")
            try:
                self._ppt.open_slideshow()
            except Exception as e:
                self._on_status_change(f"Error opening slideshow: {e}")
                self._on_finished()
                return

            total = self._ppt.slide_count

            for idx in range(self.current_slide, total):
                if self._stop_event.is_set():
                    break

                # If paused between slides, wait here until resumed or stopped
                while not self._pause_event.is_set():
                    if self._stop_event.is_set():
                        break
                    time.sleep(0.05)
                if self._stop_event.is_set():
                    break

                self.current_slide = idx
                slide_num = idx + 1

                # Advance PowerPoint to this slide (same thread → works)
                self._ppt.goto_slide(slide_num)
                self._on_slide_change(idx)
                self._on_status_change(f"Slide {slide_num} / {total}")

                notes = self._ppt.get_notes(idx)

                if notes:
                    self._on_status_change(
                        f"Slide {slide_num} / {total}  —  Speaking…"
                    )
                    self._tts.speak(notes)
                    # If stopped mid-speech, break out of the slide loop
                    if self._stop_event.is_set():
                        break
                else:
                    # No notes: 3 s pause, interruptible by stop/pause
                    self._on_status_change(
                        f"Slide {slide_num} / {total}  —  No notes, waiting 3 s…"
                    )
                    for _ in range(60):
                        if self._stop_event.is_set():
                            break
                        if not self._pause_event.is_set():
                            # paused — keep waiting without consuming the timer
                            time.sleep(0.05)
                            continue
                        time.sleep(0.05)

            if not self._stop_event.is_set():
                self._on_status_change("Presentation finished.")
                self._on_finished()

        finally:
            # Close COM objects on the same thread they were created on
            self._ppt.close()
            pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class AutoPresentApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AutoPresent")
        self.resizable(False, False)
        self.configure(bg="#1e1e2e")

        self._ppt: PPTController | None = None
        self._tts = TTSEngine()
        self._presenter: Presenter | None = None
        self._paused = False

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---- UI construction -------------------------------------------------

    def _build_ui(self):
        PAD = 12
        BG = "#1e1e2e"
        FG = "#cdd6f4"
        ACCENT = "#89b4fa"
        ENTRY_BG = "#313244"
        BTN_BG = "#45475a"
        BTN_FG = "#cdd6f4"

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TScale", background=BG, troughcolor=ENTRY_BG,
                        sliderlength=18, sliderrelief="flat")
        style.configure("TCombobox", fieldbackground=ENTRY_BG,
                        background=BTN_BG, foreground=FG,
                        selectbackground=ENTRY_BG, selectforeground=FG)
        style.map("TCombobox", fieldbackground=[("readonly", ENTRY_BG)])

        # ---- File row ----
        file_frame = tk.Frame(self, bg=BG)
        file_frame.grid(row=0, column=0, columnspan=3,
                        padx=PAD, pady=(PAD, 4), sticky="ew")

        tk.Label(file_frame, text="PowerPoint file:", bg=BG, fg=FG,
                 font=("Segoe UI", 10)).pack(side="left")

        self._file_var = tk.StringVar(value="(none selected)")
        tk.Label(file_frame, textvariable=self._file_var, bg=BG,
                 fg=ACCENT, font=("Segoe UI", 10, "italic"),
                 width=40, anchor="w").pack(side="left", padx=8)

        tk.Button(file_frame, text="Browse…", bg=BTN_BG, fg=BTN_FG,
                  font=("Segoe UI", 9), relief="flat", padx=6,
                  command=self._browse_file).pack(side="left")

        # ---- Settings ----
        settings_frame = tk.LabelFrame(self, text=" Settings ",
                                       bg=BG, fg=FG,
                                       font=("Segoe UI", 9, "bold"),
                                       labelanchor="nw", bd=1,
                                       relief="groove")
        settings_frame.grid(row=1, column=0, columnspan=3,
                            padx=PAD, pady=4, sticky="ew")
        settings_frame.columnconfigure(1, weight=1)

        # Speech rate
        tk.Label(settings_frame, text="Speech rate:", bg=BG, fg=FG,
                 font=("Segoe UI", 9)).grid(row=0, column=0,
                                            padx=8, pady=4, sticky="w")
        self._rate_var = tk.IntVar(value=175)
        rate_scale = ttk.Scale(settings_frame, from_=80, to=300,
                               variable=self._rate_var, orient="horizontal",
                               length=200,
                               command=lambda v: self._on_rate_change())
        rate_scale.grid(row=0, column=1, padx=8, pady=4, sticky="ew")
        self._rate_label = tk.Label(settings_frame, text="175 wpm",
                                    bg=BG, fg=FG, font=("Segoe UI", 9),
                                    width=8)
        self._rate_label.grid(row=0, column=2, padx=(0, 8))

        # Volume
        tk.Label(settings_frame, text="Volume:", bg=BG, fg=FG,
                 font=("Segoe UI", 9)).grid(row=1, column=0,
                                            padx=8, pady=4, sticky="w")
        self._vol_var = tk.DoubleVar(value=1.0)
        vol_scale = ttk.Scale(settings_frame, from_=0.0, to=1.0,
                              variable=self._vol_var, orient="horizontal",
                              length=200,
                              command=lambda v: self._on_vol_change())
        vol_scale.grid(row=1, column=1, padx=8, pady=4, sticky="ew")
        self._vol_label = tk.Label(settings_frame, text="100%",
                                   bg=BG, fg=FG, font=("Segoe UI", 9),
                                   width=8)
        self._vol_label.grid(row=1, column=2, padx=(0, 8))

        # Voice selector
        tk.Label(settings_frame, text="Voice:", bg=BG, fg=FG,
                 font=("Segoe UI", 9)).grid(row=2, column=0,
                                            padx=8, pady=4, sticky="w")
        self._voice_var = tk.StringVar()
        self._voice_combo = ttk.Combobox(settings_frame,
                                         textvariable=self._voice_var,
                                         state="readonly", width=35)
        self._voice_combo.grid(row=2, column=1, columnspan=2,
                               padx=8, pady=4, sticky="ew")
        self._populate_voices()
        self._voice_combo.bind("<<ComboboxSelected>>", self._on_voice_change)

        # Start slide
        tk.Label(settings_frame, text="Start from slide:", bg=BG, fg=FG,
                 font=("Segoe UI", 9)).grid(row=3, column=0,
                                            padx=8, pady=4, sticky="w")
        self._start_slide_var = tk.IntVar(value=1)
        self._start_slide_spin = tk.Spinbox(
            settings_frame, from_=1, to=999,
            textvariable=self._start_slide_var,
            width=6, bg=ENTRY_BG, fg=FG,
            buttonbackground=BTN_BG, relief="flat",
            font=("Segoe UI", 9)
        )
        self._start_slide_spin.grid(row=3, column=1, padx=8, pady=4,
                                     sticky="w")

        # ---- Progress / status ----
        prog_frame = tk.Frame(self, bg=BG)
        prog_frame.grid(row=2, column=0, columnspan=3,
                        padx=PAD, pady=4, sticky="ew")
        prog_frame.columnconfigure(0, weight=1)

        self._status_var = tk.StringVar(value="Ready. Open a .pptx file to begin.")
        tk.Label(prog_frame, textvariable=self._status_var,
                 bg=BG, fg=FG, font=("Segoe UI", 9),
                 anchor="w").grid(row=0, column=0, sticky="ew")

        self._progress = ttk.Progressbar(prog_frame, orient="horizontal",
                                         mode="determinate", length=460)
        self._progress.grid(row=1, column=0, pady=(2, 0), sticky="ew")

        # Current slide notes preview
        self._notes_text = tk.Text(self, height=7, width=58, wrap="word",
                                   bg=ENTRY_BG, fg=FG, relief="flat",
                                   font=("Segoe UI", 9),
                                   state="disabled", bd=0,
                                   padx=6, pady=4)
        self._notes_text.grid(row=3, column=0, columnspan=3,
                              padx=PAD, pady=4)

        # ---- Control buttons ----
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.grid(row=4, column=0, columnspan=3,
                       padx=PAD, pady=(4, PAD))

        btn_cfg = dict(font=("Segoe UI", 10, "bold"), relief="flat",
                       width=12, padx=4, pady=6)

        self._start_btn = tk.Button(btn_frame, text="▶  Start",
                                    bg="#a6e3a1", fg="#1e1e2e",
                                    command=self._on_start, **btn_cfg)
        self._start_btn.pack(side="left", padx=4)

        self._pause_btn = tk.Button(btn_frame, text="⏸  Pause",
                                    bg="#f9e2af", fg="#1e1e2e",
                                    command=self._on_pause_resume,
                                    state="disabled", **btn_cfg)
        self._pause_btn.pack(side="left", padx=4)

        self._stop_btn = tk.Button(btn_frame, text="⏹  Stop",
                                   bg="#f38ba8", fg="#1e1e2e",
                                   command=self._on_stop,
                                   state="disabled", **btn_cfg)
        self._stop_btn.pack(side="left", padx=4)

    # ---- Voice population -----------------------------------------------

    def _populate_voices(self):
        voices = self._tts.get_voices_sync()
        names = [v.name for v in voices]
        self._voices = voices
        self._voice_combo["values"] = names
        if names:
            self._voice_combo.current(0)
            self._voice_var.set(names[0])

    # ---- Callbacks -------------------------------------------------------

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select PowerPoint file",
            filetypes=[("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*")]
        )
        if not path:
            return
        self._load_file(path)

    def _load_file(self, path: str):
        try:
            self._ppt = PPTController(path)
            short = os.path.basename(path)
            self._file_var.set(short)
            n = self._ppt.slide_count
            self._progress["maximum"] = n
            self._progress["value"] = 0
            self._start_slide_spin.config(to=n)
            self._start_slide_var.set(1)
            self._status_var.set(f"Loaded: {short}  ({n} slides)")
            self._show_notes(0)
            self._start_btn.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")

    def _on_rate_change(self):
        rate = self._rate_var.get()
        self._rate_label.config(text=f"{rate} wpm")
        self._tts.set_rate(rate)

    def _on_vol_change(self):
        vol = self._vol_var.get()
        self._vol_label.config(text=f"{int(vol * 100)}%")
        self._tts.set_volume(vol)

    def _on_voice_change(self, _event=None):
        idx = self._voice_combo.current()
        if idx >= 0 and idx < len(self._voices):
            self._tts.set_voice(self._voices[idx].id)

    def _on_start(self):
        if self._ppt is None:
            messagebox.showwarning("No file", "Please select a .pptx file first.")
            return

        # Apply current settings
        self._tts.set_rate(self._rate_var.get())
        self._tts.set_volume(self._vol_var.get())
        self._on_voice_change()

        start_idx = self._start_slide_var.get() - 1  # convert to 0-based

        self._paused = False
        self._presenter = Presenter(
            ppt=self._ppt,
            tts=self._tts,
            on_slide_change=self._cb_slide_change,
            on_status_change=self._cb_status,
            on_finished=self._cb_finished,
        )
        self._presenter.start(from_slide=start_idx)

        self._start_btn.config(state="disabled")
        self._pause_btn.config(state="normal")
        self._stop_btn.config(state="normal")

    def _on_pause_resume(self):
        if self._presenter is None:
            return
        if self._paused:
            self._presenter.resume()
            self._pause_btn.config(text="⏸  Pause")
            self._paused = False
        else:
            self._presenter.pause()
            self._pause_btn.config(text="▶  Resume")
            self._paused = True

    def _on_stop(self):
        if self._presenter:
            # stop() signals the presenter thread; it will call ppt.close()
            # on its own thread before exiting, so we don't close here.
            self._presenter.stop()
        self._reset_controls()
        self._status_var.set("Stopped.")

    def _reset_controls(self):
        self._start_btn.config(state="normal" if self._ppt else "disabled")
        self._pause_btn.config(state="disabled", text="⏸  Pause")
        self._stop_btn.config(state="disabled")
        self._paused = False

    # ---- Thread-safe UI callbacks ----------------------------------------

    def _cb_slide_change(self, slide_idx: int):
        self.after(0, self._update_slide_ui, slide_idx)

    def _cb_status(self, message: str):
        self.after(0, self._status_var.set, message)

    def _cb_finished(self):
        self.after(0, self._on_presentation_finished)

    def _update_slide_ui(self, slide_idx: int):
        self._progress["value"] = slide_idx + 1
        self._show_notes(slide_idx)

    def _show_notes(self, slide_idx: int):
        if self._ppt is None:
            return
        notes = self._ppt.get_notes(slide_idx)
        self._notes_text.config(state="normal")
        self._notes_text.delete("1.0", "end")
        if notes:
            self._notes_text.insert("end", notes)
        else:
            self._notes_text.insert("end", "(No notes for this slide)")
        self._notes_text.config(state="disabled")

    def _on_presentation_finished(self):
        self._reset_controls()
        self._progress["value"] = self._ppt.slide_count if self._ppt else 0
        messagebox.showinfo("AutoPresent", "Presentation finished!")

    # ---- Window close ----------------------------------------------------

    def _on_close(self):
        if self._presenter and self._presenter.is_running():
            if not messagebox.askyesno("Quit",
                                       "Presentation is running. Quit anyway?"):
                return
            self._presenter.stop()
            # Give the presenter thread a moment to close COM cleanly
            self._presenter._thread.join(timeout=3)
        self.destroy()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = AutoPresentApp()
    app.mainloop()
