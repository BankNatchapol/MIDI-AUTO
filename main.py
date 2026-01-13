import time
import threading
import json
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

import mido
from pynput.keyboard import Controller

kb = Controller()

APP_DIR = Path.cwd()
MIDIS_DIR = APP_DIR / "midis"
CONFIG_FILE = APP_DIR / "presets.json"


# ===========================
# Keymaps
# ===========================
# Semitone indices: 0..11 = C, C#, D, D#, E, F, F#, G, G#, A, A#, B
#
# Mac/PlayCover (your first image):
#   HIGH: q 2 w 3 e r 5 t 6 y 7 u + (extra top C) i
#   MID : z s x d c v g b h n j m
#   LOW : , l . ; / o 0 p - [ = ]
#
# Windows (your second image):
#   HIGH: q 2 w 3 e r 5 t 6 y 7 u + (extra top C) i
#   MID : z s x d c v g b h n j m
#   LOW : l . ; ' / o 0 p - [ = ]
#
# Note: Windows LOW D# is apostrophe "'".


def get_keymaps(use_windows: bool):
    # top row (same for both)
    HIGH = {
        0: 'q',  1: '2',  2: 'w',  3: '3',  4: 'e',  5: 'r',
        6: '5',  7: 't',  8: '6',  9: 'y', 10: '7', 11: 'u',
    }
    HIGH_PLUS_C = 'i'  # extra top C

    # middle row (same for both)
    MID = {
        0: 'z',  1: 's',  2: 'x',  3: 'd',  4: 'c',  5: 'v',
        6: 'g',  7: 'b',  8: 'h',  9: 'n', 10: 'j', 11: 'm',
    }

    if use_windows:
        # Windows low row (from your Windows screenshot)
        LOW = {
            0: 'l',   # C
            1: '.',   # C#
            2: ';',   # D
            3: "'",   # D#
            4: '/',   # E
            5: 'o',   # F
            6: '0',   # F#
            7: 'p',   # G
            8: '-',   # G#
            9: '[',   # A
            10: '=',  # A#
            11: ']',  # B
        }
    else:
        # Mac/PlayCover low row (from your Mac screenshot)
        LOW = {
            0: ',',  1: 'l',  2: '.',  3: ';',  4: '/',  5: 'o',
            6: '0',  7: 'p',  8: '-',  9: '[', 10: '=', 11: ']',
        }

    return LOW, MID, HIGH, HIGH_PLUS_C


@dataclass
class Config:
    midi_path: str = ""
    base_c_midi: int = 48
    transpose: int = 0
    speed: float = 1.0
    lead_in: float = 2.0
    trim_silence: bool = True

    tap_mode: bool = True
    tap_ms: int = 18

    use_windows_map: bool = False  # checkbox toggle


def midi_note_to_key(note: int, cfg: Config) -> Optional[str]:
    """Convert MIDI note to a mapped keyboard key, or None if out of range."""
    LOW, MID, HIGH, HIGH_PLUS_C = get_keymaps(cfg.use_windows_map)

    note += cfg.transpose
    d = note - cfg.base_c_midi
    if d < 0:
        return None

    octave = d // 12
    semitone = d % 12

    if octave == 0:
        return LOW.get(semitone)
    if octave == 1:
        return MID.get(semitone)
    if octave == 2:
        return HIGH.get(semitone)
    if octave == 3 and semitone == 0:
        return HIGH_PLUS_C
    return None


def collect_abs_timed_messages(midi_path: str) -> List[Tuple[float, mido.Message]]:
    """
    Return list of (abs_time_seconds, msg) in playback order.
    When iterating a MidiFile, msg.time is seconds since previous message. :contentReference[oaicite:4]{index=4}
    """
    mid = mido.MidiFile(midi_path)
    out: List[Tuple[float, mido.Message]] = []
    t = 0.0
    for msg in mid:
        t += float(msg.time or 0.0)
        out.append((t, msg))
    return out


def find_trim_window(timed: List[Tuple[float, mido.Message]]) -> Tuple[float, float]:
    """Trim window [start, end] based on first note_on and last note event."""
    if not timed:
        return 0.0, 0.0

    start = None
    end = None
    last_time = timed[-1][0]

    for t, msg in timed:
        if msg.type == "note_on" and getattr(msg, "velocity", 0) > 0:
            start = t
            break

    for t, msg in reversed(timed):
        if msg.type == "note_off":
            end = t
            break
        if msg.type == "note_on":
            end = t
            break

    if start is None:
        start = 0.0
    if end is None:
        end = last_time
    if end < start:
        end = start
    return start, end


def group_by_time(timed: List[Tuple[float, mido.Message]], eps: float = 1e-9):
    """Yield (t, [msgs]) grouped by identical timestamps (for chord taps)."""
    if not timed:
        return
    i = 0
    n = len(timed)
    while i < n:
        t0 = timed[i][0]
        msgs = [timed[i][1]]
        i += 1
        while i < n and abs(timed[i][0] - t0) <= eps:
            msgs.append(timed[i][1])
            i += 1
        yield t0, msgs


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MIDI → Sky Piano (Mac/Windows Keymap Toggle)")
        self.geometry("940x700")

        self.cfg = Config()
        self._stop_event = threading.Event()
        self._play_thread: Optional[threading.Thread] = None

        self.presets: Dict[str, Dict[str, Any]] = self._load_presets()

        self._build_ui()
        self._refresh_presets_dropdown()
        self._update_test_button_text()

    # ---------- storage ----------
    def _ensure_dirs(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        MIDIS_DIR.mkdir(parents=True, exist_ok=True)

    def _unique_dest(self, dest: Path) -> Path:
        if not dest.exists():
            return dest
        stem, suffix = dest.stem, dest.suffix
        for i in range(1, 10_000):
            cand = dest.with_name(f"{stem} ({i}){suffix}")
            if not cand.exists():
                return cand
        raise RuntimeError("Too many duplicate MIDI filenames in storage folder.")

    def _import_midi_to_storage(self, src_path: Path) -> Path:
        self._ensure_dirs()
        if not src_path.exists():
            raise FileNotFoundError(src_path)
        dest = self._unique_dest(MIDIS_DIR / src_path.name)
        shutil.copy2(src_path, dest)  # copy into storage :contentReference[oaicite:5]{index=5}
        return dest

    # ---------- presets ----------
    def _load_presets(self) -> Dict[str, Dict[str, Any]]:
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
        return {}

    def _save_presets(self) -> None:
        self._ensure_dirs()
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.presets, f, indent=2)

    def _refresh_presets_dropdown(self) -> None:
        names = sorted(self.presets.keys(), key=str.lower)
        self.preset_combo["values"] = names
        if names and self.preset_var.get() not in names:
            self.preset_var.set(names[0])
            self.apply_preset()

    def save_preset(self) -> None:
        name = self.preset_name_var.get().strip()
        if not name:
            messagebox.showerror("Missing name", "Type a preset name first.")
            return
        if not self.cfg.midi_path:
            messagebox.showerror("No MIDI", "Choose MIDI… first (it imports into storage).")
            return

        stored_path = Path(self.cfg.midi_path)
        if not stored_path.exists():
            messagebox.showerror("Missing MIDI", "Current MIDI file does not exist.")
            return

        try:
            stored_rel = stored_path.relative_to(APP_DIR)
        except ValueError:
            messagebox.showerror(
                "MIDI not imported",
                "This MIDI isn't in app storage.\nUse Choose MIDI… (copies into storage), then save again."
            )
            return

        existed = name in self.presets  # overwrite if same name

        self.presets[name] = {
            "midi_relpath": str(stored_rel),
            "base_c_midi": int(self.base_c.get()),
            "transpose": int(self.transpose.get()),
            "speed": float(self.speed.get()),
            "lead_in": float(self.lead_in.get()),
            "trim_silence": bool(self.trim_silence.get()),
            "tap_mode": bool(self.tap_mode.get()),
            "tap_ms": int(self.tap_ms.get()),
            "use_windows_map": bool(self.use_windows_map.get()),
        }

        self._save_presets()
        self._refresh_presets_dropdown()
        self.preset_var.set(name)
        self._log(f"{'Overwrote' if existed else 'Saved'} preset '{name}' (MIDI + settings).")

    def apply_preset(self, _event=None) -> None:
        """Auto-called when selecting dropdown (no Apply button). :contentReference[oaicite:6]{index=6}"""
        name = self.preset_var.get().strip()
        if not name:
            return
        preset = self.presets.get(name)
        if not isinstance(preset, dict):
            return

        # Stop if playing
        if self._play_thread and self._play_thread.is_alive():
            self._stop_event.set()
            self._log("Stopping playback to switch preset…")

        self.base_c.set(int(preset.get("base_c_midi", 48)))
        self.transpose.set(int(preset.get("transpose", 0)))
        self.speed.set(float(preset.get("speed", 1.0)))
        self.lead_in.set(float(preset.get("lead_in", 2.0)))
        self.trim_silence.set(bool(preset.get("trim_silence", True)))
        self.tap_mode.set(bool(preset.get("tap_mode", True)))
        self.tap_ms.set(int(preset.get("tap_ms", 18)))
        self.use_windows_map.set(bool(preset.get("use_windows_map", False)))
        self._update_tap_ui_state()
        self._update_test_button_text()

        # Load MIDI from storage
        rel = preset.get("midi_relpath", "")
        midi_path = (APP_DIR / rel).resolve() if rel else None
        if not midi_path or not midi_path.exists():
            self.cfg.midi_path = ""
            self.file_label.config(text="(Missing MIDI in storage)")
            self.play_btn.config(state="disabled")
            self._log(f"Preset '{name}' MIDI missing.")
            return

        self.cfg.midi_path = str(midi_path)
        self.file_label.config(text=midi_path.name)
        self.play_btn.config(state="normal")

        self.preset_name_var.set(name)
        self._log(f"Loaded preset '{name}' → {midi_path.name}")

    # ---------- UI helpers ----------
    def _ui(self, fn):
        self.after(0, fn)

    def _log(self, s: str):
        self.log.insert("end", s + "\n")
        self.log.see("end")

    def _slider(self, parent, label, var, a, b, step):
        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label, width=34, anchor="w").pack(side="left")
        tk.Scale(row, variable=var, from_=a, to=b, orient="horizontal", resolution=step).pack(
            side="left", fill="x", expand=True
        )
        tk.Label(row, textvariable=var, width=10, anchor="e").pack(side="right")

    def _update_tap_ui_state(self):
        state = "normal" if self.tap_mode.get() else "disabled"
        try:
            self.tap_scale.configure(state=state)
        except Exception:
            pass

    def _update_test_button_text(self):
        # Show what key is LOW C for the selected map
        tmp = Config(use_windows_map=bool(self.use_windows_map.get()))
        low_c_key = midi_note_to_key(tmp.base_c_midi, tmp)  # base C itself
        low_c_key = low_c_key if low_c_key else "?"
        mode = "Windows" if self.use_windows_map.get() else "Mac/PlayCover"
        self.test_btn.configure(text=f"Test LOW C ({mode}) = '{low_c_key}'")

    # ---------- UI ----------
    def _build_ui(self):
        frm = tk.Frame(self, padx=12, pady=12)
        frm.pack(fill="both", expand=True)

        # MIDI row
        file_row = tk.Frame(frm)
        file_row.pack(fill="x", pady=(0, 10))

        self.file_label = tk.Label(file_row, text="No MIDI selected", anchor="w")
        self.file_label.pack(side="left", fill="x", expand=True)

        tk.Button(file_row, text="Choose MIDI…", command=self.choose_midi).pack(side="right")

        # Presets
        box = tk.LabelFrame(frm, text="Presets (select = auto-load MIDI + settings)", padx=10, pady=10)
        box.pack(fill="x", pady=(0, 10))

        r1 = tk.Frame(box)
        r1.pack(fill="x", pady=(0, 6))

        tk.Label(r1, text="Preset name:", width=12, anchor="w").pack(side="left")
        self.preset_name_var = tk.StringVar()
        tk.Entry(r1, textvariable=self.preset_name_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(r1, text="Save (overwrite)", command=self.save_preset).pack(side="right")

        r2 = tk.Frame(box)
        r2.pack(fill="x")

        tk.Label(r2, text="Preset:", width=12, anchor="w").pack(side="left")
        self.preset_var = tk.StringVar(value="")
        self.preset_combo = ttk.Combobox(r2, textvariable=self.preset_var, state="readonly")
        self.preset_combo.pack(side="left", fill="x", expand=True)
        self.preset_combo.bind("<<ComboboxSelected>>", self.apply_preset)  # :contentReference[oaicite:7]{index=7}

        # Settings
        settings = tk.LabelFrame(frm, text="Playback settings", padx=10, pady=10)
        settings.pack(fill="x", pady=(0, 10))

        self.base_c = tk.IntVar(value=self.cfg.base_c_midi)
        self.transpose = tk.IntVar(value=self.cfg.transpose)
        self.speed = tk.DoubleVar(value=self.cfg.speed)
        self.lead_in = tk.DoubleVar(value=self.cfg.lead_in)

        self.trim_silence = tk.BooleanVar(value=True)
        self.tap_mode = tk.BooleanVar(value=True)
        self.tap_ms = tk.IntVar(value=18)

        # Keymap toggle checkbox
        self.use_windows_map = tk.BooleanVar(value=False)  # BooleanVar is standard for checkboxes :contentReference[oaicite:8]{index=8}

        self._slider(settings, "Base C MIDI (octave align)", self.base_c, 24, 84, 1)
        self._slider(settings, "Transpose (semitones)", self.transpose, -24, 24, 1)
        self._slider(settings, "Speed", self.speed, 0.25, 3.0, 0.05)
        self._slider(settings, "Lead-in seconds (focus game)", self.lead_in, 0.0, 10.0, 0.25)

        chk_row = tk.Frame(settings)
        chk_row.pack(fill="x", pady=(8, 0))
        tk.Checkbutton(chk_row, text="Trim start/end silence", variable=self.trim_silence).pack(side="left")

        chk_row2 = tk.Frame(settings)
        chk_row2.pack(fill="x", pady=(6, 0))
        tk.Checkbutton(
            chk_row2,
            text="Tap mode (fix overlaps / retrigger every note)",
            variable=self.tap_mode,
            command=self._update_tap_ui_state
        ).pack(side="left")

        tap_row = tk.Frame(settings)
        tap_row.pack(fill="x", pady=(4, 0))
        tk.Label(tap_row, text="Tap duration (ms)", width=34, anchor="w").pack(side="left")
        self.tap_scale = tk.Scale(tap_row, variable=self.tap_ms, from_=1, to=80,
                                  orient="horizontal", resolution=1)
        self.tap_scale.pack(side="left", fill="x", expand=True)
        tk.Label(tap_row, textvariable=self.tap_ms, width=10, anchor="e").pack(side="right")
        self._update_tap_ui_state()

        map_row = tk.Frame(settings)
        map_row.pack(fill="x", pady=(10, 0))
        tk.Checkbutton(
            map_row,
            text="Use Windows keymap (black keys: 2 3 5 6 7 / S D G H J / . ' 0 - =)",
            variable=self.use_windows_map,
            command=self._update_test_button_text
        ).pack(side="left")

        # Buttons
        btns = tk.Frame(frm)
        btns.pack(fill="x", pady=(0, 10))

        self.play_btn = tk.Button(btns, text="▶ Play", command=self.play, state="disabled")
        self.play_btn.pack(side="left")

        self.stop_btn = tk.Button(btns, text="■ Stop", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        self.test_btn = tk.Button(btns, text="Test LOW C", command=self.test_note)
        self.test_btn.pack(side="right")

        # Log
        log_frame = tk.LabelFrame(frm, text="Log", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True)

        self.log = tk.Text(log_frame, height=12, wrap="word")
        self.log.pack(fill="both", expand=True)

        self._log("Choose MIDI… copies it into app storage so presets never break if you move/delete originals.")
        self._log("If notes blend into a long click: enable Tap mode and try 12–25ms.")
        self._log("Tip: click the game window during Lead-in so it receives keys.")

    # ---------- actions ----------
    def choose_midi(self):
        path = filedialog.askopenfilename(
            title="Select MIDI file",
            filetypes=[("MIDI files", "*.mid *.midi"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            stored = self._import_midi_to_storage(Path(path))
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
            return

        self.cfg.midi_path = str(stored)
        self.file_label.config(text=stored.name)
        self.play_btn.config(state="normal")
        self._log(f"Imported MIDI into storage: {stored.name}")

    def test_note(self):
        # Press and release the LOW C key according to current keymap
        tmp_cfg = Config(
            base_c_midi=int(self.base_c.get()),
            transpose=int(self.transpose.get()),
            use_windows_map=bool(self.use_windows_map.get()),
        )
        k = midi_note_to_key(tmp_cfg.base_c_midi, tmp_cfg)
        if not k:
            self._log("Test note failed: LOW C not mapped (check Base C MIDI).")
            return
        kb.press(k)
        time.sleep(0.05)
        kb.release(k)
        self._log(f"Sent test LOW C key '{k}'")

    def play(self):
        if not self.cfg.midi_path:
            messagebox.showerror("No MIDI", "Choose MIDI… first.")
            return

        # Snapshot settings
        self.cfg.base_c_midi = int(self.base_c.get())
        self.cfg.transpose = int(self.transpose.get())
        self.cfg.speed = float(self.speed.get())
        self.cfg.lead_in = float(self.lead_in.get())
        self.cfg.trim_silence = bool(self.trim_silence.get())
        self.cfg.tap_mode = bool(self.tap_mode.get())
        self.cfg.tap_ms = int(self.tap_ms.get())
        self.cfg.use_windows_map = bool(self.use_windows_map.get())

        # Stop any existing playback
        if self._play_thread and self._play_thread.is_alive():
            self._stop_event.set()
            time.sleep(0.05)

        self._stop_event.clear()
        self.play_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        self._play_thread = threading.Thread(target=self._play_worker, daemon=True)
        self._play_thread.start()

    def stop(self):
        self._stop_event.set()
        self._log("Stop requested.")

    def _play_worker(self):
        try:
            midi_path = self.cfg.midi_path  # snapshot
            if not Path(midi_path).exists():
                self._ui(lambda: self._log(f"ERROR: MIDI missing: {midi_path}"))
                return

            self._ui(lambda: self._log(f"Lead-in {self.cfg.lead_in:.2f}s — focus game window now!"))
            time.sleep(self.cfg.lead_in)

            timed = collect_abs_timed_messages(midi_path)

            # Trim silence
            if self.cfg.trim_silence and timed:
                start_t, end_t = find_trim_window(timed)
                timed = [(t, msg) for (t, msg) in timed if start_t <= t <= end_t]
                self._ui(lambda: self._log(f"Trim: start={start_t:.3f}s end={end_t:.3f}s"))

            if not timed:
                self._ui(lambda: self._log("No messages to play (empty after trim)."))
                return

            groups = list(group_by_time(timed))
            prev_t = groups[0][0]
            tap_seconds = max(0.001, self.cfg.tap_ms / 1000.0)

            for t, msgs in groups:
                if self._stop_event.is_set():
                    break

                dt = (t - prev_t) / max(self.cfg.speed, 1e-6)
                if dt > 0:
                    time.sleep(dt)
                prev_t = t

                if self.cfg.tap_mode:
                    # Tap mode: press all keys at this timestamp, then release together
                    keys: List[str] = []
                    for msg in msgs:
                        if msg.is_meta:
                            continue
                        if msg.type == "note_on" and getattr(msg, "velocity", 0) > 0:
                            k = midi_note_to_key(msg.note, self.cfg)
                            if k is not None:
                                keys.append(k)

                    # dedupe (preserve order)
                    seen = set()
                    keys = [k for k in keys if not (k in seen or seen.add(k))]

                    for k in keys:
                        kb.press(k)
                    if keys:
                        time.sleep(tap_seconds)
                        for k in keys:
                            kb.release(k)
                else:
                    # Hold mode: press on note_on, release on note_off (or note_on vel=0)
                    for msg in msgs:
                        if msg.is_meta:
                            continue
                        if msg.type == "note_on" and getattr(msg, "velocity", 0) > 0:
                            k = midi_note_to_key(msg.note, self.cfg)
                            if k is not None:
                                kb.press(k)
                        elif msg.type == "note_off" or (msg.type == "note_on" and getattr(msg, "velocity", 0) == 0):
                            k = midi_note_to_key(msg.note, self.cfg)
                            if k is not None:
                                kb.release(k)

            self._ui(lambda: self._log("Stopped." if self._stop_event.is_set() else "Done."))

        except Exception as e:
            self._ui(lambda: self._log(f"ERROR: {e}"))
        finally:
            self._ui(lambda: self.play_btn.config(state=("normal" if self.cfg.midi_path else "disabled")))
            self._ui(lambda: self.stop_btn.config(state="disabled"))


if __name__ == "__main__":
    App().mainloop()
