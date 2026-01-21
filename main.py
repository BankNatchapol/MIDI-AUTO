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
from Quartz.CoreGraphics import (
    CGEventCreateKeyboardEvent,
    CGEventPost,
    kCGHIDEventTap
)

# macOS character-to-keycode mapping
CHAR_TO_KEYCODE = {
    'a': 0, 's': 1, 'd': 2, 'f': 3, 'h': 4, 'g': 5, 'z': 6, 'x': 7,
    'c': 8, 'v': 9, 'b': 11, 'q': 12, 'w': 13, 'e': 14, 'r': 15,
    't': 17, 'y': 16, 'u': 32, 'i': 34, 'o': 31, 'p': 35,
    'l': 37, 'j': 38, 'k': 40, 'm': 46, 'n': 45,
    '0': 29, '1': 18, '2': 19, '3': 20, '4': 21, '5': 23, '6': 22,
    '7': 26, '8': 28, '9': 25,
    '-': 27, '=': 24, '[': 33, ']': 30,
    ';': 41, "'": 39, ',': 43, '.': 47, '/': 44,
    '+': 24,  # same as '='
}


class KeyController:
    """Quartz-based keyboard controller for macOS."""

    @staticmethod
    def press(char: str) -> None:
        """Press a key down."""
        code = CHAR_TO_KEYCODE.get(char)
        if code is not None:
            event = CGEventCreateKeyboardEvent(None, code, True)
            CGEventPost(kCGHIDEventTap, event)

    @staticmethod
    def release(char: str) -> None:
        """Release a key."""
        code = CHAR_TO_KEYCODE.get(char)
        if code is not None:
            event = CGEventCreateKeyboardEvent(None, code, False)
            CGEventPost(kCGHIDEventTap, event)


kb = KeyController()

APP_DIR = Path.cwd()
MIDIS_DIR = APP_DIR / "midis"
CONFIG_FILE = APP_DIR / "presets.json"

# =========================================================
# 15-key layout (YOUR PROVIDED ORDER)
# =========================================================
# q w e r t y u i
# a s d f g h j
KEYS_15 = [
    'a',  # 0
    's',  # 1
    'd',  # 2
    'f',  # 3
    'g',  # 4
    'h',  # 5
    'j',  # 6
    'q',  # 7
    'w',  # 8
    'e',  # 9
    'r',  # 10
    't',  # 11
    'y',  # 12
    'u',  # 13
    'i',  # 14
]


# Diatonic semitones within an octave for Do..Si in C major: C D E F G A B
DIATONIC_SEMITONES = [0, 2, 4, 5, 7, 9, 11]
SEMITONE_TO_DEGREE = {s: i for i, s in enumerate(DIATONIC_SEMITONES)}


# ===========================
# 21-key keymaps (with black)
# ===========================
def get_keymaps(use_windows: bool):
    # top row
    HIGH = {
        0: 'q',  1: '2',  2: 'w',  3: '3',  4: 'e',  5: 'r',
        6: '5',  7: 't',  8: '6',  9: 'y', 10: '7', 11: 'u',
    }
    HIGH_PLUS_C = 'i'

    # middle row
    MID = {
        0: 'z',  1: 's',  2: 'x',  3: 'd',  4: 'c',  5: 'v',
        6: 'g',  7: 'b',  8: 'h',  9: 'n', 10: 'j', 11: 'm',
    }

    if use_windows:
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

    always_tap: bool = True
    tap_ms: int = 18

    use_windows_map: bool = False
    use_15_keys: bool = False

    # 15-key submode:
    # False = diatonic-only (white notes)
    # True  = chromatic (include half steps)
    chromatic_15: bool = True

    # squeeze feature (15-key only; default OFF)
    squeeze_enabled: bool = False
    squeeze_lo: int = 3
    squeeze_hi: int = 11


def clamp_int(x: int, lo: int, hi: int) -> int:
    return lo if x < lo else hi if x > hi else x


def squeeze_index(idx_0_14: int, lo: int, hi: int) -> int:
    """Map 0..14 into lo..hi monotonically."""
    idx = clamp_int(idx_0_14, 0, 14)
    lo = clamp_int(lo, 0, 14)
    hi = clamp_int(hi, 0, 14)
    if lo > hi:
        lo, hi = hi, lo
    if lo == hi:
        return lo
    return clamp_int(lo + round(idx * (hi - lo) / 14.0), lo, hi)


def _quantize_to_white_floor(semitone: int) -> int:
    """Diatonic-only: map black notes down to nearest lower white note."""
    if semitone in SEMITONE_TO_DEGREE:
        return semitone
    for s in reversed(DIATONIC_SEMITONES):
        if s <= semitone:
            return s
    return 0


def midi_note_to_key(note: int, cfg: Config) -> Optional[str]:
    note += cfg.transpose
    d = note - cfg.base_c_midi
    if d < 0:
        return None

    if cfg.use_15_keys:
        # ---------- 15-key chromatic (includes half-steps) ----------
        if cfg.chromatic_15:
            idx = d  # 1 MIDI note step = 1 semitone
            if not (0 <= idx < 15):
                return None
            if cfg.squeeze_enabled:
                idx = squeeze_index(idx, cfg.squeeze_lo, cfg.squeeze_hi)
            return KEYS_15[idx]

        # ---------- 15-key diatonic-only (white notes) ----------
        octave = d // 12
        semitone = d % 12
        semitone = _quantize_to_white_floor(semitone)
        degree = SEMITONE_TO_DEGREE[semitone]  # 0..6
        idx = octave * 7 + degree
        if not (0 <= idx < 15):
            return None
        if cfg.squeeze_enabled:
            idx = squeeze_index(idx, cfg.squeeze_lo, cfg.squeeze_hi)
        return KEYS_15[idx]

    # ---------- 21-key mode ----------
    LOW, MID, HIGH, HIGH_PLUS_C = get_keymaps(cfg.use_windows_map)
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
    """Iterating MidiFile gives playback order; msg.time is seconds since previous."""
    mid = mido.MidiFile(midi_path)
    out: List[Tuple[float, mido.Message]] = []
    t = 0.0
    for msg in mid:
        t += float(msg.time or 0.0)
        out.append((t, msg))
    return out


def find_trim_window(timed: List[Tuple[float, mido.Message]]) -> Tuple[float, float]:
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
        if msg.type in ("note_off", "note_on"):
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
        self.title("MIDI → Piano Game (21-key / 15-key + chromatic + squeeze)")
        self.geometry("1080x860")

        self.cfg = Config()
        self._stop_event = threading.Event()
        self._play_thread: Optional[threading.Thread] = None

        self.presets: Dict[str, Dict[str, Any]] = self._load_presets()

        self._build_ui()
        self._refresh_presets_dropdown()
        self._update_ui_states()
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
        dest = self._unique_dest(MIDIS_DIR / src_path.name)
        shutil.copy2(src_path, dest)
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
            messagebox.showerror("No MIDI", "Choose MIDI… first.")
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
                "This MIDI isn't in app storage.\nUse Choose MIDI… (it copies into storage), then save again."
            )
            return

        existed = name in self.presets

        # sanitize squeeze bounds
        lo = clamp_int(int(self.squeeze_lo.get()), 0, 14)
        hi = clamp_int(int(self.squeeze_hi.get()), 0, 14)
        if lo > hi:
            lo, hi = hi, lo

        self.presets[name] = {
            "midi_relpath": str(stored_rel),
            "base_c_midi": int(self.base_c.get()),
            "transpose": int(self.transpose.get()),
            "speed": float(self.speed.get()),
            "lead_in": float(self.lead_in.get()),
            "trim_silence": bool(self.trim_silence.get()),
            "always_tap": bool(self.always_tap.get()),
            "tap_ms": int(self.tap_ms.get()),
            "use_windows_map": bool(self.use_windows_map.get()),
            "use_15_keys": bool(self.use_15_keys.get()),
            "chromatic_15": bool(self.chromatic_15.get()),
            "squeeze_enabled": bool(self.squeeze_enabled.get()),
            "squeeze_lo": lo,
            "squeeze_hi": hi,
        }

        self._save_presets()
        self._refresh_presets_dropdown()
        self.preset_var.set(name)
        self._log(f"{'Overwrote' if existed else 'Saved'} preset '{name}'.")

    def apply_preset(self, _event=None) -> None:
        name = self.preset_var.get().strip()
        if not name:
            return
        preset = self.presets.get(name)
        if not isinstance(preset, dict):
            return

        if self._play_thread and self._play_thread.is_alive():
            self._stop_event.set()
            self._log("Stopping playback to switch preset…")

        self.base_c.set(int(preset.get("base_c_midi", 48)))
        self.transpose.set(int(preset.get("transpose", 0)))
        self.speed.set(float(preset.get("speed", 1.0)))
        self.lead_in.set(float(preset.get("lead_in", 2.0)))
        self.trim_silence.set(bool(preset.get("trim_silence", True)))
        self.always_tap.set(bool(preset.get("always_tap", True)))
        self.tap_ms.set(int(preset.get("tap_ms", 18)))
        self.use_windows_map.set(bool(preset.get("use_windows_map", False)))
        self.use_15_keys.set(bool(preset.get("use_15_keys", False)))
        self.chromatic_15.set(bool(preset.get("chromatic_15", True)))

        self.squeeze_enabled.set(bool(preset.get("squeeze_enabled", False)))
        self.squeeze_lo.set(int(preset.get("squeeze_lo", 3)))
        self.squeeze_hi.set(int(preset.get("squeeze_hi", 11)))

        self._update_ui_states()
        self._update_test_button_text()

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
        tk.Label(row, text=label, width=42, anchor="w").pack(side="left")
        tk.Scale(row, variable=var, from_=a, to=b, orient="horizontal", resolution=step).pack(
            side="left", fill="x", expand=True
        )
        tk.Label(row, textvariable=var, width=10, anchor="e").pack(side="right")

    def _update_ui_states(self):
        self.tap_scale.configure(state=("normal" if self.always_tap.get() else "disabled"))

        # 21-key windows map only relevant when NOT 15-key
        self.winmap_chk.configure(state=("disabled" if self.use_15_keys.get() else "normal"))

        # 15-key suboptions
        fifteen = self.use_15_keys.get()
        self.chromatic_chk.configure(state=("normal" if fifteen else "disabled"))

        # squeeze controls only when 15-key + squeeze enabled
        self.squeeze_chk.configure(state=("normal" if fifteen else "disabled"))
        squeeze_on = fifteen and self.squeeze_enabled.get()
        state = "normal" if squeeze_on else "disabled"
        self.squeeze_lo_spin.configure(state=state)
        self.squeeze_hi_spin.configure(state=state)

    def _update_test_button_text(self):
        tmp_cfg = Config(
            base_c_midi=int(self.base_c.get()),
            transpose=int(self.transpose.get()),
            use_windows_map=bool(self.use_windows_map.get()),
            use_15_keys=bool(self.use_15_keys.get()),
            chromatic_15=bool(self.chromatic_15.get()),
            squeeze_enabled=bool(self.squeeze_enabled.get()),
            squeeze_lo=int(self.squeeze_lo.get()),
            squeeze_hi=int(self.squeeze_hi.get()),
        )
        k = midi_note_to_key(tmp_cfg.base_c_midi, tmp_cfg) or "?"
        if self.use_15_keys.get():
            mode = "15-key chromatic" if self.chromatic_15.get() else "15-key diatonic"
            if self.squeeze_enabled.get():
                mode += " + squeeze"
        else:
            mode = "21-key Windows" if self.use_windows_map.get() else "21-key Mac"
        self.test_btn.configure(text=f"Test Base C ({mode}) → '{k}'")

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
        self.preset_combo.bind("<<ComboboxSelected>>", self.apply_preset)

        # Settings
        settings = tk.LabelFrame(frm, text="Playback + Mapping settings", padx=10, pady=10)
        settings.pack(fill="x", pady=(0, 10))

        self.base_c = tk.IntVar(value=self.cfg.base_c_midi)
        self.transpose = tk.IntVar(value=self.cfg.transpose)
        self.speed = tk.DoubleVar(value=self.cfg.speed)
        self.lead_in = tk.DoubleVar(value=self.cfg.lead_in)

        self.trim_silence = tk.BooleanVar(value=True)
        self.always_tap = tk.BooleanVar(value=True)
        self.tap_ms = tk.IntVar(value=18)

        self.use_windows_map = tk.BooleanVar(value=False)
        self.use_15_keys = tk.BooleanVar(value=False)
        self.chromatic_15 = tk.BooleanVar(value=True)   # IMPORTANT: default ON when using 15-key

        # squeeze vars (default OFF)
        self.squeeze_enabled = tk.BooleanVar(value=False)
        self.squeeze_lo = tk.IntVar(value=3)
        self.squeeze_hi = tk.IntVar(value=11)

        self._slider(settings, "Base C MIDI (alignment)", self.base_c, 24, 84, 1)
        self._slider(settings, "Transpose (semitones)", self.transpose, -24, 24, 1)
        self._slider(settings, "Speed", self.speed, 0.25, 3.0, 0.05)
        self._slider(settings, "Lead-in seconds (focus game)", self.lead_in, 0.0, 10.0, 0.25)

        rowA = tk.Frame(settings)
        rowA.pack(fill="x", pady=(8, 0))
        tk.Checkbutton(rowA, text="Trim start/end silence", variable=self.trim_silence).pack(side="left")

        rowB = tk.Frame(settings)
        rowB.pack(fill="x", pady=(6, 0))
        tk.Checkbutton(
            rowB,
            text="Always tap (ignore MIDI note length)",
            variable=self.always_tap,
            command=self._update_ui_states
        ).pack(side="left")

        tap_row = tk.Frame(settings)
        tap_row.pack(fill="x", pady=(4, 0))
        tk.Label(tap_row, text="Tap duration (ms)", width=42, anchor="w").pack(side="left")
        self.tap_scale = tk.Scale(tap_row, variable=self.tap_ms, from_=1, to=80, orient="horizontal", resolution=1)
        self.tap_scale.pack(side="left", fill="x", expand=True)
        tk.Label(tap_row, textvariable=self.tap_ms, width=10, anchor="e").pack(side="right")

        rowC = tk.Frame(settings)
        rowC.pack(fill="x", pady=(10, 0))
        tk.Checkbutton(
            rowC,
            text="15-key mode (q w e r t y u i / a s d f g h j)",
            variable=self.use_15_keys,
            command=lambda: (self._update_ui_states(), self._update_test_button_text())
        ).pack(side="left")

        rowC2 = tk.Frame(settings)
        rowC2.pack(fill="x", pady=(4, 0))
        self.chromatic_chk = tk.Checkbutton(
            rowC2,
            text="15-key chromatic (include half-steps; C, C#, D -> q, w, e)",
            variable=self.chromatic_15,
            command=self._update_test_button_text
        )
        self.chromatic_chk.pack(side="left")

        rowD = tk.Frame(settings)
        rowD.pack(fill="x", pady=(6, 0))
        self.winmap_chk = tk.Checkbutton(
            rowD,
            text="21-key Windows keymap (black keys: 2 3 5 6 7 / S D G H J / . ' 0 - =)",
            variable=self.use_windows_map,
            command=self._update_test_button_text
        )
        self.winmap_chk.pack(side="left")

        # Squeeze UI
        squeeze_box = tk.LabelFrame(settings, text="15-key squeeze (default: OFF / unsqueezed)", padx=10, pady=8)
        squeeze_box.pack(fill="x", pady=(10, 0))

        top_sq = tk.Frame(squeeze_box)
        top_sq.pack(fill="x")
        self.squeeze_chk = tk.Checkbutton(
            top_sq,
            text="Squeeze 15-key range (remap 0..14 into [lo..hi])",
            variable=self.squeeze_enabled,
            command=lambda: (self._update_ui_states(), self._update_test_button_text())
        )
        self.squeeze_chk.pack(side="left")

        band = tk.Frame(squeeze_box)
        band.pack(fill="x", pady=(6, 0))

        tk.Label(band, text="lo (0..14):", width=11, anchor="w").pack(side="left")
        self.squeeze_lo_spin = ttk.Spinbox(
            band, from_=0, to=14, increment=1, textvariable=self.squeeze_lo, width=6,
            command=self._update_test_button_text
        )
        self.squeeze_lo_spin.pack(side="left", padx=(0, 12))

        tk.Label(band, text="hi (0..14):", width=11, anchor="w").pack(side="left")
        self.squeeze_hi_spin = ttk.Spinbox(
            band, from_=0, to=14, increment=1, textvariable=self.squeeze_hi, width=6,
            command=self._update_test_button_text
        )
        self.squeeze_hi_spin.pack(side="left")

        # Buttons
        btns = tk.Frame(frm)
        btns.pack(fill="x", pady=(0, 10))

        self.play_btn = tk.Button(btns, text="▶ Play", command=self.play, state="disabled")
        self.play_btn.pack(side="left")
        self.stop_btn = tk.Button(btns, text="■ Stop", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))
        self.test_btn = tk.Button(btns, text="Test Base C", command=self.test_note)
        self.test_btn.pack(side="right")

        # Log
        log_frame = tk.LabelFrame(frm, text="Log", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, height=12, wrap="word")
        self.log.pack(fill="both", expand=True)

        self._log("15-key chromatic maps semitone steps directly: C,C#,D -> q,w,e.")
        self._log("Squeeze is OFF by default.")
        self._log("If octave feels wrong, adjust Base C MIDI and use Test Base C.")

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
        tmp_cfg = Config(
            base_c_midi=int(self.base_c.get()),
            transpose=int(self.transpose.get()),
            use_windows_map=bool(self.use_windows_map.get()),
            use_15_keys=bool(self.use_15_keys.get()),
            chromatic_15=bool(self.chromatic_15.get()),
            squeeze_enabled=bool(self.squeeze_enabled.get()),
            squeeze_lo=int(self.squeeze_lo.get()),
            squeeze_hi=int(self.squeeze_hi.get()),
        )
        k = midi_note_to_key(tmp_cfg.base_c_midi, tmp_cfg)
        if not k:
            self._log("Test failed: Base C not mapped (check Base C MIDI).")
            return
        kb.press(k)
        time.sleep(0.05)
        kb.release(k)
        self._log(f"Sent test key '{k}'")

    def play(self):
        if not self.cfg.midi_path:
            messagebox.showerror("No MIDI", "Choose MIDI… first.")
            return

        # snapshot settings
        self.cfg.base_c_midi = int(self.base_c.get())
        self.cfg.transpose = int(self.transpose.get())
        self.cfg.speed = float(self.speed.get())
        self.cfg.lead_in = float(self.lead_in.get())
        self.cfg.trim_silence = bool(self.trim_silence.get())
        self.cfg.always_tap = bool(self.always_tap.get())
        self.cfg.tap_ms = int(self.tap_ms.get())
        self.cfg.use_windows_map = bool(self.use_windows_map.get())
        self.cfg.use_15_keys = bool(self.use_15_keys.get())
        self.cfg.chromatic_15 = bool(self.chromatic_15.get())

        self.cfg.squeeze_enabled = bool(self.squeeze_enabled.get())
        lo = clamp_int(int(self.squeeze_lo.get()), 0, 14)
        hi = clamp_int(int(self.squeeze_hi.get()), 0, 14)
        if lo > hi:
            lo, hi = hi, lo
        self.cfg.squeeze_lo = lo
        self.cfg.squeeze_hi = hi

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
            midi_path = self.cfg.midi_path
            if not Path(midi_path).exists():
                self._ui(lambda: self._log(f"ERROR: MIDI missing: {midi_path}"))
                return

            self._ui(lambda: self._log(f"Lead-in {self.cfg.lead_in:.2f}s — focus game window now!"))
            time.sleep(self.cfg.lead_in)

            timed = collect_abs_timed_messages(midi_path)

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

            # optional hold-mode state (only if always_tap OFF)
            key_down: Dict[str, bool] = {}
            MIN_UP = 0.01

            for t, msgs in groups:
                if self._stop_event.is_set():
                    break

                dt = (t - prev_t) / max(self.cfg.speed, 1e-6)
                if dt > 0:
                    time.sleep(dt)
                prev_t = t

                if self.cfg.always_tap:
                    keys: List[str] = []
                    for msg in msgs:
                        if msg.is_meta:
                            continue
                        if msg.type == "note_on" and getattr(msg, "velocity", 0) > 0:
                            k = midi_note_to_key(msg.note, self.cfg)
                            if k is not None:
                                keys.append(k)

                    # dedupe
                    seen = set()
                    keys = [k for k in keys if not (k in seen or seen.add(k))]

                    for k in keys:
                        kb.press(k)
                    if keys:
                        time.sleep(tap_seconds)
                        for k in keys:
                            kb.release(k)
                    
                else:
                    for msg in msgs:
                        if msg.is_meta:
                            continue

                        if msg.type == "note_on" and getattr(msg, "velocity", 0) > 0:
                            k = midi_note_to_key(msg.note, self.cfg)
                            if k is None:
                                continue
                            if key_down.get(k, False):
                                kb.release(k)
                                time.sleep(MIN_UP)
                                key_down[k] = False
                            kb.press(k)
                            key_down[k] = True

                        elif msg.type == "note_off" or (msg.type == "note_on" and getattr(msg, "velocity", 0) == 0):
                            k = midi_note_to_key(msg.note, self.cfg)
                            if k is None:
                                continue
                            if key_down.get(k, False):
                                kb.release(k)
                                key_down[k] = False

            for k, down in list(key_down.items()):
                if down:
                    kb.release(k)

            self._ui(lambda: self._log("Stopped." if self._stop_event.is_set() else "Done."))

        except Exception as e:
            self._ui(lambda: self._log(f"ERROR: {e}"))
        finally:
            self._ui(lambda: self.play_btn.config(state=("normal" if self.cfg.midi_path else "disabled")))
            self._ui(lambda: self.stop_btn.config(state="disabled"))

    def _ui(self, fn):
        self.after(0, fn)


if __name__ == "__main__":
    App().mainloop()
