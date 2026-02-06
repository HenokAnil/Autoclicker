"""Interactive auto clicker with a rich Tkinter UI and global hotkeys.

Dependencies
-----------
- pynput (install with `pip install pynput`)

This tool offers:
- Toggle between left, right, or middle button clicking
- Adjustable click cadence with optional random jitter
- Burst clicking (multiple clicks per cycle)
- Hold mode to keep a button pressed until stopped
- Optional click limit and deferred start countdown
- Global hotkeys to start/stop or trigger an emergency halt
- Live statistics panel (status, elapsed time, total clicks)
- Dedicated keyboard macro panel for automated key presses
- Key combinations with tap, double tap, or hold behaviour
- Macro tab to record combined mouse and keyboard sequences

Run the script directly to launch the UI.
"""
from __future__ import annotations

import random
import threading
import time
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple, Union

import tkinter as tk
from tkinter import messagebox, ttk

from pynput import keyboard, mouse


@dataclass
class HotkeyBinding:
    """Container for a user-configured hotkey."""

    label: str
    tokens: Set[str]
    display_text: str

    def matches(self, pressed: Set[str]) -> bool:
        return bool(self.tokens) and self.tokens <= pressed


@dataclass
class ClickSettings:
    """Immutable snapshot of click parameters for a run."""

    button: mouse.Button
    mode: str
    base_interval: float
    jitter: float
    burst_count: int
    start_delay: float
    limit_enabled: bool
    limit_count: int


@dataclass
class KeyPressSettings:
    """Immutable snapshot of keyboard automation parameters."""

    sequence: Tuple[Union[keyboard.Key, str], ...]
    mode: str
    base_interval: float
    jitter: float
    burst_count: int
    start_delay: float
    limit_enabled: bool
    limit_count: int


@dataclass
class MacroEvent:
    """Single recorded action for the Macro replay system."""

    kind: str
    action: str
    delay: float
    key_value: Optional[Union[keyboard.Key, keyboard.KeyCode, str]] = None
    mouse_button: Optional[mouse.Button] = None
    position: Optional[Tuple[int, int]] = None


class AutoClickerApp:
    CANONICAL_KEYS: Dict[str, str] = {
        "shift_l": "shift",
        "shift_r": "shift",
        "ctrl_l": "ctrl",
        "ctrl_r": "ctrl",
        "alt_l": "alt",
        "alt_r": "alt",
        "cmd_l": "cmd",
        "cmd_r": "cmd",
    }

    FRIENDLY_NAMES: Dict[str, str] = {
        "ctrl": "Ctrl",
        "shift": "Shift",
        "alt": "Alt",
        "cmd": "Cmd",
        "enter": "Enter",
        "space": "Space",
        "tab": "Tab",
        "esc": "Esc",
    }

    BUTTON_MAP: Dict[str, mouse.Button] = {
        "left": mouse.Button.left,
        "right": mouse.Button.right,
        "middle": mouse.Button.middle,
    }

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("AutoClicker By H")
        self.root.geometry("760x480")
        self.root.resizable(False, False)

        self.mouse_controller = mouse.Controller()
        self.key_controller = keyboard.Controller()
        self.keyboard_listener: Optional[keyboard.Listener] = None
        self.mouse_listener: Optional[mouse.Listener] = None

        self.pressed_tokens: Set[str] = set()
        self.trigger_latch: Set[str] = set()
        self.recording_target: Optional[str] = None
        self.recording_buffer: Set[str] = set()

        self.click_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.is_running = False
        self.start_timestamp: Optional[float] = None
        self.last_elapsed_seconds = 0
        self.clicks_recorded = 0
        self.run_reason: str = "Idle"
        self.hold_pressed = False

        self.key_thread: Optional[threading.Thread] = None
        self.key_stop_event = threading.Event()
        self.key_is_running = False
        self.key_start_timestamp: Optional[float] = None
        self.key_last_elapsed_seconds = 0
        self.key_cycles_recorded = 0
        self.key_run_reason: str = "Idle"
        self.key_hold_active = False
        self.active_key_sequence: Optional[Tuple[Union[keyboard.Key, str], ...]] = None

        self.macro_events: List[MacroEvent] = []
        self.macro_recording = False
        self.macro_start_time: Optional[float] = None
        self.macro_last_timestamp: Optional[float] = None
        self.macro_play_thread: Optional[threading.Thread] = None
        self.macro_stop_event = threading.Event()
        self.macro_run_reason: str = "Idle"
        self.block_macro_capture = False

        self.macro_status_var = tk.StringVar(value="Idle")
        self.macro_event_count_var = tk.StringVar(value="Events: 0")
        self.macro_loop_var = tk.IntVar(value=1)
        self.macro_play_delay_var = tk.DoubleVar(value=0.0)
        self.macro_listbox: Optional[tk.Listbox] = None

        self.button_var = tk.StringVar(value="left")
        self.mode_var = tk.StringVar(value="single")
        self.delay_var = tk.DoubleVar(value=100.0)
        self.jitter_var = tk.DoubleVar(value=0.0)
        self.burst_var = tk.IntVar(value=1)
        self.start_delay_var = tk.DoubleVar(value=0.0)
        self.limit_enabled_var = tk.BooleanVar(value=False)
        self.limit_count_var = tk.IntVar(value=500)

        self.status_var = tk.StringVar(value="Idle")
        self.hotkey_vars: Dict[str, tk.StringVar] = {}

        self.key_sequence_var = tk.StringVar(value="space")
        self.key_mode_var = tk.StringVar(value="tap")
        self.key_delay_var = tk.DoubleVar(value=150.0)
        self.key_jitter_var = tk.DoubleVar(value=0.0)
        self.key_burst_var = tk.IntVar(value=1)
        self.key_start_delay_var = tk.DoubleVar(value=0.0)
        self.key_limit_enabled_var = tk.BooleanVar(value=False)
        self.key_limit_count_var = tk.IntVar(value=500)

        self.key_status_var = tk.StringVar(value="Idle")
        self.key_elapsed_var = tk.StringVar(value="Elapsed: 00:00:00")
        self.key_cycle_var = tk.StringVar(value="Cycles: 0")

        self.hotkey_order = ("mouse_toggle", "key_toggle", "panic")
        self.hotkeys: Dict[str, HotkeyBinding] = {
            "mouse_toggle": HotkeyBinding(label="Toggle Mouse Clicker", tokens={"f6"}, display_text="F6"),
            "key_toggle": HotkeyBinding(label="Toggle Key Macro", tokens={"f8"}, display_text="F8"),
            "panic": HotkeyBinding(label="Emergency Halt", tokens={"f7"}, display_text="F7"),
        }

        self._build_ui()
        self._bind_hotkeys_to_ui()
        self._start_keyboard_listener()
        self._start_mouse_listener()
        self._schedule_status_refresh()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _safe_call_ui(self, func, *args, **kwargs) -> None:
        if threading.current_thread() is threading.main_thread():
            func(*args, **kwargs)
        else:
            self.root.after(0, lambda: func(*args, **kwargs))

    # ------------------------------------------------------------------
    # UI construction helpers
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        main_frame = ttk.Frame(notebook)
        notebook.add(main_frame, text="Click Control")
        self._build_main_tab(main_frame)

        key_frame = ttk.Frame(notebook)
        notebook.add(key_frame, text="Key Control")
        self._build_key_tab(key_frame)

        macro_frame = ttk.Frame(notebook)
        notebook.add(macro_frame, text="Macro")
        self._build_macro_tab(macro_frame)

        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="Settings")
        self._build_settings_tab(settings_frame)

    def _build_main_tab(self, container: ttk.Frame) -> None:
        click_frame = ttk.LabelFrame(container, text="Click Parameters")
        click_frame.pack(fill=tk.X, padx=8, pady=8)

        ttk.Label(click_frame, text="Mouse Button:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=4)
        for idx, (label, value) in enumerate((("Left", "left"), ("Right", "right"), ("Middle", "middle"))):
            ttk.Radiobutton(
                click_frame,
                text=label,
                value=value,
                variable=self.button_var,
            ).grid(row=0, column=idx + 1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(click_frame, text="Click Mode:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Combobox(
            click_frame,
            textvariable=self.mode_var,
            values=("single", "double", "hold"),
            state="readonly",
            width=10,
        ).grid(row=1, column=1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(click_frame, text="Delay per cycle (ms):").grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            click_frame,
            from_=1,
            to=100000,
            increment=10,
            textvariable=self.delay_var,
            width=12,
        ).grid(row=2, column=1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(click_frame, text="Random jitter ± (ms):").grid(row=2, column=2, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            click_frame,
            from_=0,
            to=5000,
            increment=5,
            textvariable=self.jitter_var,
            width=10,
        ).grid(row=2, column=3, sticky=tk.W, padx=4, pady=4)

        ttk.Label(click_frame, text="Clicks per cycle:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            click_frame,
            from_=1,
            to=50,
            increment=1,
            textvariable=self.burst_var,
            width=10,
        ).grid(row=3, column=1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(click_frame, text="Start delay (s):").grid(row=3, column=2, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            click_frame,
            from_=0,
            to=30,
            increment=0.5,
            textvariable=self.start_delay_var,
            width=10,
        ).grid(row=3, column=3, sticky=tk.W, padx=4, pady=4)

        limit_frame = ttk.Frame(click_frame)
        limit_frame.grid(row=4, column=0, columnspan=4, sticky=tk.W, padx=4, pady=4)
        ttk.Checkbutton(
            limit_frame,
            text="Stop after",
            variable=self.limit_enabled_var,
        ).pack(side=tk.LEFT)
        ttk.Spinbox(
            limit_frame,
            from_=1,
            to=1_000_000,
            increment=10,
            textvariable=self.limit_count_var,
            width=12,
        ).pack(side=tk.LEFT, padx=4)
        ttk.Label(limit_frame, text="clicks").pack(side=tk.LEFT)

        action_frame = ttk.Frame(container)
        action_frame.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(action_frame, text="Start", command=self.start_clicking).pack(side=tk.LEFT, padx=4)
        ttk.Button(action_frame, text="Stop", command=self.stop_clicking).pack(side=tk.LEFT, padx=4)
        ttk.Button(action_frame, text="Reset Stats", command=self._reset_mouse_stats).pack(side=tk.LEFT, padx=4)

        stats_frame = ttk.LabelFrame(container, text="Session Stats")
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        ttk.Label(stats_frame, textvariable=self.status_var, font=("Segoe UI", 11)).pack(anchor=tk.W, padx=8, pady=6)

        self.elapsed_var = tk.StringVar(value="Elapsed: 00:00:00")
        self.click_count_var = tk.StringVar(value="Clicks: 0")
        ttk.Label(stats_frame, textvariable=self.elapsed_var).pack(anchor=tk.W, padx=8)
        ttk.Label(stats_frame, textvariable=self.click_count_var).pack(anchor=tk.W, padx=8)

    def _build_key_tab(self, container: ttk.Frame) -> None:
        key_frame = ttk.LabelFrame(container, text="Key Parameters")
        key_frame.pack(fill=tk.X, padx=8, pady=8)

        ttk.Label(key_frame, text="Combination (use +):").grid(row=0, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Entry(key_frame, textvariable=self.key_sequence_var, width=24).grid(row=0, column=1, sticky=tk.W, padx=4, pady=4)
        ttk.Label(key_frame, text="Examples: space, enter, ctrl+shift+p").grid(row=0, column=2, columnspan=2, sticky=tk.W, padx=4, pady=4)

        ttk.Label(key_frame, text="Key Mode:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Combobox(
            key_frame,
            textvariable=self.key_mode_var,
            values=("tap", "double", "hold"),
            state="readonly",
            width=10,
        ).grid(row=1, column=1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(key_frame, text="Delay per cycle (ms):").grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            key_frame,
            from_=10,
            to=120000,
            increment=10,
            textvariable=self.key_delay_var,
            width=12,
        ).grid(row=2, column=1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(key_frame, text="Random jitter ± (ms):").grid(row=2, column=2, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            key_frame,
            from_=0,
            to=5000,
            increment=5,
            textvariable=self.key_jitter_var,
            width=10,
        ).grid(row=2, column=3, sticky=tk.W, padx=4, pady=4)

        ttk.Label(key_frame, text="Cycles per batch:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            key_frame,
            from_=1,
            to=50,
            increment=1,
            textvariable=self.key_burst_var,
            width=10,
        ).grid(row=3, column=1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(key_frame, text="Start delay (s):").grid(row=3, column=2, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            key_frame,
            from_=0,
            to=30,
            increment=0.5,
            textvariable=self.key_start_delay_var,
            width=10,
        ).grid(row=3, column=3, sticky=tk.W, padx=4, pady=4)

        key_limit_frame = ttk.Frame(key_frame)
        key_limit_frame.grid(row=4, column=0, columnspan=4, sticky=tk.W, padx=4, pady=4)
        ttk.Checkbutton(
            key_limit_frame,
            text="Stop after",
            variable=self.key_limit_enabled_var,
        ).pack(side=tk.LEFT)
        ttk.Spinbox(
            key_limit_frame,
            from_=1,
            to=1_000_000,
            increment=10,
            textvariable=self.key_limit_count_var,
            width=12,
        ).pack(side=tk.LEFT, padx=4)
        ttk.Label(key_limit_frame, text="cycles").pack(side=tk.LEFT)

        key_action_frame = ttk.Frame(container)
        key_action_frame.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(key_action_frame, text="Start Macro", command=self.start_key_pressing).pack(side=tk.LEFT, padx=4)
        ttk.Button(key_action_frame, text="Stop Macro", command=self.stop_key_pressing).pack(side=tk.LEFT, padx=4)
        ttk.Button(key_action_frame, text="Reset Stats", command=self._reset_key_stats).pack(side=tk.LEFT, padx=4)

        key_stats_frame = ttk.LabelFrame(container, text="Macro Stats")
        key_stats_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        ttk.Label(key_stats_frame, textvariable=self.key_status_var, font=("Segoe UI", 11)).pack(anchor=tk.W, padx=8, pady=6)
        ttk.Label(key_stats_frame, textvariable=self.key_elapsed_var).pack(anchor=tk.W, padx=8)
        ttk.Label(key_stats_frame, textvariable=self.key_cycle_var).pack(anchor=tk.W, padx=8)

    def _build_macro_tab(self, container: ttk.Frame) -> None:
        control_frame = ttk.LabelFrame(container, text="Recorder Controls")
        control_frame.pack(fill=tk.X, padx=8, pady=8)

        ttk.Button(control_frame, text="Start Recording", command=self.start_macro_recording).grid(row=0, column=0, padx=4, pady=4, sticky=tk.W)
        ttk.Button(control_frame, text="Stop Recording", command=self.stop_macro_recording).grid(row=0, column=1, padx=4, pady=4, sticky=tk.W)
        ttk.Button(control_frame, text="Clear", command=self.clear_macro_events).grid(row=0, column=2, padx=4, pady=4, sticky=tk.W)

        ttk.Label(control_frame, text="Loops:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            control_frame,
            from_=1,
            to=1000,
            increment=1,
            textvariable=self.macro_loop_var,
            width=8,
        ).grid(row=1, column=1, sticky=tk.W, padx=4, pady=4)

        ttk.Label(control_frame, text="Start delay (s):").grid(row=1, column=2, sticky=tk.W, padx=4, pady=4)
        ttk.Spinbox(
            control_frame,
            from_=0,
            to=60,
            increment=0.5,
            textvariable=self.macro_play_delay_var,
            width=8,
        ).grid(row=1, column=3, sticky=tk.W, padx=4, pady=4)

        ttk.Button(control_frame, text="Play", command=self.play_macro).grid(row=2, column=0, padx=4, pady=4, sticky=tk.W)
        ttk.Button(control_frame, text="Stop Playback", command=self.stop_macro_playback).grid(row=2, column=1, padx=4, pady=4, sticky=tk.W)

        ttk.Label(control_frame, text="Recorder stores global keyboard presses and mouse clicks with timing. Use hotkeys to stop if needed.", wraplength=520, justify=tk.LEFT).grid(row=3, column=0, columnspan=4, sticky=tk.W, padx=4, pady=(8, 0))

        events_frame = ttk.LabelFrame(container, text="Recorded Events")
        events_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.macro_listbox = tk.Listbox(events_frame, height=12, activestyle="none")
        self.macro_listbox.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        status_frame = ttk.LabelFrame(container, text="Macro Status")
        status_frame.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(status_frame, textvariable=self.macro_status_var, font=("Segoe UI", 11)).pack(anchor=tk.W, padx=8, pady=4)
        ttk.Label(status_frame, textvariable=self.macro_event_count_var).pack(anchor=tk.W, padx=8, pady=(0, 4))

    def _build_settings_tab(self, container: ttk.Frame) -> None:
        hotkey_frame = ttk.LabelFrame(container, text="Global Hotkeys")
        hotkey_frame.pack(fill=tk.X, padx=8, pady=8)

        for row_index, hotkey_id in enumerate(self.hotkey_order):
            binding = self.hotkeys[hotkey_id]
            ttk.Label(hotkey_frame, text=f"{binding.label}:").grid(row=row_index, column=0, sticky=tk.W, padx=4, pady=6)

            var = tk.StringVar(value=binding.display_text)
            self.hotkey_vars[hotkey_id] = var
            entry = ttk.Entry(hotkey_frame, textvariable=var, state="readonly", width=24)
            entry.grid(row=row_index, column=1, sticky=tk.W, padx=4, pady=6)

            ttk.Button(
                hotkey_frame,
                text="Set",
                command=lambda hid=hotkey_id: self._begin_hotkey_recording(hid),
            ).grid(row=row_index, column=2, sticky=tk.W, padx=4, pady=6)

        info_frame = ttk.LabelFrame(container, text="Usage Tips")
        info_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        tips = (
            "F6 toggles the clicker by default. Configure your own keys above.",
            "F7 performs an immediate emergency stop.",
            "Hold mode keeps the button pressed; use the panic hotkey to release quickly if needed.",
            "Random jitter helps mimic human input for anti-detection scenarios.",
            "Use the start delay to give yourself time to move the pointer before the run begins.",
            "Switch to Key Control for automated key combinations, including modifiers and function keys.",
            "Use the Macro tab to record and replay complete mouse and keyboard sequences.",
        )
        for tip in tips:
            ttk.Label(info_frame, text=f"• {tip}", wraplength=680, justify=tk.LEFT).pack(anchor=tk.W, padx=8, pady=2)

    def _bind_hotkeys_to_ui(self) -> None:
        for key, binding in self.hotkeys.items():
            var = self.hotkey_vars.get(key)
            if var is not None:
                var.set(binding.display_text)

    # ------------------------------------------------------------------
    # Hotkey handling
    # ------------------------------------------------------------------
    def _start_keyboard_listener(self) -> None:
        if self.keyboard_listener is not None:
            return
        listener = keyboard.Listener(on_press=self._on_key_press, on_release=self._on_key_release)
        listener.start()
        self.keyboard_listener = listener

    def _start_mouse_listener(self) -> None:
        if self.mouse_listener is not None:
            return
        listener = mouse.Listener(on_click=self._on_mouse_click)
        listener.start()
        self.mouse_listener = listener

    def _tokenize_key(self, key: keyboard.Key | keyboard.KeyCode) -> Optional[str]:
        try:
            if isinstance(key, keyboard.KeyCode):
                if key.char:
                    return key.char.lower()
                if key.vk:
                    return f"vk_{key.vk}"
                return None
            name = key.name if hasattr(key, "name") else str(key)
            if not name:
                return None
            return self.CANONICAL_KEYS.get(name, name)
        except Exception:
            return None

    def _format_tokens(self, tokens: Iterable[str]) -> str:
        ordered = []
        modifier_priority = ["ctrl", "shift", "alt", "cmd"]
        modifiers = [tok for tok in tokens if tok in modifier_priority]
        others = [tok for tok in tokens if tok not in modifier_priority]
        modifiers.sort(key=lambda tok: modifier_priority.index(tok))
        ordered.extend(modifiers)
        ordered.extend(sorted(others))

        friendly = []
        for tok in ordered:
            if tok in self.FRIENDLY_NAMES:
                friendly.append(self.FRIENDLY_NAMES[tok])
            elif tok.startswith("f") and tok[1:].isdigit():
                friendly.append(tok.upper())
            elif tok.startswith("vk_"):
                friendly.append(tok.replace("vk_", "VK"))
            elif len(tok) == 1:
                friendly.append(tok.upper())
            else:
                friendly.append(tok.capitalize())
        return " + ".join(friendly) if friendly else "Unset"

    def _on_key_press(self, key: keyboard.Key | keyboard.KeyCode) -> None:
        token = self._tokenize_key(key)
        if not token:
            return
        self.root.after(0, lambda tok=token, raw=key: self._handle_key_press(tok, raw))

    def _on_key_release(self, key: keyboard.Key | keyboard.KeyCode) -> None:
        token = self._tokenize_key(key)
        if token:
            self.root.after(0, lambda tok=token, raw=key: self._handle_key_release(tok, raw))

    def _on_mouse_click(self, x: int, y: int, button: mouse.Button, pressed: bool) -> None:
        self.root.after(0, lambda: self._handle_mouse_click(int(x), int(y), button, pressed))

    def _handle_key_press(self, token: str, raw_key: keyboard.Key | keyboard.KeyCode | None = None) -> None:
        self.pressed_tokens.add(token)
        if raw_key is not None:
            self._maybe_record_macro_key_event(raw_key, pressed=True)
        if self.recording_target:
            self.recording_buffer.add(token)
            if self.recording_buffer:
                display = self._format_tokens(self.recording_buffer)
                var = self.hotkey_vars.get(self.recording_target)
                if var is not None:
                    var.set(display)
            return

        self._process_hotkey_triggers()

    def _handle_key_release(self, token: str, raw_key: keyboard.Key | keyboard.KeyCode | None = None) -> None:
        if token in self.pressed_tokens:
            self.pressed_tokens.discard(token)

        if raw_key is not None:
            self._maybe_record_macro_key_event(raw_key, pressed=False)

        if self.recording_target and not self.pressed_tokens:
            if not self.recording_buffer:
                var = self.hotkey_vars.get(self.recording_target)
                if var is not None:
                    var.set("Unset")
            else:
                self._update_hotkey_binding(self.recording_target, set(self.recording_buffer))
            self.recording_target = None
            self.recording_buffer.clear()
            self.status_var.set(self.run_reason)
            return

        self._release_latched_hotkeys()

    def _process_hotkey_triggers(self) -> None:
        pressed = set(self.pressed_tokens)
        for hotkey_id, binding in self.hotkeys.items():
            if binding.matches(pressed) and hotkey_id not in self.trigger_latch:
                self.trigger_latch.add(hotkey_id)
                if hotkey_id == "mouse_toggle":
                    self.toggle_clicking()
                elif hotkey_id == "key_toggle":
                    self.toggle_key_pressing()
                elif hotkey_id == "panic":
                    self._panic_stop()

    def _release_latched_hotkeys(self) -> None:
        to_release = {hid for hid in self.trigger_latch if not self.hotkeys[hid].matches(self.pressed_tokens)}
        self.trigger_latch.difference_update(to_release)

    def _begin_hotkey_recording(self, hotkey_id: str) -> None:
        if self.recording_target:
            return
        self.recording_target = hotkey_id
        self.recording_buffer.clear()
        self.status_var.set(f"Press desired keys for {self.hotkeys[hotkey_id].label}")

    def _update_hotkey_binding(self, hotkey_id: str, tokens: Set[str]) -> None:
        if not tokens:
            messagebox.showwarning("Hotkey", "Please press at least one key for the hotkey.")
            return
        collision = [hid for hid, binding in self.hotkeys.items() if hid != hotkey_id and binding.tokens == tokens]
        if collision:
            name = self.hotkeys[collision[0]].label
            messagebox.showwarning("Hotkey", f"This key combination is already used by {name}.")
            return
        display = self._format_tokens(tokens)
        self.hotkeys[hotkey_id] = HotkeyBinding(
            label=self.hotkeys[hotkey_id].label,
            tokens=set(tokens),
            display_text=display,
        )
        self.hotkey_vars[hotkey_id].set(display)
        self.status_var.set(self.run_reason)

    # ------------------------------------------------------------------
    # Macro recording helpers
    # ------------------------------------------------------------------
    def _handle_mouse_click(self, x: int, y: int, button: mouse.Button, pressed: bool) -> None:
        if not self.macro_recording or self.block_macro_capture:
            return
        self._record_macro_event(
            kind="mouse",
            action="press" if pressed else "release",
            mouse_button=button,
            position=(x, y),
        )

    def _maybe_record_macro_key_event(self, raw_key: keyboard.Key | keyboard.KeyCode, pressed: bool) -> None:
        if not self.macro_recording or self.block_macro_capture or self.recording_target:
            return
        normalized = self._normalize_macro_key(raw_key)
        self._record_macro_event(
            kind="key",
            action="press" if pressed else "release",
            key_value=normalized,
        )

    def _normalize_macro_key(self, raw_key: keyboard.Key | keyboard.KeyCode | str) -> Union[keyboard.Key, keyboard.KeyCode, str]:
        if isinstance(raw_key, keyboard.KeyCode):
            if raw_key.char:
                return raw_key.char
            if raw_key.vk is not None:
                return keyboard.KeyCode(vk=raw_key.vk)
        return raw_key

    def _record_macro_event(
        self,
        *,
        kind: str,
        action: str,
        key_value: Optional[Union[keyboard.Key, keyboard.KeyCode, str]] = None,
        mouse_button: Optional[mouse.Button] = None,
        position: Optional[Tuple[int, int]] = None,
    ) -> None:
        if not self.macro_recording or self.block_macro_capture:
            return
        now = time.perf_counter()
        if self.macro_start_time is None:
            self.macro_start_time = now
        delay = 0.0
        if self.macro_last_timestamp is None:
            delay = now - self.macro_start_time
        else:
            delay = now - self.macro_last_timestamp
        self.macro_last_timestamp = now

        event = MacroEvent(
            kind=kind,
            action=action,
            delay=max(0.0, delay),
            key_value=key_value,
            mouse_button=mouse_button,
            position=position,
        )
        self.macro_events.append(event)
        self.macro_event_count_var.set(f"Events: {len(self.macro_events)}")
        self.macro_status_var.set("Recording")
        if self.macro_listbox is not None:
            self.macro_listbox.insert(tk.END, self._format_macro_event(event))
            self.macro_listbox.see(tk.END)

    def _format_macro_event(self, event: MacroEvent) -> str:
        delay_ms = event.delay * 1000.0
        if event.kind == "key":
            key_text = self._describe_macro_key(event.key_value)
            return f"[{delay_ms:7.1f} ms] Key {event.action}: {key_text}"
        button_text = self._describe_macro_button(event.mouse_button)
        position_text = f"({event.position[0]}, {event.position[1]})" if event.position else "(?, ?)"
        return f"[{delay_ms:7.1f} ms] Mouse {event.action}: {button_text} @ {position_text}"

    def _describe_macro_key(self, key_value: Union[keyboard.Key, keyboard.KeyCode, str, None]) -> str:
        if key_value is None:
            return "Unknown"
        if isinstance(key_value, str):
            return key_value.upper()
        if isinstance(key_value, keyboard.Key):
            return getattr(key_value, "name", str(key_value))
        if isinstance(key_value, keyboard.KeyCode):
            if key_value.char:
                return key_value.char.upper()
            return f"VK{key_value.vk}" if key_value.vk is not None else "KeyCode"
        return str(key_value)

    def _describe_macro_button(self, button: Optional[mouse.Button]) -> str:
        if button is None:
            return "Unknown"
        if button == mouse.Button.left:
            return "Left"
        if button == mouse.Button.right:
            return "Right"
        if button == mouse.Button.middle:
            return "Middle"
        return str(button)

    # ------------------------------------------------------------------
    # Click handling
    # ------------------------------------------------------------------
    def start_clicking(self) -> None:
        if self.is_running:
            return
        try:
            settings = self._collect_settings()
        except ValueError as exc:
            messagebox.showerror("Invalid settings", str(exc))
            return

        self.stop_event.clear()
        self.is_running = True
        self.start_timestamp = time.time()
        self.last_elapsed_seconds = 0
        self.clicks_recorded = 0
        self.run_reason = "Running"
        self.hold_pressed = False
        self.status_var.set("Running")

        self.click_thread = threading.Thread(target=self._click_worker, args=(settings,), daemon=True)
        self.click_thread.start()

    def stop_clicking(self) -> None:
        if not self.is_running:
            return
        self.run_reason = "Stopped"
        self.stop_event.set()
        if self.click_thread and self.click_thread.is_alive():
            self.click_thread.join(timeout=2.0)
        self._finalize_mouse_stop()

    def toggle_clicking(self) -> None:
        if self.is_running:
            self.stop_clicking()
        else:
            self.start_clicking()

    def _panic_stop(self) -> None:
        self.run_reason = "Emergency stop"
        self.key_run_reason = "Emergency stop"
        self.macro_run_reason = "Emergency stop"
        self.stop_event.set()
        self.key_stop_event.set()
        self.macro_stop_event.set()
        if self.is_running and self.click_thread and self.click_thread.is_alive():
            self.click_thread.join(timeout=2.0)
        if self.key_is_running and self.key_thread and self.key_thread.is_alive():
            self.key_thread.join(timeout=2.0)
        if self.macro_play_thread and self.macro_play_thread.is_alive():
            self.macro_play_thread.join(timeout=2.0)
        self._finalize_mouse_stop()
        self._finalize_key_stop()
        self._finalize_macro_stop()

    def _click_worker(self, settings: ClickSettings) -> None:
        try:
            if settings.start_delay > 0:
                start_deadline = time.time() + settings.start_delay
                while time.time() < start_deadline and not self.stop_event.is_set():
                    remaining = max(0.0, start_deadline - time.time())
                    self._safe_call_ui(self.status_var.set, f"Starting in {remaining:.1f}s")
                    time.sleep(0.1)

            self._safe_call_ui(self.status_var.set, "Running")
            while not self.stop_event.is_set():
                if settings.limit_enabled and self.clicks_recorded >= settings.limit_count:
                    self.run_reason = "Limit reached"
                    break

                if settings.mode == "hold":
                    if not self.hold_pressed:
                        self.mouse_controller.press(settings.button)
                        self.hold_pressed = True
                    if self.stop_event.wait(0.05):
                        break
                    continue

                per_click_sleep = 0.01
                for _ in range(settings.burst_count):
                    if self.stop_event.is_set():
                        break
                    click_count = 2 if settings.mode == "double" else 1
                    self.mouse_controller.click(settings.button, click_count)
                    self.clicks_recorded += click_count
                    if settings.burst_count > 1 and not self.stop_event.is_set():
                        self.stop_event.wait(per_click_sleep)

                next_delay = settings.base_interval
                if settings.jitter > 0:
                    next_delay += random.uniform(-settings.jitter, settings.jitter)
                    next_delay = max(0.005, next_delay)
                if self.stop_event.wait(next_delay):
                    break
        except Exception as exc:
            self.run_reason = f"Error: {exc}"
        finally:
            if self.hold_pressed:
                try:
                    self.mouse_controller.release(settings.button)
                except Exception:
                    pass
                self.hold_pressed = False
            self.root.after(0, self._finalize_mouse_stop)

    # ------------------------------------------------------------------
    # Key macro handling
    # ------------------------------------------------------------------
    def start_key_pressing(self) -> None:
        if self.key_is_running:
            return
        try:
            settings = self._collect_key_settings()
        except ValueError as exc:
            messagebox.showerror("Invalid settings", str(exc))
            return

        self.key_stop_event.clear()
        self.key_is_running = True
        self.key_start_timestamp = time.time()
        self.key_last_elapsed_seconds = 0
        self.key_cycles_recorded = 0
        self.key_run_reason = "Running"
        self.key_hold_active = False
        self.active_key_sequence = settings.sequence
        self.key_status_var.set("Running")

        self.key_thread = threading.Thread(target=self._key_worker, args=(settings,), daemon=True)
        self.key_thread.start()

    def stop_key_pressing(self) -> None:
        if not self.key_is_running:
            return
        self.key_run_reason = "Stopped"
        self.key_stop_event.set()
        if self.key_thread and self.key_thread.is_alive():
            self.key_thread.join(timeout=2.0)
        self._finalize_key_stop()

    def toggle_key_pressing(self) -> None:
        if self.key_is_running:
            self.stop_key_pressing()
        else:
            self.start_key_pressing()

    def _key_worker(self, settings: KeyPressSettings) -> None:
        try:
            if settings.start_delay > 0:
                start_deadline = time.time() + settings.start_delay
                while time.time() < start_deadline and not self.key_stop_event.is_set():
                    remaining = max(0.0, start_deadline - time.time())
                    self._safe_call_ui(self.key_status_var.set, f"Starting in {remaining:.1f}s")
                    time.sleep(0.1)

            self._safe_call_ui(self.key_status_var.set, "Running")
            while not self.key_stop_event.is_set():
                if settings.limit_enabled and self.key_cycles_recorded >= settings.limit_count:
                    self.key_run_reason = "Limit reached"
                    break

                if settings.mode == "hold":
                    if not self.key_hold_active:
                        self._press_sequence(settings.sequence)
                        self.key_hold_active = True
                        if self.key_cycles_recorded == 0:
                            self.key_cycles_recorded = 1
                    if self.key_stop_event.wait(0.05):
                        break
                    continue

                per_cycle_sleep = 0.02
                for _ in range(settings.burst_count):
                    if self.key_stop_event.is_set():
                        break
                    repeats = 2 if settings.mode == "double" else 1
                    self._tap_sequence(settings.sequence, repeats=repeats)
                    self.key_cycles_recorded += 1
                    if settings.burst_count > 1 and not self.key_stop_event.is_set():
                        self.key_stop_event.wait(per_cycle_sleep)

                next_delay = settings.base_interval
                if settings.jitter > 0:
                    next_delay += random.uniform(-settings.jitter, settings.jitter)
                    next_delay = max(0.01, next_delay)
                if self.key_stop_event.wait(next_delay):
                    break
        except Exception as exc:
            self.key_run_reason = f"Error: {exc}"
        finally:
            if self.key_hold_active:
                try:
                    self._release_sequence(settings.sequence)
                except Exception:
                    pass
                self.key_hold_active = False
            self.root.after(0, self._finalize_key_stop)

    # ------------------------------------------------------------------
    # Macro handling
    # ------------------------------------------------------------------
    def start_macro_recording(self) -> None:
        if self.macro_play_thread and self.macro_play_thread.is_alive():
            messagebox.showinfo("Macro", "Stop playback before recording a new sequence.")
            return
        self.macro_events.clear()
        self.macro_recording = True
        self.block_macro_capture = False
        self.macro_start_time = time.perf_counter()
        self.macro_last_timestamp = None
        self.macro_event_count_var.set("Events: 0")
        self.macro_run_reason = "Recording"
        self.macro_status_var.set("Recording")
        if self.macro_listbox is not None:
            self.macro_listbox.delete(0, tk.END)

    def stop_macro_recording(self) -> None:
        if not self.macro_recording:
            return
        self.macro_recording = False
        self.macro_last_timestamp = None
        self.macro_start_time = None
        total = len(self.macro_events)
        if total == 0:
            self.macro_status_var.set("Idle")
            self.macro_run_reason = "Idle"
        else:
            self.macro_status_var.set(f"Recorded {total} events")
            self.macro_run_reason = "Ready"

    def clear_macro_events(self) -> None:
        if self.macro_recording:
            messagebox.showinfo("Macro", "Stop recording before clearing events.")
            return
        if self.macro_play_thread and self.macro_play_thread.is_alive():
            messagebox.showinfo("Macro", "Stop playback before clearing events.")
            return
        self.macro_events.clear()
        self.macro_event_count_var.set("Events: 0")
        self.macro_status_var.set("Idle")
        self.macro_run_reason = "Idle"
        self.macro_start_time = None
        self.macro_last_timestamp = None
        if self.macro_listbox is not None:
            self.macro_listbox.delete(0, tk.END)

    def play_macro(self) -> None:
        if self.macro_recording:
            messagebox.showinfo("Macro", "Stop recording before starting playback.")
            return
        if not self.macro_events:
            messagebox.showinfo("Macro", "Record at least one event before playback.")
            return
        try:
            loops = int(self.macro_loop_var.get())
        except (TypeError, ValueError):
            messagebox.showerror("Macro", "Loops value must be a positive integer.")
            return
        if loops <= 0:
            messagebox.showerror("Macro", "Loops value must be at least 1.")
            return
        try:
            start_delay = float(self.macro_play_delay_var.get())
        except (TypeError, ValueError):
            messagebox.showerror("Macro", "Start delay must be numeric.")
            return
        if start_delay < 0:
            messagebox.showerror("Macro", "Start delay cannot be negative.")
            return
        if self.macro_play_thread and self.macro_play_thread.is_alive():
            return

        self.macro_stop_event.clear()
        self.macro_run_reason = "Playing"
        self.block_macro_capture = True
        self.macro_status_var.set("Playing")

        events_snapshot = list(self.macro_events)
        self.macro_play_thread = threading.Thread(
            target=self._macro_play_worker,
            args=(events_snapshot, loops, start_delay),
            daemon=True,
        )
        self.macro_play_thread.start()

    def stop_macro_playback(self) -> None:
        if not self.macro_play_thread or not self.macro_play_thread.is_alive():
            return
        self.macro_run_reason = "Stopped"
        self.macro_stop_event.set()
        self.macro_play_thread.join(timeout=2.0)
        self._finalize_macro_stop()

    def _macro_play_worker(self, events: Sequence[MacroEvent], loops: int, start_delay: float) -> None:
        try:
            if start_delay > 0:
                deadline = time.time() + start_delay
                while time.time() < deadline and not self.macro_stop_event.is_set():
                    remaining = max(0.0, deadline - time.time())
                    self._safe_call_ui(self.macro_status_var.set, f"Playing in {remaining:.1f}s")
                    time.sleep(0.1)

            self._safe_call_ui(self.macro_status_var.set, "Playing")
            for loop_index in range(loops):
                if self.macro_stop_event.is_set():
                    break
                loop_label = f"Playing loop {loop_index + 1}/{loops}"
                self._safe_call_ui(self.macro_status_var.set, loop_label)
                for event in events:
                    if self.macro_stop_event.is_set():
                        break
                    if event.delay > 0 and self.macro_stop_event.wait(event.delay):
                        break
                    self._apply_macro_event(event)
                else:
                    continue
                break
            else:
                if self.macro_run_reason == "Playing":
                    self.macro_run_reason = "Completed"
        except Exception as exc:
            self.macro_run_reason = f"Error: {exc}"
        finally:
            self.root.after(0, self._finalize_macro_stop)

    def _apply_macro_event(self, event: MacroEvent) -> None:
        if event.kind == "key":
            key_value = event.key_value
            if key_value is None:
                return
            try:
                if event.action == "press":
                    self.key_controller.press(key_value)
                else:
                    self.key_controller.release(key_value)
            except Exception:
                pass
            return

        if event.kind == "mouse":
            if event.position:
                try:
                    self.mouse_controller.position = event.position
                except Exception:
                    pass
            if event.mouse_button is None:
                return
            try:
                if event.action == "press":
                    self.mouse_controller.press(event.mouse_button)
                else:
                    self.mouse_controller.release(event.mouse_button)
            except Exception:
                pass

    def _finalize_macro_stop(self) -> None:
        self.block_macro_capture = False
        if self.macro_play_thread and self.macro_play_thread.is_alive():
            return
        self.macro_play_thread = None
        self.macro_stop_event.set()
        self.macro_status_var.set(self.macro_run_reason)

    def _collect_settings(self) -> ClickSettings:
        try:
            delay_ms = float(self.delay_var.get())
            jitter_ms = float(self.jitter_var.get())
            burst = int(self.burst_var.get())
            start_delay_s = float(self.start_delay_var.get())
            limit_count = int(self.limit_count_var.get())
        except (TypeError, ValueError) as exc:
            raise ValueError("One or more numeric fields contain invalid values.") from exc

        if delay_ms <= 0:
            raise ValueError("Delay per cycle must be greater than zero.")
        if jitter_ms < 0:
            raise ValueError("Random jitter cannot be negative.")
        if jitter_ms > delay_ms:
            raise ValueError("Random jitter should not exceed the base delay.")
        if burst <= 0:
            raise ValueError("Clicks per cycle must be at least 1.")
        if start_delay_s < 0:
            raise ValueError("Start delay cannot be negative.")
        if self.limit_enabled_var.get() and limit_count <= 0:
            raise ValueError("Click limit must be at least 1.")

        button_key = self.button_var.get()
        button = self.BUTTON_MAP.get(button_key)
        if button is None:
            raise ValueError("Unsupported mouse button selection.")

        mode = self.mode_var.get()
        if mode not in {"single", "double", "hold"}:
            raise ValueError("Invalid click mode selected.")

        return ClickSettings(
            button=button,
            mode=mode,
            base_interval=max(0.005, delay_ms / 1000.0),
            jitter=jitter_ms / 1000.0,
            burst_count=burst,
            start_delay=start_delay_s,
            limit_enabled=self.limit_enabled_var.get(),
            limit_count=limit_count,
        )

    def _parse_key_sequence(self, raw: str) -> Tuple[Union[keyboard.Key, str], ...]:
        tokens = [token.strip() for token in raw.split("+") if token.strip()]
        if not tokens:
            raise ValueError("Enter at least one key for the macro.")

        resolved: list[Union[keyboard.Key, str]] = []
        for token in tokens:
            lower = token.lower()
            if len(lower) == 1:
                resolved.append(lower)
                continue
            try:
                resolved.append(getattr(keyboard.Key, lower))
            except AttributeError as exc:
                raise ValueError(f"Unknown key token: {token}") from exc
        return tuple(resolved)

    def _collect_key_settings(self) -> KeyPressSettings:
        sequence = self._parse_key_sequence(self.key_sequence_var.get())

        try:
            delay_ms = float(self.key_delay_var.get())
            jitter_ms = float(self.key_jitter_var.get())
            burst = int(self.key_burst_var.get())
            start_delay_s = float(self.key_start_delay_var.get())
            limit_count = int(self.key_limit_count_var.get())
        except (TypeError, ValueError) as exc:
            raise ValueError("One or more keyboard fields contain invalid values.") from exc

        if delay_ms <= 0:
            raise ValueError("Delay per cycle must be greater than zero.")
        if jitter_ms < 0:
            raise ValueError("Random jitter cannot be negative.")
        if jitter_ms > delay_ms:
            raise ValueError("Random jitter should not exceed the base delay.")
        if burst <= 0:
            raise ValueError("Cycles per batch must be at least 1.")
        if start_delay_s < 0:
            raise ValueError("Start delay cannot be negative.")
        if self.key_limit_enabled_var.get() and limit_count <= 0:
            raise ValueError("Cycle limit must be at least 1.")

        mode = self.key_mode_var.get()
        if mode not in {"tap", "double", "hold"}:
            raise ValueError("Invalid key mode selected.")

        return KeyPressSettings(
            sequence=sequence,
            mode=mode,
            base_interval=max(0.01, delay_ms / 1000.0),
            jitter=jitter_ms / 1000.0,
            burst_count=burst,
            start_delay=start_delay_s,
            limit_enabled=self.key_limit_enabled_var.get(),
            limit_count=limit_count,
        )

    def _press_sequence(self, sequence: Sequence[Union[keyboard.Key, str]]) -> None:
        for key_token in sequence:
            self.key_controller.press(key_token)

    def _release_sequence(self, sequence: Sequence[Union[keyboard.Key, str]]) -> None:
        for key_token in reversed(sequence):
            self.key_controller.release(key_token)

    def _tap_sequence(self, sequence: Sequence[Union[keyboard.Key, str]], repeats: int = 1) -> None:
        for _ in range(repeats):
            self._press_sequence(sequence)
            time.sleep(0.01)
            self._release_sequence(sequence)
            time.sleep(0.01)

    def _finalize_mouse_stop(self) -> None:
        elapsed_snapshot = None
        if self.start_timestamp is not None:
            elapsed_snapshot = int(max(0, time.time() - self.start_timestamp))
            self.start_timestamp = None
        if elapsed_snapshot is not None:
            self.last_elapsed_seconds = elapsed_snapshot

        if self.is_running:
            self.is_running = False
            self.stop_event.set()
            self.hold_pressed = False

        self.click_thread = None
        self._safe_call_ui(self.status_var.set, self.run_reason)

    def _finalize_key_stop(self) -> None:
        elapsed_snapshot = None
        if self.key_start_timestamp is not None:
            elapsed_snapshot = int(max(0, time.time() - self.key_start_timestamp))
            self.key_start_timestamp = None
        if elapsed_snapshot is not None:
            self.key_last_elapsed_seconds = elapsed_snapshot

        if self.key_is_running:
            self.key_is_running = False
            self.key_stop_event.set()
            self.key_hold_active = False

        self.key_thread = None
        self.active_key_sequence = None
        self._safe_call_ui(self.key_status_var.set, self.key_run_reason)

    def _reset_mouse_stats(self) -> None:
        if self.is_running:
            messagebox.showinfo("Reset", "Stop the clicker before resetting statistics.")
            return
        self.clicks_recorded = 0
        self.start_timestamp = None
        self.last_elapsed_seconds = 0
        self.elapsed_var.set("Elapsed: 00:00:00")
        self.click_count_var.set("Clicks: 0")
        self.run_reason = "Idle"
        self.status_var.set("Idle")

    def _reset_key_stats(self) -> None:
        if self.key_is_running:
            messagebox.showinfo("Reset", "Stop the macro before resetting statistics.")
            return
        self.key_cycles_recorded = 0
        self.key_start_timestamp = None
        self.key_last_elapsed_seconds = 0
        self.key_elapsed_var.set("Elapsed: 00:00:00")
        self.key_cycle_var.set("Cycles: 0")
        self.key_run_reason = "Idle"
        self.key_status_var.set("Idle")
        self.active_key_sequence = None
        self.key_hold_active = False

    def _schedule_status_refresh(self) -> None:
        self._update_status_labels()
        self.root.after(200, self._schedule_status_refresh)

    def _update_status_labels(self) -> None:
        if self.is_running and self.start_timestamp is not None:
            elapsed = int(max(0, time.time() - self.start_timestamp))
            self.last_elapsed_seconds = elapsed
        else:
            elapsed = self.last_elapsed_seconds
        hours, remainder = divmod(elapsed, 3600)
        minutes, seconds = divmod(remainder, 60)
        self.elapsed_var.set(f"Elapsed: {hours:02d}:{minutes:02d}:{seconds:02d}")
        self.click_count_var.set(f"Clicks: {self.clicks_recorded}")

        if self.key_is_running and self.key_start_timestamp is not None:
            key_elapsed = int(max(0, time.time() - self.key_start_timestamp))
            self.key_last_elapsed_seconds = key_elapsed
        else:
            key_elapsed = self.key_last_elapsed_seconds
        hours, remainder = divmod(key_elapsed, 3600)
        minutes, seconds = divmod(remainder, 60)
        self.key_elapsed_var.set(f"Elapsed: {hours:02d}:{minutes:02d}:{seconds:02d}")
        self.key_cycle_var.set(f"Cycles: {self.key_cycles_recorded}")

    def _on_close(self) -> None:
        self.stop_event.set()
        self.key_stop_event.set()
        self.macro_stop_event.set()
        self.macro_recording = False
        if self.keyboard_listener:
            self.keyboard_listener.stop()
        if self.mouse_listener:
            self.mouse_listener.stop()
        if self.click_thread and self.click_thread.is_alive():
            self.click_thread.join(timeout=1.0)
        if self.key_thread and self.key_thread.is_alive():
            self.key_thread.join(timeout=1.0)
        if self.macro_play_thread and self.macro_play_thread.is_alive():
            self.macro_play_thread.join(timeout=1.0)
        if self.hold_pressed:
            try:
                self.mouse_controller.release(self.BUTTON_MAP[self.button_var.get()])
            except Exception:
                pass
        if self.key_hold_active and self.active_key_sequence:
            try:
                self._release_sequence(self.active_key_sequence)
            except Exception:
                pass
        if self.macro_run_reason not in {"Idle", "Ready"}:
            self.macro_run_reason = "Stopped"
        self._finalize_macro_stop()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = AutoClickerApp()
    app.run()


if __name__ == "__main__":
    main()
