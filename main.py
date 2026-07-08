import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog, ttk
import requests
import threading
import json
import os
import time
import sys
import numpy as np
import re

# ---------- Optional libraries ----------
try:
    import pyttsx3
    TTS_PYTTTSX3 = True
except ImportError:
    TTS_PYTTTSX3 = False

try:
    import win32com.client
    TTS_SAPI = True
except ImportError:
    TTS_SAPI = False

try:
    from vosk import Model, KaldiRecognizer
    import pyaudio
    VOSK_AVAILABLE = True
except ImportError:
    VOSK_AVAILABLE = False

try:
    import speech_recognition as sr
    POCKET_AVAILABLE = True
except ImportError:
    POCKET_AVAILABLE = False

try:
    import whisper
    WHISPER_AVAILABLE = True
except ImportError:
    WHISPER_AVAILABLE = False

CONFIG_FILE = "config.json"

class OllamaChatApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepSeek Chat")
        self.root.geometry("750x650")
        self.root.resizable(True, True)

        # Set a modern ttk theme
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10), padding=4)
        self.style.configure('TEntry', font=('Segoe UI', 10), padding=4)
        self.style.configure('TCombobox', font=('Segoe UI', 10))

        self.config = self.load_config()
        self.model_name = "nexusriot/deepseek-r1-abliterated:8b"
        self.ollama_url = "http://localhost:11434/api/generate"

        self.auto_send_stt = self.config.get("auto_send_stt", True)
        self.mute = self.config.get("mute", False)
        self.wake_word = self.config.get("wake_word", "seek").lower()
        self.wake_enabled = self.config.get("wake_enabled", True)
        self.vosk_model_path = self.config.get("vosk_model_path", "")
        self.audio_device_index = self.config.get("audio_device_index", None)
        self.tts_voice_id = self.config.get("tts_voice_id", None)
        self.debug_partials = self.config.get("debug_partials", False)
        self.stt_engine = self.config.get("stt_engine", "vosk")
        self.tts_engine_type = self.config.get("tts_engine_type", "sapi5" if TTS_SAPI else "pyttsx3")
        self.whisper_model_name = self.config.get("whisper_model", "tiny")

        # Personality
        self.personalities = [
            {"name": "Default Assistant", "prompt": "You are a helpful assistant."},
            {"name": "Creative Writer", "prompt": "You are a creative storyteller and poet."},
            {"name": "Philosopher", "prompt": "You are a thoughtful philosopher who asks deep questions."},
            {"name": "Coding Mentor", "prompt": "You are an expert programmer who gives clear code examples."},
            {"name": "Friendly Companion", "prompt": "You are a warm and empathetic friend."},
        ]
        self.current_personality_idx = 0
        self.system_prompt = self.personalities[0]["prompt"]

        self.conversation = []
        self.last_bot_reply = ""
        self.tts_lock = threading.Lock()
        self.audio_lock = threading.Lock()

        # Vosk objects
        self.vosk_model = None
        self.audio = None
        self.stream = None
        self.wake_listening = False
        self.wake_thread = None
        self.stop_wake_event = threading.Event()
        self.is_awake = False

        self.whisper_model = None
        self.whisper_available = WHISPER_AVAILABLE

        self.create_widgets()
        self.input_field.bind("<Return>", self.send_message)
        self.load_default_session()

        self.init_stt()
        if self.wake_enabled and self.stt_engine == "vosk" and self.vosk_model is not None:
            self.start_wake_listener()

    # ---------- Config ----------
    def load_config(self):
        default = {
            "session_dir": "",
            "auto_send_stt": True,
            "mute": False,
            "wake_word": "seek",
            "wake_enabled": True,
            "vosk_model_path": "",
            "audio_device_index": None,
            "tts_voice_id": None,
            "debug_partials": False,
            "stt_engine": "vosk",
            "tts_engine_type": "sapi5" if TTS_SAPI else "pyttsx3",
            "whisper_model": "tiny"
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except:
                return default
        return default

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.config, f, indent=2)
        except:
            pass

    # ---------- Vosk model path helper ----------
    def find_vosk_model_folder(self, base_path):
        if not os.path.exists(base_path):
            return None
        if os.path.exists(os.path.join(base_path, "am")) and os.path.exists(os.path.join(base_path, "conf")):
            return base_path
        for item in os.listdir(base_path):
            sub = os.path.join(base_path, item)
            if os.path.isdir(sub):
                if os.path.exists(os.path.join(sub, "am")) and os.path.exists(os.path.join(sub, "conf")):
                    return sub
        return None

    # ---------- STT Init ----------
    def init_stt(self):
        if self.stt_engine == "vosk":
            if not VOSK_AVAILABLE:
                self.status_var.set("Vosk not installed. Switch engine.")
                return False
            if not self.vosk_model_path or not os.path.exists(self.vosk_model_path):
                self.status_var.set("Vosk model path not set or invalid.")
                return False
            model_dir = self.find_vosk_model_folder(self.vosk_model_path)
            if model_dir is None:
                self.status_var.set("Vosk model folder not found (missing am/conf/).")
                return False
            try:
                self.vosk_model = Model(model_dir)
                self.status_var.set("Vosk model loaded.")
                return True
            except Exception as e:
                self.status_var.set(f"Vosk model error: {e}")
                return False

        elif self.stt_engine == "pocketsphinx":
            if not POCKET_AVAILABLE:
                self.status_var.set("PocketSphinx not installed.")
                return False
            return True

        elif self.stt_engine == "whisper":
            if not self.whisper_available:
                self.status_var.set("Whisper not installed (pip install openai-whisper).")
                return False
            self.status_var.set("Whisper selected (model loads on first use).")
            return True

        return False

    # ---------- Audio ----------
    def get_audio_devices(self):
        devices = []
        if not VOSK_AVAILABLE:
            return devices
        p = pyaudio.PyAudio()
        for i in range(p.get_device_count()):
            info = p.get_device_info_by_index(i)
            if info.get("maxInputChannels", 0) > 0:
                devices.append((i, info.get("name")))
        p.terminate()
        return devices

    def open_audio_stream(self, device_index=None, force_default=False):
        if self.stt_engine == "pocketsphinx":
            return None
        with self.audio_lock:
            if self.audio is None:
                self.audio = pyaudio.PyAudio()
            if self.stream is not None:
                return self.stream

            idx = None
            if not force_default:
                idx = device_index if device_index is not None else self.audio_device_index
            if idx is not None:
                try:
                    self.audio.get_device_info_by_index(idx)
                except:
                    idx = None

            try:
                self.stream = self.audio.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=16000,
                    input=True,
                    input_device_index=idx,
                    frames_per_buffer=4000
                )
                return self.stream
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"Audio stream error: {e}"))
                try:
                    self.stream = self.audio.open(
                        format=pyaudio.paInt16,
                        channels=1,
                        rate=16000,
                        input=True,
                        input_device_index=None,
                        frames_per_buffer=4000
                    )
                    return self.stream
                except Exception as e2:
                    self.root.after(0, lambda: self.status_var.set(f"Failed to open audio: {e2}"))
                    return None

    def close_audio_stream(self):
        with self.audio_lock:
            if self.stream:
                self.stream.stop_stream()
                self.stream.close()
                self.stream = None
            if self.audio:
                self.audio.terminate()
                self.audio = None

    # ---------- TTS ----------
    def get_voices(self):
        if self.tts_engine_type == "sapi5" and TTS_SAPI:
            try:
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                voices = speaker.GetVoices()
                result = []
                for v in voices:
                    desc = v.GetDescription() if v.GetDescription() else f"Voice {v.Id}"
                    result.append((v.Id, desc))
                return result
            except:
                return []
        elif self.tts_engine_type == "pyttsx3" and TTS_PYTTTSX3:
            try:
                engine = pyttsx3.init()
                voices = engine.getProperty('voices')
                result = [(v.id, v.name if v.name else f"Voice {v.id}") for v in voices]
                engine.stop()
                return result
            except:
                return []
        return []

    def clean_tts_text(self, text):
        allowed = re.compile(r'[^a-zA-Z0-9\s.,!?;:\-]')
        cleaned = allowed.sub('', text)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned

    def speak_text(self, text, force=False):
        if not text or (not force and self.mute):
            return
        clean_text = self.clean_tts_text(text)
        if not clean_text:
            return
        threading.Thread(target=self._speak_text, args=(clean_text,), daemon=True).start()

    def _speak_text(self, text):
        with self.tts_lock:
            try:
                if self.tts_engine_type == "sapi5" and TTS_SAPI:
                    speaker = win32com.client.Dispatch("SAPI.SpVoice")
                    if self.tts_voice_id:
                        try:
                            speaker.Voice = speaker.GetVoices().Item(self.tts_voice_id)
                        except:
                            pass
                    speaker.Speak(text)
                elif self.tts_engine_type == "pyttsx3" and TTS_PYTTTSX3:
                    engine = pyttsx3.init()
                    if self.tts_voice_id is not None:
                        try:
                            engine.setProperty('voice', self.tts_voice_id)
                        except:
                            pass
                    engine.say(text)
                    engine.runAndWait()
                    engine.stop()
                else:
                    self.root.after(0, lambda: self.status_var.set("TTS: no engine available"))
                    return
                self.root.after(0, lambda: self.status_var.set("TTS done"))
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"TTS error: {e}"))

    def test_tts(self):
        if self.mute:
            messagebox.showwarning("TTS Test", "Mute is on – un‑mute first.")
            return
        self.speak_text("This is a test of the text‑to‑speech engine.", force=True)
        self.status_var.set("TTS test sent")

    # ---------- Manual Listen ----------
    def start_listen(self, after_command=None):
        if self.stt_engine == "vosk" and self.vosk_model is None:
            messagebox.showerror("STT Error", "Vosk not ready. Check Settings (model path).")
            return
        if self.stt_engine == "pocketsphinx" and not POCKET_AVAILABLE:
            messagebox.showerror("STT Error", "PocketSphinx not installed.")
            return
        if self.stt_engine == "whisper" and not self.whisper_available:
            messagebox.showerror("STT Error", "Whisper not installed.")
            return

        if self.wake_listening:
            self.stop_wake_listener()

        self.status_var.set("Listening... speak now")
        self.listen_btn.config(state="disabled")
        threading.Thread(target=self._listen_and_process, args=(after_command,), daemon=True).start()

    def _listen_and_process(self, after_command=None):
        try:
            if self.stt_engine == "vosk":
                stream = self.open_audio_stream()
                if stream is None:
                    self.root.after(0, lambda: messagebox.showerror("STT Error", "Could not open audio stream."))
                    return
                rec = KaldiRecognizer(self.vosk_model, 16000)
                rec.SetWords(True)
                rec.SetPartialWords(True)
                timeout = time.time() + 5
                while time.time() < timeout:
                    data = stream.read(4000, exception_on_overflow=False)
                    if rec.AcceptWaveform(data):
                        result = json.loads(rec.Result())
                        text = result.get("text", "").strip()
                        if text:
                            self.root.after(0, self._on_speech_recognized, text, after_command)
                            return
                self.root.after(0, lambda: self.status_var.set("No speech detected"))

            elif self.stt_engine == "pocketsphinx":
                r = sr.Recognizer()
                with sr.Microphone(device_index=self.audio_device_index) as source:
                    r.adjust_for_ambient_noise(source, duration=0.5)
                    audio = r.listen(source, timeout=5, phrase_time_limit=5)
                    try:
                        text = r.recognize_sphinx(audio)
                        self.root.after(0, self._on_speech_recognized, text, after_command)
                    except sr.UnknownValueError:
                        self.root.after(0, lambda: self.status_var.set("Could not understand"))
                    except sr.RequestError as e:
                        self.root.after(0, lambda: messagebox.showerror("STT Error", str(e)))

            elif self.stt_engine == "whisper":
                if self.whisper_model is None:
                    self.root.after(0, lambda: self.status_var.set("Loading Whisper model..."))
                    try:
                        self.whisper_model = whisper.load_model(self.whisper_model_name)
                    except Exception as e:
                        self.root.after(0, lambda: messagebox.showerror("Whisper Error", f"Failed to load model: {e}"))
                        return

                stream = self.open_audio_stream()
                if stream is None:
                    self.root.after(0, lambda: messagebox.showerror("STT Error", "Could not open audio stream."))
                    return

                self.root.after(0, lambda: self.status_var.set("Recording... (5 sec max)"))
                frames = []
                timeout = time.time() + 5
                while time.time() < timeout:
                    data = stream.read(4000, exception_on_overflow=False)
                    frames.append(data)
                audio_bytes = b''.join(frames)
                audio_np = np.frombuffer(audio_bytes, dtype=np.int16).astype(np.float32) / 32768.0

                self.root.after(0, lambda: self.status_var.set("Transcribing with Whisper..."))
                result = self.whisper_model.transcribe(audio_np, language="en", fp16=False)
                text = result["text"].strip()
                if text:
                    self.root.after(0, self._on_speech_recognized, text, after_command)
                else:
                    self.root.after(0, lambda: self.status_var.set("No speech detected"))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("STT Error", str(e)))
        finally:
            self.root.after(0, lambda: self.listen_btn.config(state="normal"))
            if after_command:
                self.root.after(0, after_command)

    def _on_speech_recognized(self, text, after_command=None):
        self.status_var.set(f"Recognized: {text}")
        self.input_field.delete(0, tk.END)
        self.input_field.insert(0, text)
        if self.auto_send_stt:
            self.root.after(500, self.send_message)
        else:
            self.input_field.focus_set()
        if after_command:
            self.root.after(1000, after_command)

    # ---------- Wake Word (fast detection via partials) ----------
    def start_wake_listener(self):
        if self.stt_engine != "vosk" or self.vosk_model is None or not self.wake_enabled:
            return
        if self.wake_listening:
            return
        self.wake_listening = True
        self.stop_wake_event.clear()
        self.wake_thread = threading.Thread(target=self._wake_loop, daemon=True)
        self.wake_thread.start()
        self.status_var.set(f"Wake word '{self.wake_word}' listening...")
        self.wake_status_label.config(text="🟢 Wake: "+self.wake_word, foreground="green")

    def stop_wake_listener(self):
        if self.wake_listening:
            self.stop_wake_event.set()
            self.wake_listening = False
            if self.wake_thread:
                self.wake_thread.join(timeout=1)
            self.wake_thread = None
            self.status_var.set("Wake listener stopped")
            self.wake_status_label.config(text="🔴 Wake off", foreground="red")
            self.close_audio_stream()

    def _wake_loop(self):
        stream = None
        rec = None

        while not self.stop_wake_event.is_set():
            try:
                if stream is None or not self.stream:
                    stream = self.open_audio_stream(force_default=False)
                    if stream is None:
                        self.root.after(0, lambda: self.status_var.set("Wake: retrying with default mic..."))
                        stream = self.open_audio_stream(force_default=True)
                    if stream is None:
                        self.root.after(0, lambda: self.status_var.set("Wake: no microphone available"))
                        time.sleep(2)
                        continue
                    rec = KaldiRecognizer(self.vosk_model, 16000)
                    rec.SetWords(False)
                    rec.SetPartialWords(True)

                data = stream.read(4000, exception_on_overflow=False)

                # Check partial result first for low latency
                partial = json.loads(rec.PartialResult())
                ptext = partial.get("partial", "").strip().lower()
                if self.wake_word in ptext:
                    self.root.after(0, self._on_wake_detected)
                    self.stop_wake_event.wait(0.5)
                    continue

                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    text = result.get("text", "").strip().lower()
                    if self.wake_word in text:
                        self.root.after(0, self._on_wake_detected)
                        self.stop_wake_event.wait(0.5)
                else:
                    if self.debug_partials and ptext:
                        self.root.after(0, lambda t=ptext: self.status_var.set(f"Wake debug: {t}"))

            except OSError as e:
                self.root.after(0, lambda: self.status_var.set(f"Wake audio error: {e}"))
                self.close_audio_stream()
                stream = None
                time.sleep(1)
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"Wake error: {e}"))
                time.sleep(1)

    def _on_wake_detected(self):
        if self.is_awake:
            return
        self.is_awake = True
        self.status_var.set("🔊 Awake – listening for command...")
        self.wake_status_label.config(text="⚡ WOKE!", foreground="orange")
        self.stop_wake_listener()
        self.start_listen(after_command=self._after_command)

    def _after_command(self):
        self.is_awake = False
        if self.wake_enabled and self.stt_engine == "vosk" and self.vosk_model is not None:
            self.start_wake_listener()

    # ---------- Settings Window (scrollable with fixed Save/Cancel) ----------
    def open_settings_window(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.geometry("650x800")
        win.resizable(True, True)
        win.configure(bg='#f0f0f0')

        # Outer container to hold the canvas and the button bar
        outer_frame = ttk.Frame(win)
        outer_frame.pack(fill=tk.BOTH, expand=True)

        # Canvas and scrollbar for content
        canvas = tk.Canvas(outer_frame, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Frame inside canvas that will hold all settings widgets
        content_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor='nw', width=610)

        # Buttons at the bottom (outside canvas)
        btn_frame = ttk.Frame(outer_frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        # Pack canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configure content frame to update scroll region on resize
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        content_frame.bind("<Configure>", configure_scroll_region)

        # Populate content_frame with settings (using grid or pack)
        # We'll use grid inside content_frame
        row = 0
        pad = 5
        label_font = ('Segoe UI', 10, 'bold')

        # Session Directory
        ttk.Label(content_frame, text="Default Session Directory:", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        dir_frame = ttk.Frame(content_frame)
        dir_frame.grid(row=row, column=0, sticky='ew', pady=(0,10))
        dir_var = tk.StringVar(value=self.config.get("session_dir", ""))
        entry1 = ttk.Entry(dir_frame, textvariable=dir_var, width=50)
        entry1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        def browse_dir():
            selected = filedialog.askdirectory(title="Select Default Session Folder")
            if selected:
                dir_var.set(selected)
        ttk.Button(dir_frame, text="Browse...", command=browse_dir).pack(side=tk.RIGHT)
        row += 1

        # STT Engine
        ttk.Label(content_frame, text="Speech-to-Text Engine:", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        stt_frame = ttk.Frame(content_frame)
        stt_frame.grid(row=row, column=0, sticky='ew', pady=(0,10))
        stt_var = tk.StringVar(value=self.stt_engine)
        stt_choices = ["vosk", "pocketsphinx"]
        if self.whisper_available:
            stt_choices.append("whisper")
        stt_combo = ttk.Combobox(stt_frame, textvariable=stt_var, state="readonly", values=stt_choices, width=20)
        stt_combo.pack(side=tk.LEFT, padx=5)
        stt_status_label = ttk.Label(stt_frame, text="", foreground='red')
        stt_status_label.pack(side=tk.LEFT, padx=10)
        row += 1

        # Vosk Model Path
        ttk.Label(content_frame, text="Vosk Model Directory (for Vosk engine):", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        vosk_frame = ttk.Frame(content_frame)
        vosk_frame.grid(row=row, column=0, sticky='ew', pady=(0,10))
        vosk_var = tk.StringVar(value=self.vosk_model_path)
        entry2 = ttk.Entry(vosk_frame, textvariable=vosk_var, width=50)
        entry2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))

        def browse_vosk():
            selected = filedialog.askdirectory(title="Select Vosk Model Folder (or parent)")
            if selected:
                actual = self.find_vosk_model_folder(selected)
                if actual is not None:
                    vosk_var.set(actual)
                else:
                    vosk_var.set(selected)
                    messagebox.showwarning("Model Detection", "Could not find a Vosk model (am/conf/ subfolder) in the selected directory.\nPlease select the folder that contains 'am' and 'conf' subdirectories.")
        ttk.Button(vosk_frame, text="Browse...", command=browse_vosk).pack(side=tk.RIGHT)
        row += 1

        def test_vosk():
            path = vosk_var.get().strip()
            actual = self.find_vosk_model_folder(path) if path else None
            if actual is None:
                messagebox.showerror("Test", "Vosk model folder not found (missing am/conf/).")
                return
            try:
                model = Model(actual)
                messagebox.showinfo("Test", f"Vosk model loaded successfully from:\n{actual}")
                model = None
            except Exception as e:
                messagebox.showerror("Test", f"Failed to load model:\n{str(e)}")
        ttk.Button(content_frame, text="Test Vosk Model", command=test_vosk).grid(row=row, column=0, pady=5)
        row += 1

        # Whisper model size
        ttk.Label(content_frame, text="Whisper Model Size (if using Whisper):", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        whisper_frame = ttk.Frame(content_frame)
        whisper_frame.grid(row=row, column=0, sticky='ew', pady=(0,10))
        whisper_var = tk.StringVar(value=self.whisper_model_name)
        whisper_combo = ttk.Combobox(whisper_frame, textvariable=whisper_var, state="readonly",
                                     values=["tiny", "base", "small", "medium", "large"], width=15)
        whisper_combo.pack(side=tk.LEFT, padx=5)
        ttk.Label(whisper_frame, text="(larger = more accurate but slower)", foreground='gray').pack(side=tk.LEFT, padx=5)
        row += 1

        # Microphone
        ttk.Label(content_frame, text="Microphone:", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        mic_frame = ttk.Frame(content_frame)
        mic_frame.grid(row=row, column=0, sticky='ew', pady=(0,10))
        devices = self.get_audio_devices()
        mic_var = tk.StringVar()
        mic_combo = ttk.Combobox(mic_frame, textvariable=mic_var, state="readonly", width=50)
        mic_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        device_names, device_indices = [], []
        for idx, name in devices:
            device_names.append(f"{idx}: {name}")
            device_indices.append(idx)
        mic_combo['values'] = device_names
        current_idx = self.audio_device_index
        if current_idx is not None and current_idx in device_indices:
            pos = device_indices.index(current_idx)
            mic_combo.current(pos)
        elif device_names:
            mic_combo.current(0)

        mic_status_label = ttk.Label(content_frame, text="", foreground='red')
        mic_status_label.grid(row=row+1, column=0, sticky='w', pady=(0,5))
        if not devices:
            mic_status_label.config(text="⚠️ No input devices found. Check pyaudio and microphone.")
        row += 2

        def test_mic():
            if self.stt_engine == "vosk" and self.vosk_model is None:
                messagebox.showerror("Test", "Vosk not ready. Set a valid model path and test it first.")
                return
            self.status_var.set("Testing mic... speak for 3 sec")
            threading.Thread(target=self._test_mic_thread, daemon=True).start()
        ttk.Button(content_frame, text="Test Microphone", command=test_mic).grid(row=row, column=0, pady=5)
        row += 1

        # TTS Engine
        ttk.Label(content_frame, text="TTS Engine:", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        tts_engine_frame = ttk.Frame(content_frame)
        tts_engine_frame.grid(row=row, column=0, sticky='ew', pady=(0,5))
        tts_engine_var = tk.StringVar(value=self.tts_engine_type)
        tts_choices = []
        if TTS_SAPI:
            tts_choices.append("sapi5")
        if TTS_PYTTTSX3:
            tts_choices.append("pyttsx3")
        if not tts_choices:
            tts_choices = ["none"]
        tts_engine_combo = ttk.Combobox(tts_engine_frame, textvariable=tts_engine_var, state="readonly", values=tts_choices, width=15)
        tts_engine_combo.pack(side=tk.LEFT, padx=5)
        row += 1

        # TTS Voice
        ttk.Label(content_frame, text="TTS Voice:", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        voice_frame = ttk.Frame(content_frame)
        voice_frame.grid(row=row, column=0, sticky='ew', pady=(0,5))
        voices = self.get_voices()
        voice_var = tk.StringVar()
        voice_combo = ttk.Combobox(voice_frame, textvariable=voice_var, state="readonly", width=50)
        voice_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        voice_names, voice_ids = [], []
        for vid, name in voices:
            voice_names.append(name)
            voice_ids.append(vid)
        voice_combo['values'] = voice_names
        current_voice_id = self.tts_voice_id
        if current_voice_id is not None and current_voice_id in voice_ids:
            pos = voice_ids.index(current_voice_id)
            voice_combo.current(pos)
        elif voice_names:
            voice_combo.current(0)

        voice_status_label = ttk.Label(content_frame, text="", foreground='red')
        voice_status_label.grid(row=row+1, column=0, sticky='w', pady=(0,5))
        if not voices:
            voice_status_label.config(text="⚠️ No voices found. Check TTS engine.")
        else:
            ttk.Label(content_frame, text="💡 For more natural voices on Windows, install third‑party SAPI5 voices.", foreground='gray').grid(row=row+2, column=0, sticky='w', pady=(0,10))
        row += 3

        ttk.Button(content_frame, text="Test TTS", command=self.test_tts).grid(row=row, column=0, pady=5)
        row += 1

        # Wake Word (only works with Vosk)
        ttk.Label(content_frame, text="Wake Word (only for Vosk):", font=label_font).grid(row=row, column=0, sticky='w', pady=(10,5))
        row += 1
        wake_frame = ttk.Frame(content_frame)
        wake_frame.grid(row=row, column=0, sticky='ew', pady=(0,5))
        wake_var = tk.StringVar(value=self.wake_word)
        wake_entry = ttk.Entry(wake_frame, textvariable=wake_var, width=20)
        wake_entry.pack(side=tk.LEFT, padx=5)
        wake_enabled_var = tk.BooleanVar(value=self.wake_enabled)
        ttk.Checkbutton(wake_frame, text="Enable wake word", variable=wake_enabled_var).pack(side=tk.LEFT, padx=10)
        row += 1

        # Debug partials
        debug_var = tk.BooleanVar(value=self.debug_partials)
        ttk.Checkbutton(content_frame, text="Show partial transcriptions (debug)", variable=debug_var).grid(row=row, column=0, sticky='w', pady=5)
        row += 1

        # Auto-send
        auto_var = tk.BooleanVar(value=self.auto_send_stt)
        ttk.Checkbutton(content_frame, text="Auto-send after voice input", variable=auto_var).grid(row=row, column=0, sticky='w', pady=5)
        row += 1

        # Spacer to push content up
        ttk.Label(content_frame, text="").grid(row=row, column=0, pady=10)
        row += 1

        # --- Buttons at the bottom (in btn_frame) ---
        def save_settings():
            new_dir = dir_var.get().strip()
            self.config["session_dir"] = new_dir

            new_stt = stt_var.get().strip()
            self.config["stt_engine"] = new_stt
            self.stt_engine = new_stt

            new_vosk = vosk_var.get().strip()
            self.config["vosk_model_path"] = new_vosk
            self.vosk_model_path = new_vosk

            new_whisper_model = whisper_var.get().strip()
            self.config["whisper_model"] = new_whisper_model
            self.whisper_model_name = new_whisper_model
            self.whisper_model = None

            if mic_combo.current() >= 0 and device_indices:
                idx = device_indices[mic_combo.current()]
                self.audio_device_index = idx
                self.config["audio_device_index"] = idx
            else:
                self.audio_device_index = None
                self.config["audio_device_index"] = None

            new_tts_eng = tts_engine_var.get().strip()
            self.config["tts_engine_type"] = new_tts_eng
            self.tts_engine_type = new_tts_eng

            if voice_combo.current() >= 0 and voice_ids:
                vid = voice_ids[voice_combo.current()]
                self.tts_voice_id = vid
                self.config["tts_voice_id"] = vid
            else:
                self.tts_voice_id = None
                self.config["tts_voice_id"] = None

            self.auto_send_stt = auto_var.get()
            self.config["auto_send_stt"] = self.auto_send_stt

            self.debug_partials = debug_var.get()
            self.config["debug_partials"] = self.debug_partials

            new_wake = wake_var.get().strip().lower()
            self.config["wake_word"] = new_wake
            self.wake_word = new_wake
            self.wake_enabled = wake_enabled_var.get()
            self.config["wake_enabled"] = self.wake_enabled

            self.save_config()

            self.close_audio_stream()
            if self.stt_engine == "vosk":
                self.init_stt()
            else:
                self.vosk_model = None

            if self.stt_engine == "vosk" and self.wake_enabled and self.vosk_model is not None:
                self.stop_wake_listener()
                self.start_wake_listener()
            else:
                self.stop_wake_listener()
                self.wake_status_label.config(text="🔴 Wake off", foreground="red")
                if self.stt_engine != "vosk":
                    self.status_var.set("Wake word disabled for this STT engine")

            self.status_var.set("Settings saved")
            win.destroy()

        ttk.Button(btn_frame, text="Save", command=save_settings, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy, width=10).pack(side=tk.LEFT, padx=5)

        # Force canvas to update scroll region
        content_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    # ---------- Test Mic ----------
    def _test_mic_thread(self):
        try:
            if self.stt_engine == "vosk":
                stream = self.open_audio_stream()
                if stream is None:
                    self.root.after(0, lambda: self.status_var.set("Test: no audio stream"))
                    return
                rec = KaldiRecognizer(self.vosk_model, 16000)
                rec.SetWords(False)
                rec.SetPartialWords(True)
                start = time.time()
                while time.time() - start < 3:
                    data = stream.read(4000, exception_on_overflow=False)
                    if rec.AcceptWaveform(data):
                        result = json.loads(rec.Result())
                        text = result.get("text", "")
                        if text:
                            self.root.after(0, lambda: self.status_var.set(f"Test heard: {text}"))
                            return
                    else:
                        partial = json.loads(rec.PartialResult())
                        ptext = partial.get("partial", "")
                        if ptext:
                            self.root.after(0, lambda t=ptext: self.status_var.set(f"Test partial: {t}"))
                self.root.after(0, lambda: self.status_var.set("Test complete (no clear speech)"))

            elif self.stt_engine == "pocketsphinx":
                r = sr.Recognizer()
                with sr.Microphone(device_index=self.audio_device_index) as source:
                    r.adjust_for_ambient_noise(source, duration=0.5)
                    audio = r.listen(source, timeout=3, phrase_time_limit=3)
                    try:
                        text = r.recognize_sphinx(audio)
                        self.root.after(0, lambda: self.status_var.set(f"Test heard: {text}"))
                    except:
                        self.root.after(0, lambda: self.status_var.set("Test: could not recognize"))

            elif self.stt_engine == "whisper":
                if self.whisper_model is None:
                    self.root.after(0, lambda: self.status_var.set("Loading Whisper model..."))
                    try:
                        self.whisper_model = whisper.load_model(self.whisper_model_name)
                    except Exception as e:
                        self.root.after(0, lambda: messagebox.showerror("Whisper Error", f"Failed to load model: {e}"))
                        return
                stream = self.open_audio_stream()
                if stream is None:
                    self.root.after(0, lambda: self.status_var.set("Test: no audio stream"))
                    return
                self.root.after(0, lambda: self.status_var.set("Recording... (3 sec)"))
                frames = []
                start = time.time()
                while time.time() - start < 3:
                    data = stream.read(4000, exception_on_overflow=False)
                    frames.append(data)
                audio_bytes = b''.join(frames)
                audio_np = np.frombuffer(audio_bytes, dtype=np.int16).astype(np.float32) / 32768.0
                self.root.after(0, lambda: self.status_var.set("Transcribing..."))
                result = self.whisper_model.transcribe(audio_np, language="en", fp16=False)
                text = result["text"].strip()
                if text:
                    self.root.after(0, lambda: self.status_var.set(f"Test heard: {text}"))
                else:
                    self.root.after(0, lambda: self.status_var.set("Test: no speech detected"))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Test Error", str(e)))

    # ---------- GUI ----------
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill=tk.X, pady=(0,10))
        ttk.Label(toolbar, text="Model:", font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=5)
        self.model_label = ttk.Label(toolbar, text=self.model_name, foreground='blue', font=('Segoe UI', 10, 'bold'))
        self.model_label.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Button(toolbar, text="Personalities", command=self.open_personality_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Settings", command=self.open_settings_window).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Save Session", command=self.save_session).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Load Session", command=self.load_session).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Clear History", command=self.clear_history).pack(side=tk.LEFT, padx=5)

        self.mute_btn = ttk.Button(toolbar, text="🔇 Mute" if self.mute else "🔊 Unmuted",
                                   command=self.toggle_mute)
        self.mute_btn.pack(side=tk.LEFT, padx=5)

        self.listen_btn = ttk.Button(toolbar, text="🎙️ Listen", command=self.start_listen, state="normal")
        self.listen_btn.pack(side=tk.LEFT, padx=5)

        self.wake_status_label = ttk.Label(toolbar, text="🔴 Wake off", foreground='red')
        self.wake_status_label.pack(side=tk.LEFT, padx=10)

        chat_frame = ttk.Frame(main_frame)
        chat_frame.pack(fill=tk.BOTH, expand=True, pady=(0,10))
        self.chat_display = scrolledtext.ScrolledText(
            chat_frame, wrap=tk.WORD, state='disabled', font=('Segoe UI', 11),
            bg='#ffffff', fg='#333333', relief='flat', borderwidth=0
        )
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        self.chat_display.tag_config("user", foreground="#2c3e50", font=('Segoe UI', 11, 'bold'))
        self.chat_display.tag_config("bot", foreground="#2980b9", font=('Segoe UI', 11))

        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(0,5))
        self.input_field = ttk.Entry(bottom_frame, font=('Segoe UI', 12))
        self.input_field.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        self.send_button = ttk.Button(bottom_frame, text="Send", command=self.send_message, width=10)
        self.send_button.pack(side=tk.RIGHT)

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief='sunken', anchor='w',
                               font=('Segoe UI', 9), background='#e0e0e0')
        status_bar.pack(fill=tk.X, pady=(5,0))

    def toggle_mute(self):
        self.mute = not self.mute
        self.config["mute"] = self.mute
        self.save_config()
        self.mute_btn.config(text="🔇 Mute" if self.mute else "🔊 Unmuted")
        self.status_var.set("Muted" if self.mute else "Unmuted")

    # ---------- Personality ----------
    def open_personality_window(self):
        win = tk.Toplevel(self.root)
        win.title("Select Personality")
        win.geometry("400x300")
        win.resizable(False, False)

        ttk.Label(win, text="Choose a personality for the AI:", font=('Segoe UI', 10, 'bold')).pack(pady=10)

        listbox = tk.Listbox(win, height=len(self.personalities), font=('Segoe UI', 10), relief='flat')
        listbox.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        for p in self.personalities:
            listbox.insert(tk.END, p["name"])
        listbox.selection_set(self.current_personality_idx)

        def apply_selection():
            idx = listbox.curselection()
            if idx:
                self.current_personality_idx = idx[0]
                self.system_prompt = self.personalities[idx[0]]["prompt"]
                self.status_var.set(f"Personality: {self.personalities[idx[0]]['name']}")
                win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Apply", command=apply_selection, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy, width=10).pack(side=tk.LEFT, padx=5)

    # ---------- Model Interaction ----------
    def build_prompt(self):
        messages = [{"role": "system", "content": self.system_prompt}] + self.conversation
        prompt_text = ""
        for msg in messages:
            if msg["role"] == "system":
                prompt_text += f"System: {msg['content']}\n"
            elif msg["role"] == "user":
                prompt_text += f"User: {msg['content']}\n"
            elif msg["role"] == "assistant":
                prompt_text += f"Assistant: {msg['content']}\n"
        prompt_text += "Assistant: "
        return prompt_text

    def send_message(self, event=None):
        user_input = self.input_field.get().strip()
        if not user_input:
            return

        self.input_field.delete(0, tk.END)
        self.display_message("User", user_input)
        self.conversation.append({"role": "user", "content": user_input})

        self.set_ui_state(False)
        self.status_var.set("Generating...")

        thread = threading.Thread(target=self.get_model_response)
        thread.daemon = True
        thread.start()

    def get_model_response(self):
        try:
            prompt = self.build_prompt()
            payload = {
                "model": self.model_name,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.7,
                    "top_p": 0.9,
                }
            }
            response = requests.post(self.ollama_url, json=payload, timeout=60)
            response.raise_for_status()
            data = response.json()
            bot_reply = data.get("response", "").strip()
            self.root.after(0, self.on_response_received, bot_reply)

        except requests.exceptions.RequestException as e:
            self.root.after(0, self.on_error, f"Network/API error: {e}")
        except Exception as e:
            self.root.after(0, self.on_error, f"Unexpected error: {e}")

    def on_response_received(self, bot_reply):
        self.display_message("DeepSeek-R1", bot_reply)
        self.conversation.append({"role": "assistant", "content": bot_reply})
        self.last_bot_reply = bot_reply
        self.set_ui_state(True)
        self.status_var.set("Ready")
        self.speak_text(bot_reply)

    def on_error(self, error_msg):
        messagebox.showerror("Error", error_msg)
        self.set_ui_state(True)
        self.status_var.set("Error")

    # ---------- Session ----------
    def save_session(self, filename=None):
        if not filename:
            save_dir = self.config.get("session_dir", "")
            if save_dir and os.path.exists(save_dir):
                filename = os.path.join(save_dir, f"session_{int(time.time())}.json")
            else:
                filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
                if not filename:
                    return
        try:
            with open(filename, 'w') as f:
                json.dump({
                    "conversation": self.conversation,
                    "system_prompt": self.system_prompt,
                    "personality_idx": self.current_personality_idx
                }, f, indent=2)
            self.status_var.set(f"Session saved to {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    def load_session(self, filename=None):
        if not filename:
            save_dir = self.config.get("session_dir", "")
            if save_dir and os.path.exists(save_dir):
                files = [f for f in os.listdir(save_dir) if f.endswith('.json')]
                if files:
                    filename = filedialog.askopenfilename(initialdir=save_dir, filetypes=[("JSON files", "*.json")])
                else:
                    filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
            else:
                filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
            if not filename:
                return
        try:
            with open(filename, 'r') as f:
                data = json.load(f)
            self.conversation = data.get("conversation", [])
            self.system_prompt = data.get("system_prompt", self.personalities[0]["prompt"])
            self.current_personality_idx = data.get("personality_idx", 0)
            self.chat_display.config(state='normal')
            self.chat_display.delete(1.0, tk.END)
            for msg in self.conversation:
                if msg["role"] == "user":
                    self.display_message("User", msg["content"])
                elif msg["role"] == "assistant":
                    self.display_message("DeepSeek-R1", msg["content"])
            self.chat_display.config(state='disabled')
            self.status_var.set(f"Session loaded from {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def load_default_session(self):
        save_dir = self.config.get("session_dir", "")
        if save_dir and os.path.exists(save_dir):
            default_path = os.path.join(save_dir, "last_session.json")
            if os.path.exists(default_path):
                try:
                    with open(default_path, 'r') as f:
                        data = json.load(f)
                    self.conversation = data.get("conversation", [])
                    self.system_prompt = data.get("system_prompt", self.personalities[0]["prompt"])
                    self.current_personality_idx = data.get("personality_idx", 0)
                    for msg in self.conversation:
                        if msg["role"] == "user":
                            self.display_message("User", msg["content"])
                        elif msg["role"] == "assistant":
                            self.display_message("DeepSeek-R1", msg["content"])
                    self.status_var.set("Loaded last session")
                except:
                    pass

    def clear_history(self):
        if self.conversation:
            if messagebox.askyesno("Clear History", "Are you sure you want to clear the conversation?"):
                self.conversation.clear()
                self.chat_display.config(state='normal')
                self.chat_display.delete(1.0, tk.END)
                self.chat_display.config(state='disabled')
                self.status_var.set("History cleared")

    def display_message(self, sender, text):
        self.chat_display.config(state='normal')
        if sender == "User":
            prefix = "You: "
            tag = "user"
        else:
            prefix = "DeepSeek-R1: "
            tag = "bot"
        self.chat_display.insert(tk.END, prefix + text + "\n\n", tag)
        self.chat_display.config(state='disabled')
        self.chat_display.see(tk.END)

    def set_ui_state(self, enabled):
        state = "normal" if enabled else "disabled"
        self.input_field.config(state=state)
        self.send_button.config(state=state)
        if enabled:
            self.input_field.focus_set()

    def on_closing(self):
        self.stop_wake_listener()
        self.close_audio_stream()
        if messagebox.askyesno("Quit", "Save session before exiting?"):
            save_dir = self.config.get("session_dir", "")
            if save_dir and os.path.exists(save_dir):
                default_path = os.path.join(save_dir, "last_session.json")
            else:
                default_path = "last_session.json"
            self.save_session(default_path)
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = OllamaChatApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
