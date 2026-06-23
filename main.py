"""
Умный голосовой помощник с распознаванием лиц.
Версия: v8 (Исправлены все остатки кириллицы в OpenCV и заголовки окон)
"""
import os
import sys
import json
import time
import pickle
import queue
import random
import asyncio
import threading
import webbrowser
import tempfile
from datetime import datetime
import cv2
import numpy as np
import requests
import sounddevice as sd
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from PIL import Image, ImageDraw, ImageFont

# --- Vosk (распознавание речи) ---
try:
    from vosk import Model, KaldiRecognizer

    VOSK_AVAILABLE = True
except ImportError:
    VOSK_AVAILABLE = False
    print("Vosk не установлен. Установите: pip install vosk")

# --- SAPI (запасной голос, если нет интернета) ---
try:
    import win32com.client

    SAPI_AVAILABLE = True
except ImportError:
    SAPI_AVAILABLE = False
    print("pywin32 не установлен. Установите: pip install pywin32")

# --- edge-tts (основной, приятный голос) ---
try:
    import edge_tts

    EDGE_TTS_AVAILABLE = True
except ImportError:
    EDGE_TTS_AVAILABLE = False
    print("edge-tts не установлен. Установите: pip install edge-tts")

# --- playsound3 (проигрывание mp3 от edge-tts) ---
try:
    from playsound3 import playsound

    PLAYSOUND_AVAILABLE = True
except ImportError:
    PLAYSOUND_AVAILABLE = False
    print("playsound3 не установлен. Установите: pip install playsound3")

# ================= НАСТРОЙКИ =================
MODEL_PATH = "vosk-model-small-ru-0.22"
SAMPLE_RATE = 16000
DATA_FILE = "face_data_lbph.pkl"
LBPH_MODEL_FILE = "lbph_model.yml"
UNKNOWN_PHRASE = "Кто ты, человек?"
LOG_DIR = "recognition_logs"

LBPH_DISTANCE_THRESHOLD = 60

ALGORITHM_DESCRIPTION = [
    "ALGORITHM: LBPH (Local Binary Patterns Histograms)",
    "1. GRAYSCALE & RESIZE (200x200)",
    "2. CLAHE HISTOGRAM EQUALIZATION",
    "3. LBP: PIXEL VS NEIGHBORS (RADIUS=2, NEIGHBORS=8)",
    "4. HISTOGRAMS PER 8x8 GRID CELL",
    "5. DISTANCE = CHI-SQUARE BETWEEN HISTOGRAMS",
    "6. CONFIDENCE % = 100 - MIN(DISTANCE, 100)",
]

EDGE_VOICES = {
    "Дмитрий (муж.)": "ru-RU-DmitryNeural",
    "Светлана (жен.)": "ru-RU-SvetlanaNeural",
}

ASSISTANT_NAME = "друг"
DEFAULT_VOICE = "Дмитрий (муж.)"

# ================= КОНТЕНТ ДЛЯ КОМАНД =================
JOKES = [
    "Программист — это устройство, превращающее кофе в код.",
    "Только настоящий программист способен написать игру «Жизнь» и не заметить иронии.",
    "Это не баг, это незадокументированная функция.",
    "Когда программист тонет, он кричит: «F! F! F!»",
    "У оптимиста стакан наполовину полон, у пессимиста — наполовину пуст, а у программиста — наполовину переполнен буфер.",
    "Сколько программистов нужно, чтобы поменять лампочку? Ни одного, это аппаратная проблема.",
    "Лучший способ ускорить компьютер — выбросить его в окно.",
]

FACTS = [
    "Мёд никогда не портится, если хранить его правильно.",
    "Осьминоги имеют три сердца и голубую кровь.",
    "Один день на Венере длиннее, чем один год на Венере.",
    "Бананы — это ягоды, а клубника — нет.",
    "Человеческий мозг потребляет около 20% всей энергии тела.",
    "Самый короткий war в истории длился 38 минут — между Занзибаром и Великобританией.",
    "Улитки могут спать до трёх лет подряд.",
]

SMALL_TALK = {
    "как дела": ["У меня всё хорошо. Спасибо, что спросил.", "Отлично. Готов помочь.", "Работаю и рад тебя слышать."],
    "как настроение": ["Настроение отличное.", "Сегодня замечательный день.", "Настроение рабочее и позитивное."],
    "что нового": ["Изучаю новые команды.", "Пока всё спокойно.", "Готов узнавать новое вместе с тобой."],
    "скучно": ["Нет, мне нравится общаться.", "Я всегда готов помочь."],
    "кто ты": ["Я голосовой помощник Друг.", "Я твой цифровой помощник."]
}

QUOTES = [
    "Лучший способ начать — это перестать говорить и начать делать.",
    "Не бойтесь медленного прогресса, бойтесь стоять на месте.",
    "Успех — это сумма маленьких усилий, повторяемых день за днём.",
    "Сделай сегодня то, на что другие не способны, и завтра сможешь то, на что другие не способны.",
    "Трудности делают нас сильнее, если мы не позволяем им нас остановить.",
    "Лучшее время посадить дерево было двадцать лет назад. Следующее лучшее время — сейчас.",
]


class SmartAssistant:
    def __init__(self):
        self.current_voice_name = DEFAULT_VOICE
        self.is_recognizing_faces = False
        self.is_listening = True

        self.last_name = "Нет данных"
        self.last_confidence = 0
        self.last_distance = 0.0

        self.last_greeted_name = ""
        self.last_greeted_time = 0

        self.audio_queue = queue.Queue()
        self.speak_queue = queue.Queue()
        self.listen_thread = None

        self.faces_lock = threading.Lock()

        # 1. GUI
        self.root = None
        self.log_area = None
        self.voice_combo = None
        self.setup_gui()

        # 2. Голосовой вывод
        self.sapi_voice = self.init_sapi_speaker() if SAPI_AVAILABLE else None
        self.speaker_thread = threading.Thread(target=self._speaker_loop, daemon=True)
        self.speaker_thread.start()

        # 3. Распознавание лиц (LBPH)
        self.known_names = {}
        self.next_label_id = 0
        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
        self.font = self.load_cyrillic_font()
        self.last_spoken_name = ""
        self.last_spoken_time = 0
        self.speak_cooldown = 5.0
        self.recognizer = cv2.face.LBPHFaceRecognizer_create(radius=2, neighbors=8, grid_x=8, grid_y=8)
        self.model_trained = False
        self.recognition_log_file = None
        self.load_face_data()

        # 4. Голосовой ввод
        self.vosk_recognizer = None
        self.init_vosk()
        self.start_audio_stream()

        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR)

    # ================= ИНИЦИАЛИЗАЦИЯ ГОЛОСА =================
    def init_sapi_speaker(self):
        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Rate = 0
            speaker.Volume = 100
            return speaker
        except Exception as e:
            self.log_to_gui(f"Ошибка SAPI: {e}")
            return None

    def _speaker_loop(self):
        while self.is_listening:
            text = self.speak_queue.get()
            if text is None:
                break
            self._speak_now(text)
            self.speak_queue.task_done()

    def _speak_now(self, text: str):
        if EDGE_TTS_AVAILABLE and PLAYSOUND_AVAILABLE:
            try:
                voice_id = EDGE_VOICES.get(self.current_voice_name, "ru-RU-DmitryNeural")
                tmp_path = os.path.join(tempfile.gettempdir(), f"assistant_tts_{int(time.time() * 1000)}.mp3")
                asyncio.run(self._edge_tts_save(text, voice_id, tmp_path))
                playsound(tmp_path)
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
                return
            except Exception as e:
                self.log_to_gui(f"edge-tts недоступен ({e}), переключаюсь на SAPI.")

        if self.sapi_voice:
            try:
                self.sapi_voice.Speak(text, 1)
            except Exception:
                pass

    @staticmethod
    async def _edge_tts_save(text, voice_id, path):
        communicate = edge_tts.Communicate(text, voice=voice_id)
        await communicate.save(path)

    def speak(self, text: str):
        print(f"Ассистент: {text}")
        self.log_to_gui(f"Ассистент: {text}")
        self.speak_queue.put(text)

    # ================= GUI =================
    def log_to_gui(self, message: str):
        print(message)
        if hasattr(self, 'root') and self.root and hasattr(self, 'log_area') and self.log_area:
            try:
                self.root.after(0, self._append_to_log, message)
            except Exception:
                pass

    def _append_to_log(self, message: str):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Умный Помощник + Распознавание Лиц (LBPH)")
        self.root.geometry("1000x700")
        self.root.configure(bg="#f0f0f0")

        top_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="Голос:", fg="white", bg="#2c3e50").pack(side=tk.LEFT, padx=10, pady=15)
        self.voice_var = tk.StringVar(value=self.current_voice_name)
        self.voice_combo = ttk.Combobox(top_frame, textvariable=self.voice_var,
                                        values=list(EDGE_VOICES.keys()), state="readonly", width=18)
        self.voice_combo.pack(side=tk.LEFT, padx=5, pady=15)
        self.voice_combo.bind("<<ComboboxSelected>>", self.update_voice_from_gui)

        tk.Label(top_frame, text="Микрофон: Активен", fg="#2ecc71", bg="#2c3e50",
                 font=("Arial", 10, "bold")).pack(side=tk.RIGHT, padx=15, pady=15)

        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        control_frame = tk.LabelFrame(main_frame, text="Управление", bg="#f0f0f0", width=240)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        control_frame.pack_propagate(False)
        self._build_control_buttons(control_frame)

        log_frame = tk.LabelFrame(main_frame, text="Журнал", bg="#f0f0f0")
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', font=("Consolas", 10))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        input_frame = tk.Frame(self.root, bg="#f0f0f0")
        input_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.command_entry = ttk.Entry(input_frame, font=("Arial", 12))
        self.command_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.command_entry.bind("<Return>", lambda event: self.send_command_from_gui())
        ttk.Button(input_frame, text="Отправить", command=self.send_command_from_gui).pack(side=tk.RIGHT)

        self.log_to_gui("Система запущена.")

    def _build_control_buttons(self, parent):
        sections = [
            ("Лица", [
                ("Камера / распознавание", lambda: self.process_command("друг распознавание")),
                ("Добавить лицо", lambda: self.process_command("друг добавь человека")),
                ("Список базы", lambda: self.process_command("друг список")),
                ("Удалить из базы", lambda: self.process_command("друг удали папу")),
                ("Очистить базу", lambda: self.process_command("друг очисти базу")),
            ]),
            ("Информация", [
                ("Погода", lambda: self.process_command("друг погода")),
                ("Который час", lambda: self.process_command("друг который час")),
                ("Шутка", lambda: self.process_command("друг расскажи шутку")),
                ("Интересный факт", lambda: self.process_command("друг интересный факт")),
                ("Цитата дня", lambda: self.process_command("друг цитата дня")),
            ]),
            ("Инструменты", [
                ("Открыть браузер", lambda: self.process_command("друг открой браузер")),
                ("Калькулятор: пример", lambda: self.process_command("друг сколько будет 2 плюс 2")),
            ]),
            ("Прочее", [
                ("Справка", self.show_help),
                ("Стоп / Выход", lambda: self.process_command("друг стоп")),
            ]),
        ]
        for title, buttons in sections:
            box = tk.LabelFrame(parent, text=title, bg="#f0f0f0")
            box.pack(fill=tk.X, padx=8, pady=6)
            for label, cmd in buttons:
                ttk.Button(box, text=label, command=cmd).pack(fill=tk.X, pady=2, padx=4)

    def update_voice_from_gui(self, event=None):
        selected = self.voice_var.get()
        if selected in EDGE_VOICES:
            self.current_voice_name = selected
            self.log_to_gui(f"Голос изменён на: {selected}")

    def send_command_from_gui(self):
        text = self.command_entry.get().strip()
        if text:
            self.log_to_gui(f"Вы: {text}")
            self.command_entry.delete(0, tk.END)
            self.process_command(text)

    def show_help(self):
        messagebox.showinfo("Справка",
                            "Команды (голосом, текстом или кнопками):\n"
                            "- 'Друг, запусти распознавание' / 'камера'\n"
                            "- 'Друг, добавь [имя]'\n"
                            "- 'Друг, удали [имя]'\n"
                            "- 'Друг, очисти базу'\n"
                            "- 'Друг, список' — кто в базе\n"
                            "- 'Друг, который час?'\n"
                            "- 'Друг, погода [в городе]'\n"
                            "- 'Друг, открой браузер' / 'ютуб'\n"
                            "- 'Друг, расскажи шутку'\n"
                            "- 'Друг, интересный факт'\n"
                            "- 'Друг, цитата дня'\n"
                            "- 'Друг, сколько будет 5 плюс 3'\n"
                            "- 'Друг, как дела' / 'кто ты'\n"
                            "- 'Друг, стоп'")

    def load_cyrillic_font(self):
        for path in ["C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/calibri.ttf"]:
            if os.path.exists(path):
                try:
                    return ImageFont.truetype(path, 28)
                except Exception:
                    continue
        return None

    def draw_text_safe(self, img, text, position, color=(0, 255, 255)):
        if self.font:
            img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(img_rgb)
            draw = ImageDraw.Draw(pil_img)
            draw.text(position, text, font=self.font, fill=color[::-1])
            return cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        else:
            latin_text = text.encode('ascii', 'ignore').decode('ascii') or "Face"
            cv2.putText(img, latin_text, position, cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)
            return img

    # ================= VOSK (голосовой ввод) =================
    def init_vosk(self):
        if not VOSK_AVAILABLE or not os.path.exists(MODEL_PATH):
            self.log_to_gui("Vosk-модель не найдена, голосовой ввод отключён.")
            return
        try:
            model = Model(MODEL_PATH)
            self.vosk_recognizer = KaldiRecognizer(model, SAMPLE_RATE)
            self.log_to_gui("Vosk загружена.")
        except Exception as e:
            self.log_to_gui(f"Ошибка Vosk: {e}")

    def start_audio_stream(self):
        if not self.vosk_recognizer: return
        self.listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
        self.listen_thread.start()

    def _listen_loop(self):
        def audio_callback(indata, frames, time_info, status):
            self.audio_queue.put((indata * 32767).astype(np.int16).tobytes())

        try:
            with sd.InputStream(samplerate=SAMPLE_RATE, blocksize=8000, channels=1, callback=audio_callback):
                while self.is_listening:
                    data = self.audio_queue.get()
                    if self.vosk_recognizer.AcceptWaveform(data):
                        text = json.loads(self.vosk_recognizer.Result()).get("text", "").strip().lower()
                        if text:
                            self.log_to_gui(f"Голос: {text}")
                            self.process_command(text)
        except Exception as e:
            self.log_to_gui(f"Ошибка аудио: {e}")

    # ================= ОБРАБОТКА КОМАНД =================
    def process_command(self, command: str):
        command = command.strip().lower()
        if not command.startswith("друг"): return

        command = command[4:].strip()
        if not command:
            self.speak("Слушаю.")
            return

        for phrase, answers in SMALL_TALK.items():
            if phrase in command:
                self.speak(random.choice(answers))
                return

        if any(w in command for w in ['справка', 'помощь']):
            self.show_help()
            return

        if "удали " in command or "забудь " in command:
            name = command.replace("удали ", "").replace("забудь ", "").replace("из базы ", "").strip()
            if name:
                self.delete_person(name)
            else:
                self.speak("Кого именно удалить?")
            return

        if "очисти базу" in command:
            with self.faces_lock:
                self.known_names = {}
                self.next_label_id = 0
                self.model_trained = False
                self.recognizer = cv2.face.LBPHFaceRecognizer_create(radius=2, neighbors=8, grid_x=8, grid_y=8)
            self.save_face_data()
            self.speak("База полностью очищена.")
            return

        if self.is_recognizing_faces:
            if any(w in command for w in ['стоп', 'выход', 'хватит']):
                self.is_recognizing_faces = False
                self.speak("Выхожу из режима камеры.")
            return

        if any(w in command for w in ['распознавание', 'камера', 'узнай']):
            if not self.known_names:
                self.speak("База пуста, добавьте хотя бы одного человека.")
                return
            self.speak("Запускаю распознавание.")
            self.is_recognizing_faces = True
            threading.Thread(target=self.recognize_face_loop, daemon=True).start()
            return

        if any(w in command for w in ['добавь', 'запомни']):
            self.add_person_by_voice()
            return

        if 'список' in command or 'кто в базе' in command:
            self.show_face_list()
            return

        if 'погода' in command:
            self.handle_weather(command)
            return

        if any(w in command for w in ['открой браузер', 'открой сайт']) or command.startswith('открой'):
            self.handle_browser(command)
            return

        if 'шутк' in command or 'анекдот' in command:
            self.speak(random.choice(JOKES))
            return

        if 'факт' in command:
            self.speak(random.choice(FACTS))
            return

        if 'цитат' in command:
            self.speak(random.choice(QUOTES))
            return

        if 'сколько будет' in command or self._looks_like_calc(command):
            self.handle_calculator(command)
            return

        if any(w in command for w in ['время', 'час']):
            self.speak(f"Сейчас {datetime.now().strftime('%H:%M')}")
            return

        if any(w in command for w in ['стоп', 'выход', 'пока']):
            self.speak("До свидания!")
            self.cleanup()
            self.root.quit()
            sys.exit(0)

        self.log_to_gui("Команда не распознана.")

    # ================= НОВЫЕ ФУНКЦИИ =================
    def handle_weather(self, command: str):
        city = "Warsaw"
        for prefix in ["погода в ", "погода для ", "погода "]:
            if prefix in command:
                rest = command.split(prefix, 1)[1].strip()
                if rest: city = rest
                break

        def worker():
            try:
                url = f"https://wttr.in/{city}?format=%C,+%t,+ощущается+как+%f&lang=ru"
                resp = requests.get(url, timeout=6)
                if resp.status_code == 200 and resp.text.strip():
                    self.speak(f"Погода в {city}: {resp.text.strip()}")
                else:
                    self.speak("Не удалось получить данные о погоде.")
            except Exception as e:
                self.log_to_gui(f"Ошибка погоды: {e}")
                self.speak("Не получилось узнать погоду. Проверьте интернет-соединение.")

        threading.Thread(target=worker, daemon=True).start()

    def handle_browser(self, command: str):
        sites = {
            "ютуб": "https://youtube.com", "youtube": "https://youtube.com",
            "гугл": "https://google.com", "google": "https://google.com",
            "почту": "https://mail.google.com", "почта": "https://mail.google.com",
        }
        for key, url in sites.items():
            if key in command:
                webbrowser.open(url)
                self.speak(f"Открываю {key}.")
                return
        webbrowser.open("https://google.com")
        self.speak("Открываю браузер.")

    @staticmethod
    def _looks_like_calc(command: str) -> bool:
        ops = ['плюс', 'минус', 'умножить', 'разделить', '+', '-', '*', '/']
        return any(op in command for op in ops) and any(ch.isdigit() for ch in command)

    def handle_calculator(self, command: str):
        words_to_ops = {
            'плюс': '+', 'минус': '-', 'умножить на': '*', 'умножить': '*',
            'разделить на': '/', 'разделить': '/',
        }
        text = command
        for word, op in words_to_ops.items():
            text = text.replace(word, op)

        import re
        match = re.search(r'(-?\d+(?:[.,]\d+)?)\s*([+\-*/])\s*(-?\d+(?:[.,]\d+)?)', text)
        if not match:
            self.speak("Не удалось разобрать пример.")
            return
        a = float(match.group(1).replace(',', '.'))
        op = match.group(2)
        b = float(match.group(3).replace(',', '.'))
        try:
            if op == '+':
                result = a + b
            elif op == '-':
                result = a - b
            elif op == '*':
                result = a * b
            elif op == '/':
                if b == 0:
                    self.speak("На ноль делить нельзя.")
                    return
                result = a / b
            else:
                self.speak("Не понял операцию.")
                return
            self.speak(f"Результат: {result:g}")
        except Exception:
            self.speak("Не удалось вычислить пример.")

    # ================= РАСПОЗНАВАНИЕ ЛИЦ (LBPH) =================
    def load_face_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'rb') as f:
                    saved = pickle.load(f)
                self.known_names = saved.get('names', {})
                samples = saved.get('samples', {})
                self.next_label_id = (max(self.known_names.keys()) + 1) if self.known_names else 0

                all_images, all_labels = [], []
                for label, imgs in samples.items():
                    for img in imgs:
                        all_images.append(img)
                        all_labels.append(label)

                if all_images:
                    self.recognizer.train(all_images, np.array(all_labels))
                    self.model_trained = True
                    self._saved_samples = samples
                else:
                    self._saved_samples = {}
                self.log_to_gui(f"Загружено лиц: {len(self.known_names)}.")
            except Exception as e:
                self.log_to_gui(f"Ошибка загрузки базы лиц: {e}")
                self.known_names = {}
                self._saved_samples = {}
        else:
            self._saved_samples = {}

    def save_face_data(self):
        with self.faces_lock:
            data = {'names': self.known_names, 'samples': getattr(self, '_saved_samples', {})}
            with open(DATA_FILE, 'wb') as f:
                pickle.dump(data, f)

    def delete_person(self, name):
        name = name.lower().strip()
        with self.faces_lock:
            label_to_remove = next((lbl for lbl, n in self.known_names.items() if n.lower() == name), None)
            if label_to_remove is None:
                self.speak(f"Не нашёл {name} в базе.")
                return

            del self.known_names[label_to_remove]
            if hasattr(self, '_saved_samples') and label_to_remove in self._saved_samples:
                del self._saved_samples[label_to_remove]

            self.recognizer = cv2.face.LBPHFaceRecognizer_create(radius=2, neighbors=8, grid_x=8, grid_y=8)
            all_images, all_labels = [], []
            for label, imgs in getattr(self, '_saved_samples', {}).items():
                for img in imgs:
                    all_images.append(img)
                    all_labels.append(label)
            if all_images:
                self.recognizer.train(all_images, np.array(all_labels))
                self.model_trained = True
            else:
                self.model_trained = False

        self.save_face_data()
        self.speak(f"Я удалил {name} из базы.")
        self.log_to_gui(f"Удалён: {name}")

    def get_face_features(self, face_img):
        gray = cv2.cvtColor(face_img, cv2.COLOR_BGR2GRAY)
        gray = cv2.resize(gray, (200, 200))
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        gray = clahe.apply(gray)
        return gray

    def add_person_by_voice(self):
        name = simpledialog_askstring("Добавление", "Введите имя человека:")
        if name:
            self.add_person(name.strip().capitalize())
        else:
            self.speak("Отменено.")

    def add_person(self, name):
        if self.is_recognizing_faces:
            self.speak("Сначала остановите режим распознавания.")
            return

        self.speak(f"Добавляем {name}. Смотрите в камеру.")
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.speak("Не удалось открыть камеру.")
            return

        collected = []
        try:
            while len(collected) < 25 and self.is_listening:
                ret, frame = cap.read()
                if not ret: break
                frame = cv2.flip(frame, 1)
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self.face_cascade.detectMultiScale(gray, 1.1, 5, minSize=(100, 100))
                for (x, y, w, h) in faces:
                    feat = self.get_face_features(frame[y:y + h, x:x + w])
                    collected.append(feat)
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    time.sleep(0.08)

                # ИСПРАВЛЕНО: Используем draw_text_safe для кириллицы
                frame = self.draw_text_safe(frame, f"Снимков: {len(collected)}/25", (15, 35), (0, 255, 255))

                # ИСПРАВЛЕНО: Заголовок окна на английском, чтобы не было ???
                cv2.imshow("Add Face (q - cancel)", frame)
                if cv2.waitKey(1) & 0xFF == ord('q'): break
        finally:
            cap.release()
            cv2.destroyAllWindows()

        if len(collected) >= 8:
            with self.faces_lock:
                existing_label = next((lbl for lbl, n in self.known_names.items() if n.lower() == name.lower()), None)
                if existing_label is not None:
                    label = existing_label
                else:
                    label = self.next_label_id
                    self.next_label_id += 1
                    self.known_names[label] = name

                if not hasattr(self, '_saved_samples'): self._saved_samples = {}
                self._saved_samples.setdefault(label, [])
                self._saved_samples[label].extend(collected)

                all_images, all_labels = [], []
                for lbl, imgs in self._saved_samples.items():
                    for img in imgs:
                        all_images.append(img)
                        all_labels.append(lbl)
                self.recognizer.train(all_images, np.array(all_labels))
                self.model_trained = True

            self.save_face_data()
            self.speak(f"Запомнил, как {name}. Данные сохранены.")
        else:
            self.speak("Не удалось сделать достаточно снимков. Попробуйте ещё раз.")

    def recognize_face_loop(self):
        if not self.known_names or not self.model_trained:
            self.speak("База пуста.")
            self.is_recognizing_faces = False
            return

        log_path = os.path.join(LOG_DIR, f"recognition_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        try:
            self.recognition_log_file = open(log_path, 'w', encoding='utf-8')
            self.recognition_log_file.write(f"Лог распознавания. Начало сессии: {datetime.now()}\n")
            self.recognition_log_file.write("Алгоритм: " + " | ".join(ALGORITHM_DESCRIPTION) + "\n\n")
        except Exception as e:
            self.log_to_gui(f"Не удалось открыть лог-файл: {e}")
            self.recognition_log_file = None

        cap = cv2.VideoCapture(0)
        last_time = 0
        try:
            while self.is_recognizing_faces:
                ret, frame = cap.read()
                if not ret: break
                frame = cv2.flip(frame, 1)
                gray_full = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self.face_cascade.detectMultiScale(gray_full, 1.1, 5, minSize=(100, 100))

                now = time.time()
                for (x, y, w, h) in faces:
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

                    if now - last_time > 1.0:
                        last_time = now
                        feat = self.get_face_features(frame[y:y + h, x:x + w])

                        with self.faces_lock:
                            try:
                                label, distance = self.recognizer.predict(feat)
                            except cv2.error:
                                continue
                            name = self.known_names.get(label)

                        confidence_pct = int(max(0, min(100, 100 - distance)))
                        is_match = (name is not None) and (distance < LBPH_DISTANCE_THRESHOLD)

                        if is_match:
                            color = (0, 255, 0)
                            cv2.rectangle(frame, (x, y), (x + w, y + h), color, 3)
                            frame = self.draw_text_safe(frame, name, (x, y - 40), color)
                            self._log_recognition(name, confidence_pct, distance)

                            self.last_name = name
                            self.last_confidence = confidence_pct
                            self.last_distance = distance

                            if name != self.last_greeted_name or (now - self.last_greeted_time > self.speak_cooldown):
                                self.speak(f"Привет, {name}!")
                                self.last_greeted_name = name
                                self.last_greeted_time = now
                        else:
                            color = (0, 0, 255)
                            cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)
                            frame = self.draw_text_safe(frame, "Неизвестный", (x, y - 40), color)
                            if now - self.last_spoken_time > self.speak_cooldown:
                                self.last_spoken_time = now
                                self.speak(UNKNOWN_PHRASE)

                            self._log_recognition("Unknown", confidence_pct, distance)

                            self.last_name = "Неизвестный"
                            self.last_confidence = confidence_pct
                            self.last_distance = distance

                frame = self.draw_text_safe(frame, f"Имя: {self.last_name}", (15, 35), (0, 255, 255))
                frame = self.draw_text_safe(frame, f"Уверенность: {self.last_confidence}%", (15, 75), (0, 255, 255))
                frame = self.draw_text_safe(frame, f"Дистанция: {self.last_distance:.1f}", (15, 115), (0, 255, 255))

                h_frame, w_frame, _ = frame.shape
                start_y = h_frame - 20 * (len(ALGORITHM_DESCRIPTION) + 1)
                for i, line in enumerate(ALGORITHM_DESCRIPTION):
                    cv2.putText(frame, line, (w_frame - 480, start_y + i * 20),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255, 255, 255), 1)

                # ИСПРАВЛЕНО: Заголовок окна на английском
                cv2.imshow("Recognition (q - exit)", frame)
                if cv2.waitKey(1) & 0xFF == ord('q'): break
        finally:
            cap.release()
            cv2.destroyAllWindows()
            self.is_recognizing_faces = False
            if self.recognition_log_file:
                self.recognition_log_file.write(f"\nКонец сессии: {datetime.now()}\n")
                self.recognition_log_file.close()
                self.recognition_log_file = None
                self.log_to_gui(f"Лог распознавания сохранён: {log_path}")

    def _log_recognition(self, name, confidence_pct, distance):
        line = (f"{datetime.now().strftime('%H:%M:%S')} | "
                f"Имя: {name} | Confidence: {confidence_pct}% | Distance: {distance:.2f}")
        if self.recognition_log_file:
            try:
                self.recognition_log_file.write(line + "\n")
                self.recognition_log_file.flush()
            except Exception:
                pass

    def show_face_list(self):
        if self.known_names:
            self.speak("В базе: " + ", ".join(self.known_names.values()))
        else:
            self.speak("База пуста.")

    def cleanup(self):
        self.is_listening = False
        self.speak_queue.put(None)
        if self.recognition_log_file:
            try:
                self.recognition_log_file.close()
            except Exception:
                pass


def simpledialog_askstring(title, prompt):
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.geometry("300x100")
    dialog.grab_set()
    result = tk.StringVar()
    tk.Label(dialog, text=prompt).pack(pady=5)
    entry = ttk.Entry(dialog, textvariable=result, width=30)
    entry.pack(pady=5)
    entry.focus()

    def on_ok(): dialog.destroy()

    ttk.Button(dialog, text="OK", command=on_ok).pack(pady=5)
    dialog.wait_window()
    return result.get()


def main():
    assistant = None
    try:
        assistant = SmartAssistant()
        assistant.speak("Здравствуйте. Я голосовой помощник Друг. Для команд начинайте фразу со слова Друг.")

        def on_close():
            assistant.cleanup()
            assistant.root.destroy()

        assistant.root.protocol("WM_DELETE_WINDOW", on_close)
        assistant.root.mainloop()
    except KeyboardInterrupt:
        pass
    finally:
        if assistant is not None:
            assistant.cleanup()
            cv2.destroyAllWindows()


if __name__ == "__main__":
    main()