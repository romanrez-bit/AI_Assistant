# -*- coding: utf-8 -*-
"""
Умный голосовой помощник с распознаванием лиц.

Изменения в этой версии (v5):
- Распознавание лиц переведено на LBPH (cv2.face.LBPHFaceRecognizer) —
  это настоящий ML-алгоритм на основе текстурных гистограмм, входит в
  opencv-contrib-python, не требует dlib/CMake.
- Озвучка переведена на Microsoft Edge TTS (библиотека edge-tts) —
  два приятных нейросетевых голоса: ru-RU-DmitryNeural (мужской) и
  ru-RU-SvetlanaNeural (женский). Нужен интернет. SAPI оставлен как
  запасной вариант, если edge-tts недоступен (нет сети).
- Новые команды: погода (wttr.in, без API-ключа), открыть браузер/сайт,
  шутка, интересный факт, калькулятор, цитата дня.
- В окне распознавания: вероятность (confidence, %) — верхний левый
  угол; описание алгоритма — нижний правый угол; всё это построчно
  пишется в лог-файл до момента закрытия окна камеры.
- GUI: все команды доступны как кнопки в левой панели; всё это же
  работает и голосом, и через текстовый ввод.

Установка зависимостей (Windows, PyCharm, терминал):
    pip install opencv-contrib-python pillow pywin32 vosk sounddevice
    pip install edge-tts playsound3 requests

Модель Vosk должна лежать в папке vosk-model-small-ru-0.22 рядом со
скриптом (распаковать архив с https://alphacephei.com/vosk/models).
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

# Пороговое значение LBPH (расстояние; меньше = увереннее).
# Подобрано эмпирически: < 60 — уверенное совпадение.
LBPH_DISTANCE_THRESHOLD = 60

# Описание алгоритма для отображения в правом нижнем углу
ALGORITHM_DESCRIPTION = [
    "ALGORITHM: LBPH (Local Binary Patterns Histograms)",
    "1. GRAYSCALE & RESIZE (200x200)",
    "2. CLAHE HISTOGRAM EQUALIZATION",
    "3. LBP: PIXEL VS NEIGHBORS (RADIUS=2, NEIGHBORS=8)",
    "4. HISTOGRAMS PER 8x8 GRID CELL",
    "5. DISTANCE = CHI-SQUARE BETWEEN HISTOGRAMS",
    "6. CONFIDENCE % = 100 - MIN(DISTANCE, 100)",
]

# Голоса edge-tts (приятные нейро-голоса). Профилей больше нет — один
# ассистент ("Друг"), голос выбирается вручную в GUI и не меняется
# автоматически по распознанному лицу.
EDGE_VOICES = {
    "Дмитрий (муж.)": "ru-RU-DmitryNeural",
    "Светлана (жен.)": "ru-RU-SvetlanaNeural",
}
DEFAULT_VOICE_NAME = "Дмитрий (муж.)"

# Ключевое слово, с которого должна начинаться голосовая команда.
# В тексте/кнопках GUI это слово не требуется.
WAKE_WORD = "друг"

# Имена для отображения в окне распознавания. ИИ распознаёт по базе лиц
# (известные имена), а для неизвестного лица показывается это слово.
UNKNOWN_LABEL = "Неизвестен"

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
    "Самая короткая война в истории длилась 38 минут — между Занзибаром и Великобританией.",
    "Улитки могут спать до трёх лет подряд.",
]

QUOTES = [
    "Лучший способ начать — это перестать говорить и начать делать.",
    "Не бойтесь медленного прогресса, бойтесь стоять на месте.",
    "Успех — это сумма маленьких усилий, повторяемых день за днём.",
    "Сделай сегодня то, на что другие не способны, и завтра сможешь то, на что другие не способны.",
    "Трудности делают нас сильнее, если мы не позволяем им нас остановить.",
    "Лучшее время посадить дерево было двадцать лет назад. Следующее лучшее время — сейчас.",
]

# Эмпатичные реплики на простые личные вопросы. Ключ — список фраз-триггеров,
# значение — варианты ответа (выбирается случайно, чтобы не звучало роботизированно).
EMPATHY_RESPONSES = [
    (["как дела", "как у тебя дела", "как жизнь"],
     ["Спасибо, что спросил! У меня всё отлично, я готов помогать.",
      "Всё хорошо, работаю в обычном режиме и рад быть полезным.",
      "Дела хорошо, особенно когда есть с кем поговорить!"]),
    (["как настроение", "какое у тебя настроение"],
     ["Настроение бодрое, готов к работе!",
      "Настроение хорошее, спасибо, что интересуешься.",
      "Чувствую себя отлично, давай чем-нибудь займёмся."]),
    (["как ты", "ты как"],
     ["Я в порядке, спасибо! А у тебя как дела?",
      "Всё хорошо, спасибо за заботу."]),
    (["что делаешь", "чем занимаешься", "что нового"],
     ["Слушаю тебя и готов помочь — с погодой, лицами, шутками, чем угодно.",
      "Жду твоих команд, ничего не пропускаю."]),
    (["спасибо"],
     ["Пожалуйста! Обращайся в любое время.",
      "Рад был помочь!",
      "Всегда пожалуйста."]),
    (["устал", "устала", "грустно", "плохое настроение", "плохо"],
     ["Жаль это слышать. Если хочешь — расскажу шутку или интересный факт, чтобы немного отвлечься.",
      "Понимаю. Иногда помогает небольшая пауза. Могу рассказать что-нибудь интересное, если хочешь."]),
    (["я тебя люблю", "ты молодец", "ты хороший"],
     ["Очень приятно это слышать, спасибо!",
      "Спасибо тебе! Стараюсь быть полезным."]),
]


class SmartAssistant:
    def __init__(self):
        self.is_recognizing_faces = False
        self.is_listening = True

        self.audio_queue = queue.Queue()
        self.speak_queue = queue.Queue()
        self.listen_thread = None

        # Текущий голос edge-tts (короткое имя из EDGE_VOICES). Один голос
        # для всего ассистента, профилей и авто-смены по лицу больше нет.
        self.current_voice_name = DEFAULT_VOICE_NAME

        # Лок для безопасного доступа к базе лиц из разных потоков
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
        self.known_names = {}   # label_id (int) -> имя
        self.next_label_id = 0
        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
        self.font = self.load_cyrillic_font()
        self.last_spoken_name = ""
        self.last_spoken_time = 0
        self.speak_cooldown = 3.0
        self.recognizer = cv2.face.LBPHFaceRecognizer_create(radius=2, neighbors=8,
                                                               grid_x=8, grid_y=8)
        self.model_trained = False
        self.recognition_log_file = None
        # Последний результат распознавания на лицо — чтобы рисовать
        # рамку/подпись/вероятность на КАЖДОМ кадре, а не только в момент
        # пересчёта (иначе текст мигает/пропадает между обновлениями).
        # Формат: {face_key: {"name", "confidence_pct", "distance", "color",
        #                      "box", "last_update"}}
        self.last_results = {}
        self.load_face_data()

        # 4. Голосовой ввод
        self.vosk_recognizer = None
        self.init_vosk()
        self.start_audio_stream()

        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR)

    # ================= ИНИЦИАЛИЗАЦИЯ ГОЛОСА =================
    def init_sapi_speaker(self):
        """Запасной голос через Windows SAPI (используется только если
        edge-tts недоступен — например, нет интернета)."""
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
        """Озвучивает текст: пытается edge-tts, при ошибке падает на SAPI."""
        if EDGE_TTS_AVAILABLE and PLAYSOUND_AVAILABLE:
            try:
                voice_id = EDGE_VOICES.get(self.current_voice_name, "ru-RU-DmitryNeural")
                tmp_path = os.path.join(tempfile.gettempdir(),
                                         f"assistant_tts_{int(time.time()*1000)}.mp3")
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
        self.root.title("Друг — Умный Помощник с Распознаванием Лиц (LBPH)")
        self.root.geometry("1000x700")
        self.root.configure(bg="#f0f0f0")

        # Верхняя панель
        top_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="🤖 Друг", fg="white", bg="#2c3e50",
                 font=("Arial", 13, "bold")).pack(side=tk.LEFT, padx=15, pady=15)

        tk.Label(top_frame, text="Голос:", fg="white", bg="#2c3e50").pack(
            side=tk.LEFT, padx=(20, 5), pady=15)
        self.voice_var = tk.StringVar(value=self.current_voice_name)
        self.voice_combo = ttk.Combobox(top_frame, textvariable=self.voice_var,
                                         values=list(EDGE_VOICES.keys()),
                                         state="readonly", width=18)
        self.voice_combo.pack(side=tk.LEFT, padx=5, pady=15)
        self.voice_combo.bind("<<ComboboxSelected>>", self.update_voice_from_gui)

        tk.Label(top_frame, text="Микрофон: Активен", fg="#2ecc71", bg="#2c3e50",
                 font=("Arial", 10, "bold")).pack(side=tk.RIGHT, padx=15, pady=15)

        # Основная область
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Левая панель управления (все команды здесь)
        control_frame = tk.LabelFrame(main_frame, text="Управление", bg="#f0f0f0", width=240)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        control_frame.pack_propagate(False)

        self._build_control_buttons(control_frame)

        # Лог справа
        log_frame = tk.LabelFrame(main_frame, text="Журнал", bg="#f0f0f0")
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled',
                                                   font=("Consolas", 10))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Ввод команд
        input_frame = tk.Frame(self.root, bg="#f0f0f0")
        input_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.command_entry = ttk.Entry(input_frame, font=("Arial", 12))
        self.command_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.command_entry.bind("<Return>", lambda event: self.send_command_from_gui())
        ttk.Button(input_frame, text="Отправить", command=self.send_command_from_gui).pack(
            side=tk.RIGHT)

        self.log_to_gui("Система запущена. Ассистент «Друг» к работе готов.")

    def _build_control_buttons(self, parent):
        """Все команды ассистента доступны как кнопки в левой панели."""
        sections = [
            ("Лица", [
                ("Камера / распознавание", lambda: self.process_command("распознавание")),
                ("Добавить лицо", lambda: self.process_command("добавь человека")),
                ("Список базы", lambda: self.process_command("список")),
                ("Удалить из базы", lambda: self.delete_person_by_dialog()),
                ("Очистить базу", lambda: self.process_command("очисти базу")),
            ]),
            ("Информация", [
                ("Погода", lambda: self.process_command("погода")),
                ("Который час", lambda: self.process_command("который час")),
                ("Шутка", lambda: self.process_command("расскажи шутку")),
                ("Интересный факт", lambda: self.process_command("интересный факт")),
                ("Цитата дня", lambda: self.process_command("цитата дня")),
            ]),
            ("Инструменты", [
                ("Открыть браузер", lambda: self.process_command("открой браузер")),
                ("Калькулятор: пример", lambda: self.process_command("сколько будет 2 плюс 2")),
            ]),
            ("Прочее", [
                ("Как дела?", lambda: self.process_command("как дела")),
                ("Как настроение?", lambda: self.process_command("как настроение")),
                ("Справка", self.show_help),
                ("Стоп / Выход", lambda: self.process_command("стоп")),
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
            "Голосовые команды начинаются со слова «Друг», например:\n"
            "«Друг, погода», «Друг, расскажи шутку».\n"
            "В тексте и кнопках слово «Друг» не нужно.\n\n"
            "Доступные команды:\n"
            "- 'Запусти распознавание' / 'камера'\n"
            "- 'Добавь [имя]'\n"
            "- 'Удали [имя]'\n"
            "- 'Очисти базу'\n"
            "- 'Список' — кто в базе\n"
            "- 'Который час?'\n"
            "- 'Погода [в городе]'\n"
            "- 'Открой браузер' / 'открой ютуб'\n"
            "- 'Расскажи шутку'\n"
            "- 'Интересный факт'\n"
            "- 'Цитата дня'\n"
            "- 'Сколько будет 5 плюс 3'\n"
            "- 'Как дела?', 'Как настроение?' — простые вопросы\n"
            "- 'Стоп'")

    def delete_person_by_dialog(self):
        name = simpledialog_askstring("Удаление", "Введите имя человека для удаления:")
        if name:
            self.delete_person(name.strip())
        else:
            self.speak("Отменено.")

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
        if not self.vosk_recognizer:
            return
        self.listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
        self.listen_thread.start()

    def _listen_loop(self):
        def audio_callback(indata, frames, time_info, status):
            self.audio_queue.put((indata * 32767).astype(np.int16).tobytes())
        try:
            with sd.InputStream(samplerate=SAMPLE_RATE, blocksize=8000, channels=1,
                                 callback=audio_callback):
                while self.is_listening:
                    data = self.audio_queue.get()
                    if self.vosk_recognizer.AcceptWaveform(data):
                        text = json.loads(self.vosk_recognizer.Result()).get("text", "").strip().lower()
                        if text:
                            self.log_to_gui(f"Голос (распознано): {text}")
                            command = self._strip_wake_word(text)
                            if command is not None:
                                self.process_command(command)
                            # Если в фразе не было слова "друг" — игнорируем
                            # её как случайную/фоновую речь.
        except Exception as e:
            self.log_to_gui(f"Ошибка аудио: {e}")

    @staticmethod
    def _strip_wake_word(text: str):
        """Голосовые команды должны начинаться с ключевого слова (WAKE_WORD,
        по умолчанию «друг»). Возвращает остаток фразы без этого слова,
        либо None, если ключевого слова не было (фраза игнорируется)."""
        text = text.strip().lower()
        if text == WAKE_WORD or text.startswith(WAKE_WORD + " ") or \
           text.startswith(WAKE_WORD + ","):
            rest = text[len(WAKE_WORD):].lstrip(", ").strip()
            return rest
        return None

    # ================= ОБРАБОТКА КОМАНД =================
    def process_command(self, command: str):
        command = command.strip().lower()
        if not command:
            return

        if any(w in command for w in ['справка', 'помощь']):
            self.show_help()
            return

        if "удали" in command or "забудь" in command:
            name = command.replace("удали", "").replace("забудь", "").replace("из базы", "").strip()
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
                self.recognizer = cv2.face.LBPHFaceRecognizer_create(radius=2, neighbors=8,
                                                                       grid_x=8, grid_y=8)
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

        if any(w in command for w in ['открой браузер', 'открой сайт']) or \
           command.startswith('открой '):
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

        empathy_reply = self._match_empathy(command)
        if empathy_reply:
            self.speak(empathy_reply)
            return

        self.log_to_gui("Команда не распознана.")

    @staticmethod
    def _match_empathy(command: str):
        """Ищет совпадение с простыми личными вопросами (как дела, как
        настроение и т.п.) и возвращает случайный тёплый ответ, либо None."""
        for triggers, replies in EMPATHY_RESPONSES:
            if any(trigger in command for trigger in triggers):
                return random.choice(replies)
        return None

    # ================= НОВЫЕ ФУНКЦИИ =================
    def handle_weather(self, command: str):
        """Погода через wttr.in, без API-ключа."""
        city = "Warsaw"
        for prefix in ["погода в", "погода для", "погода"]:
            if prefix in command:
                rest = command.split(prefix, 1)[1].strip()
                if rest:
                    city = rest
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
            "ютуб": "https://youtube.com",
            "youtube": "https://youtube.com",
            "гугл": "https://google.com",
            "google": "https://google.com",
            "почту": "https://mail.google.com",
            "почта": "https://mail.google.com",
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
            'плюс': '+', 'минус': '-',
            'умножить на': '*', 'умножить': '*',
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
            result_str = f"{result:g}"
            self.speak(f"Результат: {result_str}")
        except Exception:
            self.speak("Не удалось вычислить пример.")

    # ================= РАСПОЗНАВАНИЕ ЛИЦ (LBPH) =================
    def load_face_data(self):
        """Загружает сохранённые лица и переобучает LBPH-модель."""
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'rb') as f:
                    saved = pickle.load(f)
                # saved: {'names': {label: name}, 'samples': {label: [images]}}
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
            label_to_remove = next((lbl for lbl, n in self.known_names.items()
                                     if n.lower() == name), None)
            if label_to_remove is None:
                self.speak(f"Не нашёл {name} в базе.")
                return

            del self.known_names[label_to_remove]
            if hasattr(self, '_saved_samples') and label_to_remove in self._saved_samples:
                del self._saved_samples[label_to_remove]

            # Переобучаем модель на оставшихся данных
            self.recognizer = cv2.face.LBPHFaceRecognizer_create(radius=2, neighbors=8,
                                                                   grid_x=8, grid_y=8)
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
        """Подготовка изображения лица для LBPH: grayscale, resize, CLAHE."""
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
                if not ret:
                    break
                frame = cv2.flip(frame, 1)
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self.face_cascade.detectMultiScale(gray, 1.1, 5, minSize=(100, 100))
                for (x, y, w, h) in faces:
                    feat = self.get_face_features(frame[y:y + h, x:x + w])
                    collected.append(feat)
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    time.sleep(0.08)
                frame = self.draw_text_safe(frame, f"Снимков: {len(collected)}/25",
                                             (15, 15), (0, 255, 255))
                cv2.imshow("Добавление лица (q - отмена)", frame)
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
        finally:
            cap.release()
            cv2.destroyAllWindows()

        if len(collected) >= 8:
            with self.faces_lock:
                # Если имя уже есть — дополняем его данные, иначе создаём новую метку
                existing_label = next((lbl for lbl, n in self.known_names.items()
                                       if n.lower() == name.lower()), None)
                if existing_label is not None:
                    label = existing_label
                else:
                    label = self.next_label_id
                    self.next_label_id += 1
                    self.known_names[label] = name

                if not hasattr(self, '_saved_samples'):
                    self._saved_samples = {}
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
            self.speak(f"Запомнил, как {name}. Готов узнавать при следующей встрече.")
        else:
            self.speak("Не удалось сделать достаточно снимков. Попробуйте ещё раз.")

    def recognize_face_loop(self):
        if not self.known_names or not self.model_trained:
            self.speak("База пуста.")
            self.is_recognizing_faces = False
            return

        # Открываем лог-файл на время сессии распознавания
        log_path = os.path.join(
            LOG_DIR, f"recognition_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        try:
            self.recognition_log_file = open(log_path, 'w', encoding='utf-8')
            self.recognition_log_file.write(
                f"Лог распознавания. Начало сессии: {datetime.now()}\n")
            self.recognition_log_file.write("Алгоритм: " + " | ".join(ALGORITHM_DESCRIPTION) + "\n\n")
        except Exception as e:
            self.log_to_gui(f"Не удалось открыть лог-файл: {e}")
            self.recognition_log_file = None

        # last_result хранит последний посчитанный результат распознавания.
        # Используется, чтобы вероятность/имя оставались на экране ПОСТОЯННО
        # (пока открыто окно камеры), а не мигали между кадрами — даже если
        # детектор лиц на долю секунды не нашёл лицо.
        self.last_results = {"main": None}

        cap = cv2.VideoCapture(0)
        try:
            while self.is_recognizing_faces:
                ret, frame = cap.read()
                if not ret:
                    break
                frame = cv2.flip(frame, 1)
                gray_full = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self.face_cascade.detectMultiScale(gray_full, 1.1, 5, minSize=(100, 100))

                now = time.time()

                if len(faces) > 0:
                    # Берём самое крупное лицо в кадре (наиболее вероятный
                    # "главный" объект, если в кадре несколько лиц).
                    x, y, w, h = max(faces, key=lambda f: f[2] * f[3])

                    feat = self.get_face_features(frame[y:y + h, x:x + w])
                    with self.faces_lock:
                        try:
                            label, distance = self.recognizer.predict(feat)
                            name = self.known_names.get(label)
                        except cv2.error:
                            label, distance, name = None, 999.0, None

                    confidence_pct = int(max(0, min(100, 100 - distance)))
                    is_match = (name is not None) and (distance < LBPH_DISTANCE_THRESHOLD)
                    display_name = name if is_match else UNKNOWN_LABEL

                    self.last_results["main"] = {
                        "name": display_name,
                        "is_match": is_match,
                        "confidence_pct": confidence_pct,
                        "distance": distance,
                        "box": (x, y, w, h),
                    }

                    if is_match:
                        if name != self.last_spoken_name:
                            self.last_spoken_name = name
                            self.speak(f"Это {name}!")
                    else:
                        self.last_spoken_name = ""
                        if now - self.last_spoken_time > self.speak_cooldown:
                            self.last_spoken_time = now
                            self.speak(UNKNOWN_PHRASE)

                    self._log_recognition(display_name, confidence_pct, distance)

                # Отрисовываем последний известный результат на каждом
                # кадре — вероятность не мигает, пока окно камеры открыто.
                result = self.last_results.get("main")
                if result:
                    x, y, w, h = result["box"]
                    color = (0, 255, 0) if result["is_match"] else (0, 0, 255)
                    thickness = 3 if result["is_match"] else 2
                    cv2.rectangle(frame, (x, y), (x + w, y + h), color, thickness)
                    # Имя рисуем кириллицей через PIL (cv2.putText не умеет кириллицу)
                    frame = self.draw_text_safe(frame, result["name"], (x, max(0, y - 40)), color)

                    # Вероятность — верхний левый угол, на экране постоянно.
                    frame = self.draw_text_safe(
                        frame, f"Вероятность: {result['confidence_pct']}%",
                        (15, 15), (0, 255, 255))
                    frame = self.draw_text_safe(
                        frame, f"Расстояние: {result['distance']:.1f}",
                        (15, 50), (0, 255, 255))

                # Описание алгоритма — нижний правый угол, видно всегда,
                # независимо от того, найдено лицо в данном кадре или нет.
                h_frame, w_frame, _ = frame.shape
                start_y = h_frame - 20 * (len(ALGORITHM_DESCRIPTION) + 1)
                for i, line in enumerate(ALGORITHM_DESCRIPTION):
                    cv2.putText(frame, line, (w_frame - 480, start_y + i * 20),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255, 255, 255), 1)

                cv2.imshow("Распознавание (q - выход)", frame)
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
        finally:
            cap.release()
            cv2.destroyAllWindows()
            self.is_recognizing_faces = False
            self.last_results = {}
            if self.recognition_log_file:
                self.recognition_log_file.write(f"\nКонец сессии: {datetime.now()}\n")
                self.recognition_log_file.close()
                self.recognition_log_file = None
                self.log_to_gui(f"Лог распознавания сохранён: {log_path}")

    def _log_recognition(self, name, confidence_pct, distance):
        line = (f"{datetime.now().strftime('%H:%M:%S')} | "
                f"Имя: {name} | Вероятность: {confidence_pct}% | Расстояние: {distance:.2f}")
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
    ttk.Button(dialog, text="OK", command=dialog.destroy).pack(pady=5)
    dialog.wait_window()
    return result.get()


def main():
    assistant = None
    try:
        assistant = SmartAssistant()
        assistant.speak("Система запущена. Я твой друг и готов помогать.")
        assistant.root.protocol("WM_DELETE_WINDOW",
                                 lambda: [assistant.cleanup(), assistant.root.quit(), sys.exit(0)])
        assistant.root.mainloop()
    except KeyboardInterrupt:
        pass
    finally:
        if assistant is not None:
            assistant.cleanup()
        cv2.destroyAllWindows()


if __name__ == "__main__":
    main()