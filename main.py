# -*- coding: utf-8 -*-
"""
Умный голосовой помощник «Друг» с распознаванием лиц и эмоций.

Возможности:
- Распознавание лиц: LBPH (cv2.face.LBPHFaceRecognizer) — настоящий ML-
  алгоритм на основе текстурных гистограмм, входит в opencv-contrib-python,
  не требует dlib/CMake.
- Распознавание эмоций: DeepFace (deep learning, TensorFlow). Эмоция
  показывается в окне камеры постоянно (с эмодзи) и озвучивается по
  команде «друг, какая у меня эмоция».
- Озвучка: Microsoft Edge TTS (edge-tts) — приятные нейро-голоса Дмитрий
  и Светлана. SAPI как запасной вариант, если нет интернета.
- Умные ответы: при нераспознанной команде вопрос уходит в локальную
  модель GPT4All (если включён её API-сервер), и ассистент отвечает как
  чат-бот.
- Команды: погода, браузер, шутка, факт, цитата дня, калькулятор,
  эмпатичные ответы (как дела/настроение), время, выход.
- Голосовые команды начинаются со слова «друг». Кнопки и текстовый ввод
  работают без него.

Установка зависимостей (Windows, PyCharm, терминал):
    pip install opencv-contrib-python pillow pywin32 vosk sounddevice
    pip install edge-tts playsound3 requests
    pip install deepface tf-keras       (для распознавания эмоций)

Vosk-модель: папка vosk-model-small-ru-0.22 рядом со скриптом
    (скачать с https://alphacephei.com/vosk/models).

GPT4All (умные ответы): установить приложение, загрузить модель, затем
    Settings -> Application -> Advanced -> включить «Enable Local API Server»
    (порт 4891). Без этого ассистент просто скажет, что не понял команду.

DeepFace при первом запуске скачивает веса моделей (~несколько сотен МБ)
    с github — нужен интернет один раз.
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

# --- DeepFace (распознавание эмоций). Тяжёлая библиотека (TensorFlow);
# импортируем лениво при первом использовании, чтобы не замедлять старт.
DEEPFACE_AVAILABLE = None  # None = ещё не проверяли, True/False после проверки
DeepFace = None


def _try_import_deepface():
    """Ленивая загрузка DeepFace. Возвращает True, если доступен."""
    global DEEPFACE_AVAILABLE, DeepFace
    if DEEPFACE_AVAILABLE is not None:
        return DEEPFACE_AVAILABLE
    try:
        from deepface import DeepFace as _DF
        DeepFace = _DF
        DEEPFACE_AVAILABLE = True
    except Exception as e:
        DEEPFACE_AVAILABLE = False
        print(f"DeepFace недоступен ({e}). Распознавание эмоций отключено. "
              f"Установите: pip install deepface tf-keras")
    return DEEPFACE_AVAILABLE


# ================= НАСТРОЙКИ =================
MODEL_PATH = "vosk-model-small-ru-0.22"
SAMPLE_RATE = 16000
DATA_FILE = "face_data_lbph.pkl"
LBPH_MODEL_FILE = "lbph_model.yml"

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

# ================= GPT4All (локальная LLM для умных ответов) =================
GPT4ALL_API_URL = "http://localhost:4891/v1/chat/completions"
# "auto" — взять первую загруженную в приложении модель. Можно прописать
# точное имя модели из GPT4All, если "auto" не сработает.
GPT4ALL_MODEL = "auto"
GPT4ALL_TIMEOUT = 90  # секунд на ответ локальной модели (8B на CPU медленный)
GPT4ALL_SYSTEM_PROMPT = (
    "Ты — дружелюбный голосовой ассистент по имени Друг. Отвечай кратко, "
    "по-русски, в одном-двух предложениях, тёплым и простым языком."
)

# ================= ЭМОЦИИ (DeepFace) =================
# DeepFace возвращает эмоции на английском. Переводим на русский, эмодзи
# (для GUI-лога и текста) и ASCII-смайл (для окна камеры — цветные эмодзи
# в OpenCV/PIL рисуются ненадёжно, часто выходит пустой квадрат).
EMOTION_RU = {
    "angry":    ("злость",      "😠", ">:("),
    "disgust":  ("отвращение",  "🤢", ":S"),
    "fear":     ("страх",       "😨", ":-O"),
    "happy":    ("радость",     "😊", ":)"),
    "sad":      ("грусть",      "😢", ":("),
    "surprise": ("удивление",   "😲", ":O"),
    "neutral":  ("спокойствие", "😐", ":|"),
}
# Как часто пересчитывать эмоцию (DeepFace тяжёлый — нельзя каждый кадр).
EMOTION_RECALC_INTERVAL = 1.5  # секунд

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
        # Управление прерыванием речи (для команды "стоп").
        self._current_sound = None      # текущий проигрываемый звук (playsound3)
        self._stop_speaking = False     # флаг: прервать текущую речь
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
        self.log_to_gui(f"Шрифт для окна камеры: {getattr(self, '_font_source', 'неизвестен')}")
        self.last_spoken_name = ""
        self.last_spoken_time = 0
        self.speak_cooldown = 3.0
        self.recognizer = cv2.face.LBPHFaceRecognizer_create(radius=2, neighbors=8,
                                                               grid_x=8, grid_y=8)
        self.model_trained = False
        # Последний результат распознавания на лицо — чтобы рисовать
        # рамку/подпись/вероятность на КАЖДОМ кадре, а не только в момент
        # пересчёта (иначе текст мигает/пропадает между обновлениями).
        # Формат: {face_key: {"name", "confidence_pct", "distance", "color",
        #                      "box", "last_update"}}
        self.last_results = {}
        # Последняя распознанная эмоция (для отображения в окне и озвучки).
        # Формат: {"ru": str, "emoji": str, "ascii": str, "score": int}
        self.last_emotion = None
        self.emotion_lock = threading.Lock()
        self._emotion_busy = False  # идёт ли сейчас анализ эмоции
        # Доступность GPT4All проверяем при первом обращении.
        self.gpt4all_available = None
        self.gpt4all_model_name = None
        self.load_face_data()

        # 4. Голосовой ввод
        self.vosk_recognizer = None
        self.init_vosk()
        self.start_audio_stream()

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
        """Озвучивает текст: пытается edge-tts, при ошибке падает на SAPI.
        Воспроизведение прерываемое — команда "стоп" обрывает его."""
        if self._stop_speaking:
            return
        if EDGE_TTS_AVAILABLE and PLAYSOUND_AVAILABLE:
            try:
                voice_id = EDGE_VOICES.get(self.current_voice_name, "ru-RU-DmitryNeural")
                tmp_path = os.path.join(tempfile.gettempdir(),
                                         f"assistant_tts_{int(time.time()*1000)}.mp3")
                asyncio.run(self._edge_tts_save(text, voice_id, tmp_path))
                if self._stop_speaking:
                    return
                # Неблокирующее воспроизведение, чтобы можно было прервать.
                sound = playsound(tmp_path, block=False)
                self._current_sound = sound
                # Ждём окончания, периодически проверяя флаг остановки.
                while sound.is_alive():
                    if self._stop_speaking:
                        try:
                            sound.stop()
                        except Exception:
                            pass
                        break
                    time.sleep(0.05)
                self._current_sound = None
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
                return
            except Exception as e:
                self.log_to_gui(f"edge-tts недоступен ({e}), переключаюсь на SAPI.")

        if self.sapi_voice and not self._stop_speaking:
            try:
                self.sapi_voice.Speak(text, 1)
            except Exception:
                pass

    def shut_up(self):
        """Немедленно обрывает текущую речь и очищает очередь озвучки."""
        self._stop_speaking = True
        # Останавливаем то, что играет прямо сейчас.
        if self._current_sound is not None:
            try:
                self._current_sound.stop()
            except Exception:
                pass
        # Очищаем очередь невыговоренных фраз.
        try:
            while True:
                self.speak_queue.get_nowait()
                self.speak_queue.task_done()
        except queue.Empty:
            pass
        # Сбрасываем флаг чуть позже, чтобы успели «проскочить» уже
        # запущенные фоновые приветствия, и снова можно было говорить.
        def _reset():
            time.sleep(0.5)
            self._stop_speaking = False
        threading.Thread(target=_reset, daemon=True).start()

    @staticmethod
    async def _edge_tts_save(text, voice_id, path):
        communicate = edge_tts.Communicate(text, voice=voice_id)
        await communicate.save(path)

    def speak(self, text: str):
        if self._stop_speaking:
            return
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
            ("Лица и эмоции", [
                ("Камера / распознавание", lambda: self.process_command("распознавание")),
                ("Добавить лицо", lambda: self.process_command("добавь человека")),
                ("Моя эмоция", lambda: self.process_command("какая у меня эмоция")),
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
                ("Спросить ИИ", lambda: self.ask_ai_by_dialog()),
            ]),
            ("Общение", [
                ("Как дела?", lambda: self.process_command("как дела")),
                ("Как настроение?", lambda: self.process_command("как настроение")),
                ("Справка", self.show_help),
            ]),
        ]
        for title, buttons in sections:
            box = tk.LabelFrame(parent, text=title, bg="#f0f0f0")
            box.pack(fill=tk.X, padx=8, pady=6)
            for label, cmd in buttons:
                ttk.Button(box, text=label, command=cmd).pack(fill=tk.X, pady=2, padx=4)

        # Отдельная заметная кнопка выхода внизу панели.
        exit_box = tk.Frame(parent, bg="#f0f0f0")
        exit_box.pack(fill=tk.X, side=tk.BOTTOM, padx=8, pady=10)
        exit_btn = tk.Button(exit_box, text="🚪 ВЫХОД", command=self.exit_app,
                             bg="#e74c3c", fg="white", font=("Arial", 11, "bold"),
                             activebackground="#c0392b", activeforeground="white",
                             relief=tk.RAISED, bd=2)
        exit_btn.pack(fill=tk.X, ipady=6)

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
            # "q" в текстовом поле — быстрый выход.
            if text.lower() == "q":
                self.exit_app()
                return
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
            "- 'Какая у меня эмоция' — распознать эмоцию\n"
            "- 'Который час?'\n"
            "- 'Погода [в городе]'\n"
            "- 'Открой браузер' / 'открой ютуб'\n"
            "- 'Расскажи шутку'\n"
            "- 'Интересный факт'\n"
            "- 'Цитата дня'\n"
            "- 'Сколько будет 5 плюс 3'\n"
            "- 'Как дела?', 'Как настроение?' — простые вопросы\n"
            "- Любой другой вопрос — ответит ИИ (GPT4All)\n"
            "- 'Выход' / 'q' / 'стоп' — закрыть программу")

    def ask_ai_by_dialog(self):
        question = simpledialog_askstring("Спросить ИИ", "Введите вопрос для ИИ:")
        if question and question.strip():
            self.log_to_gui(f"Вы (ИИ): {question}")
            self.ask_gpt4all(question.strip())
        else:
            self.speak("Отменено.")

    def exit_app(self):
        """Корректный выход из программы."""
        self.speak("До свидания!")
        self.is_recognizing_faces = False
        # Небольшая задержка, чтобы успело проговориться прощание.
        def _do_exit():
            self.cleanup()
            try:
                self.root.quit()
                self.root.destroy()
            except Exception:
                pass
            os._exit(0)
        threading.Thread(target=lambda: (time.sleep(1.5), _do_exit()), daemon=True).start()

    def delete_person_by_dialog(self):
        name = simpledialog_askstring("Удаление", "Введите имя человека для удаления:")
        if name:
            self.delete_person(name.strip())
        else:
            self.speak("Отменено.")

    def load_cyrillic_font(self, size=28):
        """Ищет шрифт с поддержкой кириллицы. Перебирает стандартные
        шрифты Windows по нескольким путям. Если не найдёт — вернёт
        встроенный шрифт PIL (он тоже умеет кириллицу в новых версиях)."""
        # Каталог шрифтов Windows (обычно C:\Windows\Fonts).
        win_fonts = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
        candidates = [
            os.path.join(win_fonts, "arial.ttf"),
            os.path.join(win_fonts, "calibri.ttf"),
            os.path.join(win_fonts, "tahoma.ttf"),
            os.path.join(win_fonts, "segoeui.ttf"),
            os.path.join(win_fonts, "verdana.ttf"),
            "C:/Windows/Fonts/arial.ttf",
            "arial.ttf",  # иногда PIL находит по имени в системных путях
        ]
        for path in candidates:
            try:
                if os.path.exists(path):
                    font = ImageFont.truetype(path, size)
                    self._font_source = path
                    return font
            except Exception:
                continue
        # Последняя попытка — PIL может найти arial по имени без полного пути.
        try:
            font = ImageFont.truetype("arial.ttf", size)
            self._font_source = "arial.ttf (по имени)"
            return font
        except Exception:
            pass
        # Совсем запасной вариант — встроенный шрифт PIL (мелкий, но рабочий).
        self._font_source = "встроенный PIL (load_default)"
        try:
            return ImageFont.load_default(size)
        except Exception:
            return ImageFont.load_default()

    def draw_text_safe(self, img, text, position, color=(0, 255, 255)):
        if self.font is not None:
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
                # Обрываем текущую речь и все приветствия в очереди.
                self.shut_up()
                # Подтверждаем выход чуть позже, когда флаг обрыва сбросится.
                def _confirm():
                    time.sleep(0.7)
                    self.speak("Выхожу из режима камеры.")
                threading.Thread(target=_confirm, daemon=True).start()
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

        if 'эмоци' in command or 'настроение у меня' in command or \
           'у меня настроение' in command or 'моё настроение' in command or \
           'мое настроение' in command or 'какое у меня лицо' in command or \
           'что у меня на лице' in command:
            self.handle_emotion_query()
            return

        if any(w in command for w in ['время', 'час']):
            self.speak(f"Сейчас {datetime.now().strftime('%H:%M')}")
            return

        if any(w in command for w in ['стоп', 'выход', 'пока', 'закройся', 'выключись']):
            self.exit_app()
            return

        empathy_reply = self._match_empathy(command)
        if empathy_reply:
            self.speak(empathy_reply)
            return

        # Если ничего не подошло — отправляем вопрос локальной модели
        # GPT4All (если её API-сервер запущен). Иначе сообщаем, что не поняли.
        self.ask_gpt4all(command)

    @staticmethod
    def _match_empathy(command: str):
        """Ищет совпадение с простыми личными вопросами (как дела, как
        настроение и т.п.) и возвращает случайный тёплый ответ, либо None."""
        for triggers, replies in EMPATHY_RESPONSES:
            if any(trigger in command for trigger in triggers):
                return random.choice(replies)
        return None

    # ================= GPT4All (умные ответы) =================
    def ask_gpt4all(self, question: str):
        """Отправляет вопрос локальной модели GPT4All и озвучивает ответ.
        Работает в отдельном потоке, чтобы не блокировать интерфейс."""
        def worker():
            if self.gpt4all_available is None:
                self._check_gpt4all()

            if not self.gpt4all_available:
                self.speak("Я не понял команду. Чтобы я мог отвечать на любые "
                           "вопросы, включите API-сервер в приложении GPT4All.")
                return

            self.log_to_gui("Думаю над ответом (GPT4All)...")
            answer = self._gpt4all_complete(question)
            if answer:
                self.speak(answer)
            else:
                self.speak("Не получилось получить ответ от ИИ.")

        threading.Thread(target=worker, daemon=True).start()

    def _gpt4all_complete(self, prompt: str, max_tokens: int = 200):
        """Синхронный запрос к GPT4All. Возвращает текст ответа или None.
        НЕ озвучивает сам — это делает вызывающая сторона. Используется
        и для ответов на вопросы, и для приветствий в режиме камеры.
        max_tokens поменьше = быстрее ответ (важно для приветствий)."""
        if self.gpt4all_available is None:
            self._check_gpt4all()
        if not self.gpt4all_available:
            return None
        try:
            model_name = self.gpt4all_model_name or "auto"
            payload = {
                "model": model_name,
                "messages": [
                    {"role": "system", "content": GPT4ALL_SYSTEM_PROMPT},
                    {"role": "user", "content": prompt},
                ],
                "max_tokens": max_tokens,
                "temperature": 0.7,
            }
            resp = requests.post(GPT4ALL_API_URL, json=payload, timeout=GPT4ALL_TIMEOUT)
            if resp.status_code == 200:
                data = resp.json()
                answer = data["choices"][0]["message"]["content"].strip()
                return answer or None
            else:
                self.log_to_gui(f"GPT4All HTTP {resp.status_code}: {resp.text[:200]}")
                return None
        except requests.exceptions.ConnectionError:
            self.gpt4all_available = False
            self.log_to_gui("Потеряна связь с GPT4All.")
            return None
        except requests.exceptions.Timeout:
            self.log_to_gui("GPT4All слишком долго думает.")
            return None
        except Exception as e:
            self.log_to_gui(f"Ошибка GPT4All: {e}")
            return None

    def _check_gpt4all(self):
        """Проверяет доступность GPT4All API и определяет имя модели."""
        try:
            resp = requests.get("http://localhost:4891/v1/models", timeout=3)
            if resp.status_code == 200:
                data = resp.json()
                models = data.get("data", [])
                if models:
                    # Берём первую загруженную модель, если GPT4ALL_MODEL="auto"
                    if GPT4ALL_MODEL == "auto":
                        self.gpt4all_model_name = models[0].get("id", "auto")
                    else:
                        self.gpt4all_model_name = GPT4ALL_MODEL
                    self.gpt4all_available = True
                    self.log_to_gui(f"GPT4All подключён, модель: {self.gpt4all_model_name}")
                    return
            self.gpt4all_available = False
        except Exception:
            self.gpt4all_available = False
            self.log_to_gui("GPT4All API недоступен (это нормально, если вы его не включили).")

    # ================= ЭМОЦИИ (DeepFace) =================
    def handle_emotion_query(self):
        """Озвучивает текущую эмоцию. Если режим камеры активен — берёт
        последнюю распознанную эмоцию; иначе делает разовый снимок."""
        if not _try_import_deepface():
            self.speak("Распознавание эмоций недоступно. Нужно установить "
                       "библиотеку DeepFace.")
            return

        with self.emotion_lock:
            emotion = self.last_emotion

        if self.is_recognizing_faces and emotion:
            self._speak_emotion(emotion)
        else:
            # Камера не активна — делаем разовый снимок и анализируем.
            self.speak("Секунду, посмотрю на вас.")
            threading.Thread(target=self._analyze_emotion_snapshot, daemon=True).start()

    def _analyze_emotion_snapshot(self):
        """Делает один снимок с камеры и определяет эмоцию."""
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.speak("Не удалось открыть камеру.")
            return
        try:
            # Прогреваем камеру несколькими кадрами.
            frame = None
            for _ in range(10):
                ret, frame = cap.read()
                if not ret:
                    frame = None
                time.sleep(0.05)
        finally:
            cap.release()

        if frame is None:
            self.speak("Не удалось получить изображение с камеры.")
            return

        emotion = self._detect_emotion(frame)
        if emotion:
            with self.emotion_lock:
                self.last_emotion = emotion
            self._speak_emotion(emotion)
        else:
            self.speak("Не получилось распознать эмоцию. Попробуйте смотреть прямо в камеру.")

    def _detect_emotion(self, frame_bgr):
        """Запускает DeepFace на кадре, возвращает dict эмоции или None.
        Применяет сглаживание: усредняет вероятности по нескольким
        последним измерениям, чтобы название эмоции не скакало."""
        if not _try_import_deepface():
            return None
        try:
            result = DeepFace.analyze(frame_bgr, actions=['emotion'],
                                       detector_backend='opencv',
                                       enforce_detection=False, silent=True)
            if isinstance(result, list):
                if not result:
                    return None
                result = result[0]
            scores = result.get('emotion', {})
            if not scores:
                return None

            # Сглаживание: храним последние N распределений вероятностей и
            # усредняем их. Это убирает резкие скачки названия эмоции.
            if not hasattr(self, '_emotion_history'):
                self._emotion_history = []
            self._emotion_history.append(scores)
            if len(self._emotion_history) > 4:
                self._emotion_history.pop(0)

            # Усредняем вероятности по всем эмоциям.
            avg = {}
            for d in self._emotion_history:
                for k, v in d.items():
                    avg[k] = avg.get(k, 0.0) + float(v)
            n = len(self._emotion_history)
            for k in avg:
                avg[k] /= n

            dominant = max(avg, key=avg.get)
            ru, emoji, ascii_face = EMOTION_RU.get(
                dominant, (dominant, "", ":|"))
            score = int(avg.get(dominant, 0))
            return {"en": dominant, "ru": ru, "emoji": emoji,
                    "ascii": ascii_face, "score": score}
        except Exception as e:
            self.log_to_gui(f"Ошибка DeepFace: {e}")
            return None

    def _speak_emotion(self, emotion):
        """Озвучивает эмоцию в виде вопроса (как просил пользователь)."""
        ru = emotion["ru"]
        emoji = emotion["emoji"]
        # Формулируем как вопрос-предположение об эмоции.
        phrases = {
            "радость":     "Вы выглядите радостным! У вас хорошее настроение?",
            "грусть":      "Кажется, вам немного грустно. Всё в порядке?",
            "злость":      "Вы выглядите рассерженным. Что-то случилось?",
            "удивление":   "Вы чем-то удивлены?",
            "страх":       "Вы выглядите встревоженным. Всё хорошо?",
            "отвращение":  "Что-то вам не по душе?",
            "спокойствие": "Вы выглядите спокойным и сосредоточенным.",
        }
        phrase = phrases.get(ru, f"Похоже, ваша эмоция — {ru}.")
        self.log_to_gui(f"Эмоция: {ru} {emoji} ({emotion['score']}%)")
        self.speak(phrase)

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
                cv2.imshow("Dobavlenie lica (q - otmena)", frame)
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

        # last_result хранит последний посчитанный результат распознавания
        # для постоянной отрисовки. Сбрасывается, когда лицо исчезает.
        self.last_results = {"main": None}
        emotion_enabled = _try_import_deepface()
        last_emotion_time = 0
        with self.emotion_lock:
            self.last_emotion = None
        self._emotion_history = []  # сброс сглаживания эмоций

        # Сглаживание распознавания: имя считается "подтверждённым" только
        # если совпадает в нескольких последних кадрах подряд. Это убирает
        # скачки между "Рома" и "Неизвестен".
        recent_names = []          # имена за последние кадры
        SMOOTH_WINDOW = 5          # сколько кадров держим для голосования
        stable_name = None         # текущее подтверждённое имя
        stable_is_match = False

        # Озвучка "человек + эмоция":
        greeted_name = None        # кого последний раз приветствовали
        last_greet_time = 0
        GREET_INTERVAL = 15        # секунд между повторными приветствиями
        frames_without_face = 0    # счётчик кадров без лица (для сброса)

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
                    frames_without_face = 0
                    # Берём самое крупное лицо в кадре.
                    x, y, w, h = max(faces, key=lambda f: f[2] * f[3])

                    feat = self.get_face_features(frame[y:y + h, x:x + w])
                    with self.faces_lock:
                        try:
                            label, distance = self.recognizer.predict(feat)
                            name = self.known_names.get(label)
                        except cv2.error:
                            label, distance, name = None, 999.0, None

                    confidence_pct = int(max(0, min(100, 100 - distance)))
                    raw_is_match = (name is not None) and (distance < LBPH_DISTANCE_THRESHOLD)
                    raw_name = name if raw_is_match else UNKNOWN_LABEL

                    # Голосование по последним кадрам — берём самое частое имя.
                    recent_names.append(raw_name)
                    if len(recent_names) > SMOOTH_WINDOW:
                        recent_names.pop(0)
                    # Подтверждаем имя, только если оно встречается в большинстве
                    # последних кадров (>= половины окна).
                    most_common = max(set(recent_names), key=recent_names.count)
                    if recent_names.count(most_common) >= max(2, SMOOTH_WINDOW // 2):
                        stable_name = most_common
                        stable_is_match = (stable_name != UNKNOWN_LABEL)

                    display_name = stable_name if stable_name else raw_name
                    is_match = stable_is_match

                    self.last_results["main"] = {
                        "name": display_name,
                        "is_match": is_match,
                        "confidence_pct": confidence_pct,
                        "distance": distance,
                        "box": (x, y, w, h),
                    }

                    # Анализ эмоции в фоне с интервалом. Передаём лицо с
                    # небольшим запасом по краям — так DeepFace точнее.
                    if emotion_enabled and not self._emotion_busy and \
                       (now - last_emotion_time) > EMOTION_RECALC_INTERVAL:
                        last_emotion_time = now
                        pad = int(0.2 * h)  # запас 20% вокруг лица
                        y1 = max(0, y - pad)
                        y2 = min(frame.shape[0], y + h + pad)
                        x1 = max(0, x - pad)
                        x2 = min(frame.shape[1], x + w + pad)
                        face_crop = frame[y1:y2, x1:x2].copy()
                        self._start_emotion_worker(face_crop)

                    # Озвучка приветствия — только для ПОДТВЕРЖДЁННОГО имени,
                    # и только когда: сменился человек ИЛИ прошёл интервал.
                    if stable_name is not None:
                        changed_person = (stable_name != greeted_name)
                        time_passed = (now - last_greet_time) > GREET_INTERVAL
                        if changed_person or time_passed:
                            greeted_name = stable_name
                            last_greet_time = now
                            with self.emotion_lock:
                                emotion_snapshot = self.last_emotion
                            self._announce_person_emotion(
                                stable_name, is_match, emotion_snapshot)
                else:
                    # Лица нет в кадре. Через несколько пустых кадров сбрасываем
                    # состояние — чтобы при новом появлении снова поздороваться,
                    # а старые надписи/эмоция не висели на экране.
                    frames_without_face += 1
                    if frames_without_face > 10:
                        self.last_results["main"] = None
                        recent_names.clear()
                        stable_name = None
                        greeted_name = None
                        with self.emotion_lock:
                            self.last_emotion = None

                # Отрисовываем последний результат (если лицо есть/недавно было).
                result = self.last_results.get("main")
                if result:
                    x, y, w, h = result["box"]
                    color = (0, 255, 0) if result["is_match"] else (0, 0, 255)
                    thickness = 3 if result["is_match"] else 2
                    cv2.rectangle(frame, (x, y), (x + w, y + h), color, thickness)
                    frame = self.draw_text_safe(frame, result["name"], (x, max(0, y - 40)), color)

                    with self.emotion_lock:
                        emotion = self.last_emotion
                    if emotion:
                        emo_text = f"{emotion['ru']} {emotion['ascii']} ({emotion['score']}%)"
                        frame = self.draw_text_safe(
                            frame, emo_text, (x, y + h + 5), (255, 200, 0))

                    frame = self.draw_text_safe(
                        frame, f"Вероятность: {result['confidence_pct']}%",
                        (15, 15), (0, 255, 255))
                    frame = self.draw_text_safe(
                        frame, f"Расстояние: {result['distance']:.1f}",
                        (15, 50), (0, 255, 255))

                # Описание алгоритма — нижний правый угол.
                h_frame, w_frame, _ = frame.shape
                start_y = h_frame - 20 * (len(ALGORITHM_DESCRIPTION) + 1)
                for i, line in enumerate(ALGORITHM_DESCRIPTION):
                    cv2.putText(frame, line, (w_frame - 480, start_y + i * 20),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255, 255, 255), 1)

                cv2.imshow("Raspoznavanie (q - vyhod)", frame)
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
        finally:
            cap.release()
            cv2.destroyAllWindows()
            self.is_recognizing_faces = False
            self.last_results = {}

    def _announce_person_emotion(self, display_name, is_match, emotion):
        """Озвучивает человека и его эмоцию. Фразу формулирует GPT4All
        (если включён), иначе — простой шаблон. Запускается в отдельном
        потоке, чтобы не тормозить видео."""
        emo_ru = emotion["ru"] if emotion else None

        def worker():
            # Готовим запрос для GPT4All.
            if is_match:
                who = f"знакомый человек по имени {display_name}"
            else:
                who = "незнакомый человек"

            if emo_ru:
                task = (f"Перед камерой {who}. Его эмоция сейчас: {emo_ru}. "
                        f"Поприветствуй его одной короткой дружелюбной фразой "
                        f"по-русски и мягко отреагируй на эмоцию. "
                        f"Если человек незнакомый — прояви тёплое участие. "
                        f"Только сама фраза, без пояснений.")
            else:
                task = (f"Перед камерой {who}. Поприветствуй его одной "
                        f"короткой дружелюбной фразой по-русски. "
                        f"Только сама фраза, без пояснений.")

            # Проверяем GPT4All один раз.
            if self.gpt4all_available is None:
                self._check_gpt4all()

            spoken = False
            if self.gpt4all_available:
                answer = self._gpt4all_complete(task, max_tokens=60)
                # Пока GPT4All думал, режим камеры мог уже выключиться
                # (нажали "стоп") — тогда не озвучиваем.
                if answer and self.is_recognizing_faces:
                    self.speak(answer)
                    spoken = True

            # Запасной вариант — шаблонная фраза (если GPT4All выключен/молчит).
            if not spoken and self.is_recognizing_faces:
                self.speak(self._template_greeting(display_name, is_match, emo_ru))

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _template_greeting(display_name, is_match, emo_ru):
        """Простая шаблонная фраза, когда GPT4All недоступен."""
        emotion_part = ""
        if emo_ru:
            reactions = {
                "радость":     "выглядишь радостным, приятно видеть!",
                "грусть":      "кажется, тебе немного грустно. Всё наладится.",
                "злость":      "выглядишь напряжённым. Надеюсь, всё в порядке.",
                "удивление":   "ты чем-то удивлён?",
                "страх":       "не волнуйся, здесь всё спокойно.",
                "отвращение":  "что-то не по душе?",
                "спокойствие": "выглядишь спокойным и собранным.",
            }
            emotion_part = " " + reactions.get(emo_ru, f"твоя эмоция — {emo_ru}.")

        if is_match:
            return f"Привет, {display_name}!{emotion_part}"
        else:
            return f"Здравствуй, незнакомец!{emotion_part}"

    def _start_emotion_worker(self, face_crop):
        """Запускает анализ эмоции в отдельном потоке (DeepFace тяжёлый)."""
        self._emotion_busy = True

        def worker():
            try:
                emotion = self._detect_emotion(face_crop)
                if emotion:
                    with self.emotion_lock:
                        self.last_emotion = emotion
            finally:
                self._emotion_busy = False

        threading.Thread(target=worker, daemon=True).start()

    def show_face_list(self):
        if self.known_names:
            self.speak("В базе: " + ", ".join(self.known_names.values()))
        else:
            self.speak("База пуста.")

    def cleanup(self):
        self.is_listening = False
        self.speak_queue.put(None)


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

        def on_close():
            assistant.cleanup()
            try:
                assistant.root.quit()
                assistant.root.destroy()
            except Exception:
                pass
            os._exit(0)

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