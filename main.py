import os
import cv2
import pickle
import numpy as np
import threading
import time
import queue
import json
import random
import webbrowser
import sys
from datetime import datetime
import sounddevice as sd
import win32com.client
from PIL import Image, ImageDraw, ImageFont

# --- GUI ---
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

# --- Vosk ---
try:
    from vosk import Model, KaldiRecognizer

    VOSK_AVAILABLE = True
except ImportError:
    VOSK_AVAILABLE = False
    print("❌ Vosk не установлен. Установите: pip install vosk")

# ================= НАСТРОЙКИ =================
MODEL_PATH = "vosk-model-small-ru-0.22"
SAMPLE_RATE = 16000
DATA_FILE = "face_data_v2.pkl"
UNKNOWN_NAME = "я тебя не знаю"

USER_PROFILES = {
    "папа": {"voice_index": 0, "greeting": "Привет, папа! Я слушаю."},
    "мама": {"voice_index": 0, "greeting": "Здравствуй, мама!"},
    "сын": {"voice_index": 0, "greeting": "Привет, сынок!"},
    "ирина": {"voice_index": 0, "greeting": "Ирина, я готова."}
}


class SmartAssistant:
    def __init__(self):
        self.current_profile_name = list(USER_PROFILES.keys())[0]
        self.current_profile = USER_PROFILES[self.current_profile_name]
        self.is_recognizing_faces = False
        self.is_listening = True

        self.audio_queue = queue.Queue()
        self.speak_queue = queue.Queue()
        self.listen_thread = None

        # 1. ИНИЦИАЛИЗАЦИЯ GUI (ПЕРВЫМ ДЕЛОМ)
        self.root = None
        self.log_area = None
        self.setup_gui()

        # 2. Инициализация голоса
        self.sapi_voice, self.voices_list, self.voice_descriptions = self.init_speaker()
        self.speaker_thread = threading.Thread(target=self._speaker_loop, daemon=True)
        self.speaker_thread.start()

        # 3. Распознавание лиц
        self.known_faces = []
        self.face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
        self.font = self.load_cyrillic_font()
        self.last_spoken_name = ""
        self.last_spoken_time = 0
        self.speak_cooldown = 3.0
        self.similarity_threshold = 0.55
        self.load_face_data()

        # 4. Голосовой ввод
        self.vosk_recognizer = None
        self.init_vosk()

        # 5. Запуск аудио-потока
        self.start_audio_stream()

    def init_speaker(self):
        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            voices = speaker.GetVoices()
            voices_list = [voices.Item(i) for i in range(voices.Count)]
            voice_descriptions = {i: v.GetDescription() for i, v in enumerate(voices_list)}

            idx = self.current_profile["voice_index"]
            if 0 <= idx < len(voices_list):
                speaker.Voice = voices_list[idx]

            speaker.Rate = 0
            speaker.Volume = 100
            return speaker, voices_list, voice_descriptions
        except Exception as e:
            print(f"❌ Ошибка инициализации SAPI: {e}")
            sys.exit(1)

    def _speaker_loop(self):
        while self.is_listening:
            text = self.speak_queue.get()
            if text is None:
                break
            try:
                self.sapi_voice.Speak(text, 1)  # 1 = асинхронно
            except Exception as e:
                print(f"Ошибка озвучивания: {e}")
            self.speak_queue.task_done()

    def speak(self, text: str):
        print(f"🤖 Ассистент: {text}")
        self.log_to_gui(f"🤖 Ассистент: {text}")
        self.speak_queue.put(text)

    def log_to_gui(self, message: str):
        print(message)
        if hasattr(self, 'root') and self.root is not None and hasattr(self, 'log_area') and self.log_area is not None:
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
        self.root.title("🤖 Умный Голосовой Помощник + Распознавание Лиц")
        self.root.geometry("850x650")
        self.root.configure(bg="#f0f0f0")

        top_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        top_frame.pack(fill=tk.X)

        self.profile_var = tk.StringVar(value=self.current_profile_name)
        profile_combo = ttk.Combobox(top_frame, textvariable=self.profile_var, values=list(USER_PROFILES.keys()),
                                     state="readonly", width=15)
        profile_combo.pack(side=tk.LEFT, padx=15, pady=15)
        profile_combo.bind("<<ComboboxSelected>>", lambda e: self.set_profile(self.profile_var.get()))

        self.status_label = tk.Label(top_frame, text="🎤 Микрофон: Активен | Говорите 'друг' или команду", fg="#2ecc71",
                                     bg="#2c3e50", font=("Arial", 10, "bold"))
        self.status_label.pack(side=tk.RIGHT, padx=15, pady=15)

        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        log_frame = tk.LabelFrame(main_frame, text="📜 Журнал действий и диалог", bg="#f0f0f0",
                                  font=("Arial", 10, "bold"))
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', font=("Consolas", 10), bg="#ffffff",
                                                  fg="#333333")
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        control_frame = tk.LabelFrame(main_frame, text="⚡ Управление и Возможности", bg="#f0f0f0",
                                      font=("Arial", 10, "bold"), width=250)
        control_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        control_frame.pack_propagate(False)

        help_text = (
            "🗣️ Голосовые команды:\n"
            "• 'Друг, справка'\n"
            "• 'Я папа' (смена профиля)\n"
            "• 'Который час?', 'Погода'\n"
            "• 'Запусти распознавание'\n"
            "• 'Добавь человека'\n"
            "• 'Стоп' / 'Выход'\n\n"
            "⌨️ Введите команду в поле ниже и нажмите Enter."
        )
        help_label = tk.Label(control_frame, text=help_text, justify=tk.LEFT, bg="#f0f0f0", font=("Arial", 9),
                              anchor="w")
        help_label.pack(fill=tk.X, padx=10, pady=10)

        btn_frame = tk.Frame(control_frame, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(btn_frame, text="📷 Камера", command=lambda: self.process_command("распознавание")).pack(fill=tk.X,
                                                                                                           pady=2)
        ttk.Button(btn_frame, text="➕ Добавить лицо", command=lambda: self.process_command("добавь человека")).pack(
            fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="👥 Кто в базе?", command=lambda: self.process_command("список лиц")).pack(fill=tk.X,
                                                                                                             pady=2)
        ttk.Button(btn_frame, text="📋 Полная справка", command=self.show_help).pack(fill=tk.X, pady=10)

        input_frame = tk.Frame(self.root, bg="#f0f0f0")
        input_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.command_entry = ttk.Entry(input_frame, font=("Arial", 12))
        self.command_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.command_entry.bind("<Return>", lambda event: self.send_command_from_gui())

        send_btn = ttk.Button(input_frame, text="Отправить ➤", command=self.send_command_from_gui)
        send_btn.pack(side=tk.RIGHT)

        self.log_to_gui(f"✅ Система запущена. Профиль: {self.current_profile_name.capitalize()}")
        self.log_to_gui("💡 Скажите 'Друг, справка' или нажмите кнопку 'Полная справка'.")

    def send_command_from_gui(self):
        text = self.command_entry.get().strip()
        if text:
            self.log_to_gui(f"👤 Вы (клавиатура): {text}")
            self.command_entry.delete(0, tk.END)
            self.process_command(text)

    def show_help(self):
        help_msg = (
            "📋 ПОЛНЫЙ СПИСОК ВОЗМОЖНОСТЕЙ:\n\n"
            "1. УПРАВЛЕНИЕ ПРОФИЛЯМИ:\n"
            "   - 'Смени профиль на [имя]'\n"
            "   - 'Я [имя]' (например, 'Я папа')\n\n"
            "2. ИНФОРМАЦИЯ:\n"
            "   - 'Который час?', 'Какое сегодня число?'\n"
            "   - 'Как дела?'\n\n"
            "3. ИНТЕРНЕТ:\n"
            "   - 'Открой браузер'\n"
            "   - 'Найди в Яндексе [запрос]'\n"
            "   - 'Покажи погоду'\n\n"
            "4. РАЗВЛЕЧЕНИЯ:\n"
            "   - 'Расскажи шутку' / 'Анекдот'\n\n"
            "5. РАСПОЗНАВАНИЕ ЛИЦ:\n"
            "   - 'Запусти распознавание' (откроет камеру)\n"
            "   - 'Добавь человека' (сделает снимки для базы)\n"
            "   - 'Кто в базе?' (покажет список)\n\n"
            "6. ЗАВЕРШЕНИЕ:\n"
            "   - 'Стоп', 'Выход', 'Пока'"
        )
        messagebox.showinfo("Справка по возможностям", help_msg)

    def set_profile(self, profile_name: str):
        prof_name = profile_name.lower().strip()
        matched = None
        for key in USER_PROFILES:
            if key in prof_name or prof_name in key:
                matched = key
                break

        if not matched:
            self.speak(f"Профиль '{prof_name}' не найден.")
            return False

        prof = USER_PROFILES[matched]
        idx = prof["voice_index"]

        if 0 <= idx < len(self.voices_list):
            self.sapi_voice.Voice = self.voices_list[idx]
            self.current_profile = prof
            self.current_profile_name = matched
            self.profile_var.set(matched)
            self.speak(prof.get("greeting", f"Профиль {matched} активирован."))
            return True
        return False

    def load_cyrillic_font(self):
        font_paths = [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibri.ttf",
            "C:/Windows/Fonts/times.ttf"
        ]
        for path in font_paths:
            if os.path.exists(path):
                try:
                    font = ImageFont.truetype(path, 32)
                    return font
                except Exception as e:
                    print(f"Не удалось загрузить шрифт {path}: {e}")

        print("⚠️ Кириллический шрифт не найден. Имена в окне камеры могут отображаться как '???'.")
        return None

    def draw_text_with_pil(self, img, text, position, color=(0, 255, 255)):
        if self.font is None:
            cv2.putText(img, "Face Detected", position, cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)
            return img

        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(img_rgb)
        draw = ImageDraw.Draw(pil_img)
        draw.text(position, text, font=self.font, fill=color[::-1])
        return cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)

    def init_vosk(self):
        if not VOSK_AVAILABLE:
            return
        if os.path.exists(MODEL_PATH):
            try:
                model = Model(MODEL_PATH)
                self.vosk_recognizer = KaldiRecognizer(model, SAMPLE_RATE)
                self.vosk_recognizer.SetWords(True)
                self.log_to_gui("🎤 Модель Vosk успешно загружена.")
            except Exception as e:
                self.log_to_gui(f"❌ Ошибка загрузки модели Vosk: {e}")
        else:
            self.log_to_gui(f"❌ Модель Vosk не найдена в папке '{MODEL_PATH}'.")

    def start_audio_stream(self):
        if not self.vosk_recognizer:
            return
        try:
            self.listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
            self.listen_thread.start()
        except Exception as e:
            self.log_to_gui(f"❌ Ошибка запуска микрофона: {e}")

    def _listen_loop(self):
        def audio_callback(indata, frames, time_info, status):
            if status:
                print(status)
            int_data = (indata * 32767).astype(np.int16)
            self.audio_queue.put(int_data.tobytes())

        try:
            with sd.InputStream(samplerate=SAMPLE_RATE, blocksize=8000, channels=1, callback=audio_callback):
                while self.is_listening:
                    data = self.audio_queue.get()
                    if self.vosk_recognizer.AcceptWaveform(data):
                        result = json.loads(self.vosk_recognizer.Result())
                        text = result.get("text", "").strip().lower()
                        if text:
                            print(f"🎤 Распознано: {text}")
                            self.log_to_gui(f"🎤 Распознано (голос): {text}")
                            self.process_command(text)
        except Exception as e:
            self.log_to_gui(f"❌ Ошибка аудиопотока: {e}")

    def process_command(self, command: str):
        command = command.strip().lower()
        if not command:
            return

        if any(word in command for word in ['справка', 'помощь', 'помоги', 'что ты умеешь']):
            self.show_help()
            return

        if command.startswith("я ") or command.startswith("это "):
            name = command.replace("я ", "").replace("это ", "").strip()
            self.set_profile(name)
            return

        if "смени профиль" in command or "переключи профиль" in command:
            self.set_profile(command)
            return

        if self.is_recognizing_faces:
            if any(word in command for word in ['стоп', 'выход', 'отбой', 'хватит']):
                self.is_recognizing_faces = False
                self.speak("Выхожу из режима распознавания лиц.")
            return

        if any(word in command for word in ['распознавание', 'запусти камеру', 'узнай меня', 'камера']):
            self.speak("Запускаю распознавание лиц. Скажите 'стоп' или нажмите кнопку в окне камеры.")
            self.is_recognizing_faces = True
            threading.Thread(target=self.recognize_face_loop, daemon=True).start()
            return

        if any(word in command for word in ['добавь лицо', 'добавь человека', 'запомни меня', 'добавить']):
            self.speak("Смотрите в камеру. Я сделаю несколько снимков.")
            self.add_person_by_voice()
            return

        # ИСПРАВЛЕНИЕ: Добавлено 'in command' к слову 'список'
        if 'список лиц' in command or 'кто в базе' in command or 'список' in command:
            self.show_face_list()
            return

        if any(word in command for word in ['привет', 'здравствуй', 'добрый день']):
            hour = datetime.now().hour
            if 6 <= hour < 12:
                self.speak("Доброе утро!")
            elif 12 <= hour < 18:
                self.speak("Добрый день!")
            elif 18 <= hour < 23:
                self.speak("Добрый вечер!")
            else:
                self.speak("Доброй ночи!")
            return

        if 'как дела' in command:
            self.speak("Спасибо, у меня всё отлично. Я готов помочь.")
            return

        if any(word in command for word in ['время', 'который час', 'сколько времени']):
            now = datetime.now().strftime("%H часов %M минут")
            self.speak(f"Сейчас {now}")
            return

        if any(word in command for word in ['дата', 'число', 'какое сегодня']):
            months = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября",
                      "ноября", "декабря"]
            now = datetime.now()
            self.speak(f"Сегодня {now.day} {months[now.month - 1]} {now.year} года")
            return

        if 'браузер' in command:
            self.speak("Открываю браузер")
            webbrowser.open("https://ya.ru")
            return

        if any(word in command for word in ['найди', 'поищи', 'погугли']):
            for kw in ['найди', 'поищи', 'погугли']:
                if kw in command:
                    query = command.split(kw, 1)[1].strip()
                    if query:
                        self.speak(f"Ищу {query} в Яндексе")
                        webbrowser.open(f"https://yandex.ru/search/?text={query}")
                        return
            self.speak("Что именно вы хотите найти?")
            return

        if 'погода' in command:
            self.speak("Открываю погоду")
            webbrowser.open("https://yandex.ru/pogoda")
            return

        if any(word in command for word in ['анекдот', 'шутка', 'рассмеши']):
            jokes = [
                "Почему программисты путают Хэллоуин и Рождество? Потому что 31 OCT == 25 DEC.",
                "Идёт медведь по лесу, видит — машина горит. Сел в неё и сгорел.",
                "Штирлиц открыл окно — дуло. Штирлиц закрыл окно — дуло исчезло."
            ]
            self.speak(random.choice(jokes))
            return

        if any(word in command for word in ['стоп', 'выход', 'пока', 'отбой', 'закрыть программу']):
            self.speak("До свидания!")
            self.cleanup()
            self.root.quit()
            sys.exit(0)

        if "друг" in command:
            parts = command.split("друг", 1)
            command_part = parts[1].strip() if len(parts) > 1 else ""
            if command_part:
                self.process_command(command_part)
            else:
                self.speak("Слушаю вас. Скажите команду или нажмите 'Полная справка'.")
            return

        self.log_to_gui("⚠️ Команда не распознана. Проверьте ввод или нажмите 'Полная справка'.")

    def load_face_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'rb') as f:
                    self.known_faces = pickle.load(f)
                self.log_to_gui(f"📂 Загружено {len(self.known_faces)} человек в базу лиц.")
            except Exception:
                self.known_faces = []

    def save_face_data(self):
        with open(DATA_FILE, 'wb') as f:
            pickle.dump(self.known_faces, f)

    def get_face_features(self, face_img):
        try:
            gray = cv2.cvtColor(face_img, cv2.COLOR_BGR2GRAY)
            gray = cv2.resize(gray, (64, 64))
            gray = cv2.GaussianBlur(gray, (5, 5), 0)
            gray = cv2.equalizeHist(gray)
            reduced = cv2.resize(gray, (32, 32))
            features = reduced.flatten().astype(np.float32) / 255.0
            return features
        except Exception as e:
            print(f"Ошибка извлечения признаков: {e}")
            return None

    def add_person_by_voice(self):
        name = simpledialog_askstring("Добавление лица", "Введите имя человека:")
        if name:
            self.add_person(name.strip().capitalize())
        else:
            self.speak("Добавление отменено.")

    def add_person(self, name):
        self.log_to_gui(f"\n--- Добавление: {name} ---")
        self.speak(f"Добавляем человека {name}. Смотрите в камеру и поворачивайте голову.")
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.speak("Ошибка камеры.")
            return

        features_list = []
        samples = 20

        try:
            while len(features_list) < samples and self.is_listening:
                ret, frame = cap.read()
                if not ret: break
                frame = cv2.flip(frame, 1)
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self.face_cascade.detectMultiScale(gray, 1.1, 5, minSize=(100, 100))

                for (x, y, w, h) in faces:
                    roi = frame[y:y + h, x:x + w]
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    feat = self.get_face_features(roi)
                    if feat is not None and len(features_list) < samples:
                        features_list.append(feat)
                        self.log_to_gui(f"Снимок {len(features_list)}/{samples}")
                        time.sleep(0.1)
                    cv2.putText(frame, f"{len(features_list)}/{samples}", (x, y - 10),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)

                cv2.putText(frame, f"Name: {name}", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
                cv2.imshow("Добавление лица", frame)

                key = cv2.waitKey(1) & 0xFF
                if key == ord(' '):
                    break
                elif key == ord('q'):
                    return
        finally:
            cap.release()
            cv2.destroyAllWindows()

        if len(features_list) < 5:
            self.speak("Недостаточно снимков, попробуйте снова.")
            return

        avg = np.mean(features_list, axis=0)
        self.known_faces.append({'name': name, 'features': avg})
        self.save_face_data()
        self.speak(f"Запомнил, как {name}.")

    def recognize_face_loop(self):
        if len(self.known_faces) == 0:
            self.speak("База лиц пуста. Сначала добавьте человека.")
            self.is_recognizing_faces = False
            return

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.speak("Ошибка камеры.")
            self.is_recognizing_faces = False
            return

        self.log_to_gui("\n--- Режим распознавания лиц ---")
        last_time = 0
        interval = 2.0

        try:
            while self.is_recognizing_faces:
                ret, frame = cap.read()
                if not ret: break
                frame = cv2.flip(frame, 1)
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self.face_cascade.detectMultiScale(gray, 1.1, 5, minSize=(100, 100))

                now = time.time()
                for (x, y, w, h) in faces:
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

                    if now - last_time > interval:
                        last_time = now
                        roi = frame[y:y + h, x:x + w]
                        feat = self.get_face_features(roi)

                        if feat is not None:
                            best_name = None
                            best_sim = -1
                            for person in self.known_faces:
                                try:
                                    if len(feat) != len(person['features']):
                                        continue

                                    dot = np.dot(feat, person['features'])
                                    n1 = np.linalg.norm(feat)
                                    n2 = np.linalg.norm(person['features'])
                                    if n1 > 0 and n2 > 0:
                                        sim = dot / (n1 * n2)
                                        if sim > best_sim:
                                            best_sim = sim
                                            best_name = person['name']
                                except Exception:
                                    continue

                            if best_name and best_sim > self.similarity_threshold:
                                label = best_name
                                color = (0, 255, 0)
                                cv2.rectangle(frame, (x, y), (x + w, y + h), color, 3)

                                if best_name != self.last_spoken_name:
                                    self.auto_switch_profile_by_face(best_name)
                                    self.last_spoken_name = best_name
                                    self.last_spoken_time = now
                                    self.speak(f"Ура, это {best_name}!")

                                self.log_to_gui(f"✅ РАСПОЗНАНО: {best_name} (сходство: {best_sim:.2f})")
                            else:
                                label = UNKNOWN_NAME
                                color = (0, 0, 255)
                                cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)
                                if now - self.last_spoken_time > self.speak_cooldown:
                                    self.last_spoken_time = now
                                    self.speak(UNKNOWN_NAME)

                            frame = self.draw_text_with_pil(frame, label, (x, y - 40), color)

                cv2.putText(frame, "Нажмите 'q' или скажите 'стоп' для выхода", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6,
                            (255, 255, 255), 2)
                cv2.imshow("Распознавание лиц", frame)

                if cv2.waitKey(1) & 0xFF == ord('q'):
                    self.is_recognizing_faces = False
                    break
        except Exception as e:
            self.log_to_gui(f"❌ Ошибка в цикле распознавания: {e}")
        finally:
            cap.release()
            cv2.destroyAllWindows()
            self.log_to_gui("Распознавание завершено.")
            self.is_recognizing_faces = False

    def auto_switch_profile_by_face(self, recognized_name):
        target_profile = None
        for key in USER_PROFILES:
            if key.lower() in recognized_name.lower() or recognized_name.lower() in key.lower():
                target_profile = key
                break

        if target_profile and target_profile != self.current_profile_name:
            self.log_to_gui(
                f"👁️ Лицо распознано как '{recognized_name}'. Автоматически переключаю профиль на '{target_profile}'.")
            self.set_profile(target_profile)

    def show_face_list(self):
        if self.known_faces:
            names = [p['name'] for p in self.known_faces]
            text = "В базе: " + ", ".join(names)
            self.speak(text)
        else:
            self.speak("База лиц пуста.")

    def cleanup(self):
        self.is_listening = False
        self.speak_queue.put(None)
        if self.listen_thread:
            self.listen_thread.join(timeout=1)


def simpledialog_askstring(title, prompt):
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.geometry("300x100")
    dialog.transient()
    dialog.grab_set()

    result = tk.StringVar()

    tk.Label(dialog, text=prompt).pack(pady=5)
    entry = ttk.Entry(dialog, textvariable=result, width=30)
    entry.pack(pady=5)
    entry.focus()

    def on_ok():
        dialog.destroy()

    entry.bind("<Return>", lambda e: on_ok())

    ttk.Button(dialog, text="OK", command=on_ok).pack(pady=5)

    dialog.wait_window()
    return result.get()


def main():
    try:
        assistant = SmartAssistant()
        assistant.speak(f"Система запущена. Профиль: {assistant.current_profile_name}.")

        assistant.root.protocol("WM_DELETE_WINDOW",
                                lambda: [assistant.speak("До свидания!"), assistant.cleanup(), assistant.root.quit(),
                                         sys.exit(0)])
        assistant.root.mainloop()

    except KeyboardInterrupt:
        print("\n⏹️ Прервано пользователем")
    finally:
        if 'assistant' in locals():
            assistant.cleanup()
        cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
