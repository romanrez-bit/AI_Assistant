import os
import cv2
import pickle
import face_recognition
import numpy as np
import pyttsx3
import threading
import time

# Конфигурация
DATA_FILE = "face_encodings.pkl"  # Файл, где хранятся "слепки" лиц
UNKNOWN_NAME = "объект не опознан"  # Фраза для чужих


class SmartFaceRecognizer:
    def __init__(self):
        self.known_face_encodings = []  # Список кодов лиц
        self.known_face_names = []  # Список имен (порядок совпадает)

        # Голос
        self.tts_engine = pyttsx3.init()
        self.tts_engine.setProperty('rate', 150)
        self.setup_voice()

        # Для предотвращения повторений
        self.last_spoken_name = ""
        self.last_spoken_time = 0
        self.speak_cooldown = 2.5  # Не здороваться каждую секунду

        # Загрузка сохраненных данных при старте
        self.load_encodings()

    def setup_voice(self):
        """Пытаемся найти русский голос"""
        try:
            voices = self.tts_engine.getProperty('voices')
            for voice in voices:
                if 'russian' in voice.name.lower() or 'ru' in voice.id.lower():
                    self.tts_engine.setProperty('voice', voice.id)
                    print(f"Используется голос: {voice.name}")
                    break
        except:
            pass

    def speak(self, text, is_welcome=False):
        """Озвучка в отдельном потоке"""
        # Для приветствий убираем проверку на повтор, для остальных оставляем
        if not is_welcome:
            current_time = time.time()
            if text == self.last_spoken_name and current_time - self.last_spoken_time < self.speak_cooldown:
                return
            self.last_spoken_name = text
            self.last_spoken_time = current_time

        def _speak():
            try:
                # Создаем новый движок на каждую фразу, чтобы избежать зависаний
                engine = pyttsx3.init()
                engine.setProperty('rate', 150)
                # Настройка голоса (попытка русский)
                for voice in engine.getProperty('voices'):
                    if 'russian' in voice.name.lower() or 'ru' in voice.id.lower():
                        engine.setProperty('voice', voice.id)
                        break
                engine.say(text)
                engine.runAndWait()
                engine.stop()
            except Exception as e:
                print(f"Voice error: {e}")

        threading.Thread(target=_speak, daemon=True).start()

    def load_encodings(self):
        """Загрузка базы знаний из файла"""
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'rb') as f:
                    data = pickle.load(f)
                    self.known_face_encodings = data['encodings']
                    self.known_face_names = data['names']
                print(f"Загружено {len(self.known_face_names)} персон.")
                for name in self.known_face_names:
                    print(f" - {name}")
            except Exception as e:
                print(f"Ошибка загрузки: {e}")
                self.known_face_encodings = []
                self.known_face_names = []
        else:
            print("База лиц не найдена. Сначала добавьте людей.")

    def save_encodings(self):
        """Сохранение базы знаний"""
        data = {
            'encodings': self.known_face_encodings,
            'names': self.known_face_names
        }
        with open(DATA_FILE, 'wb') as f:
            pickle.dump(data, f)
        print(f"Сохранено {len(self.known_face_names)} персон.")

    def add_person(self, name):
        """Режим обучения: снимаем лицо с камеры и добавляем в базу"""
        print(f"\n--- Добавление персоны: {name} ---")
        cap = cv2.VideoCapture(0)

        if not cap.isOpened():
            print("Ошибка камеры")
            return False

        face_encodings_for_person = []

        print("Смотрите в камеру. Сканирование лица...")
        print("Для завершения нажмите 'Space' (Пробел). Для отмены 'q'.")

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            # Показываем кадр
            cv2.imshow("Добавление лица - Нажмите Пробел", frame)

            # Ищем лица на каждом кадре
            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            face_locations = face_recognition.face_locations(rgb_frame, model="hog")  # hog быстрее на CPU

            if face_locations:
                # Кодируем (превращаем в вектор) первое найденное лицо
                encodings = face_recognition.face_encodings(rgb_frame, face_locations)
                if encodings:
                    face_encodings_for_person.append(encodings[0])
                    # Рисуем рамку
                    for (top, right, bottom, left) in face_locations:
                        cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
                        cv2.putText(frame, f"Scanning {name}...", (left, top - 10),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)
                    cv2.imshow("Добавление лица - Нажмите Пробел", frame)
                    print(f"Сделано снимков: {len(face_encodings_for_person)}", end='\r')

            key = cv2.waitKey(1) & 0xFF
            if key == ord(' '):  # Пробел - закончить съемку
                break
            elif key == ord('q'):
                cap.release()
                cv2.destroyAllWindows()
                print("\nОтмена добавления.")
                return False

        cap.release()
        cv2.destroyAllWindows()

        if len(face_encodings_for_person) == 0:
            print("Не удалось захватить лицо. Попробуйте снова.")
            return False

        # Усредняем все снимки лица для лучшей точности (опционально)
        avg_encoding = np.mean(face_encodings_for_person, axis=0)
        self.known_face_encodings.append(avg_encoding)
        self.known_face_names.append(name)
        self.save_encodings()

        print(f"\nПерсона '{name}' успешно добавлена! (Всего снимков: {len(face_encodings_for_person)})")
        self.speak(f"Запомнил, как {name}", is_welcome=True)
        return True

    def recognize_face(self):
        """Основной цикл распознавания"""
        if len(self.known_face_names) == 0:
            print("База лиц пуста. Пожалуйста, сначала добавьте людей.")
            return

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("Ошибка камеры")
            return

        print("\n--- Режим распознавания ---")
        print("При обнаружении знакомого лица, программа скажет 'Привет, Имя'")
        print("Если лицо не опознано: 'объект не опознан'")
        print("Для выхода нажмите 'q'")

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            # Уменьшаем кадр для скорости (опционально)
            small_frame = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)
            rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

            # Находим все лица на кадре
            face_locations = face_recognition.face_locations(rgb_small_frame, model="hog")
            face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

            # Список имен для отображения на этом кадре
            face_names = []

            for face_encoding in face_encodings:
                # Сравниваем текущее лицо с базой
                matches = face_recognition.compare_faces(self.known_face_encodings, face_encoding, tolerance=0.5)
                name = UNKNOWN_NAME

                # Ищем лучшее совпадение
                face_distances = face_recognition.face_distance(self.known_face_encodings, face_encoding)
                if len(face_distances) > 0:
                    best_match_index = np.argmin(face_distances)
                    if matches[best_match_index]:
                        name = self.known_face_names[best_match_index]
                        # Озвучиваем имя (с проверкой на повтор)
                        self.speak(f"Привет, {name}")

                face_names.append(name)

            # Отрисовка (масштабируем обратно координаты)
            for (top, right, bottom, left), name in zip(face_locations, face_names):
                # Масштабируем координаты обратно в 100%
                top *= 2
                right *= 2
                bottom *= 2
                left *= 2

                # Цвет рамки: Зеленый если знакомый, Красный если нет
                color = (0, 255, 0) if name != UNKNOWN_NAME else (0, 0, 255)

                cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
                cv2.rectangle(frame, (left, bottom - 35), (right, bottom), color, cv2.FILLED)
                cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_DUPLEX, 0.8, (255, 255, 255), 1)

            cv2.imshow('Умная Камера - Распознавание лиц', frame)

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()
        print("Распознавание завершено.")


def show_menu():
    print("\n" + "=" * 40)
    print("   СИСТЕМА РАСПОЗНАВАНИЯ ЛИЦ С ГОЛОСОМ")
    print("=" * 40)
    print("1. Обучить нейросеть (Добавить человека)")
    print("2. Запустить распознавание (Камера + Голос)")
    print("3. Посмотреть базу лиц")
    print("4. Выйти")
    print("-" * 40)


def main():
    recognizer = SmartFaceRecognizer()

    while True:
        show_menu()
        choice = input("Ваш выбор: ").strip()

        if choice == '1':
            name = input("Введите имя для обучения: ").strip()
            if name:
                recognizer.add_person(name)
            else:
                print("Имя не может быть пустым.")
        elif choice == '2':
            recognizer.recognize_face()
        elif choice == '3':
            if len(recognizer.known_face_names) == 0:
                print("База лиц пуста.")
            else:
                print("\nСписок обученных персон:")
                for i, name in enumerate(recognizer.known_face_names, 1):
                    print(f"{i}. {name}")
        elif choice == '4':
            print("До свидания!")
            break
        else:
            print("Неверный ввод. Попробуйте снова.")


if __name__ == "__main__":
    main()