import cv2
import numpy as np

# Создаём чёрное изображение
img = np.zeros((400, 600, 3), dtype=np.uint8)
cv2.putText(img, "Press Q to exit", (50, 200), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)

window_name = "Key Test - Click here and press Q"
cv2.namedWindow(window_name)

while True:
    cv2.imshow(window_name, img)
    key = cv2.waitKey(50) & 0xFF

    print(f"Нажата клавиша: {key} (символ: {chr(key) if 32 <= key <= 126 else '?'})")

    if key == ord('q') or key == ord('Q'):
        print("Q нажата - выход")
        break

cv2.destroyAllWindows()