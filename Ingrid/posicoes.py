import pyautogui
import time

time.sleep(3)

posicao = pyautogui.position()
print(f"A posição é: {posicao}.")

#tela = pyautogui.size()
#print(f"Dimensões: {tela.width} x {tela.height}.")