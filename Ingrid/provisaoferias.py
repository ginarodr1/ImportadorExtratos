import pyautogui 
import time
import keyboard

pyautogui.click(368, 118)
pyautogui.hotkey("ctrl", "a")
pyautogui.press("5")
pyautogui.click(349, 709) #visualizar

time.sleep(7)

pyautogui.click(42, 33) #salvar
pyautogui.click(974, 345)
time.sleep(2)

pyautogui.click(912, 59)
pyautogui.write("K:\003 - Pessoal\2025\{empresafilial}\25 - Provisões")
pyautogui.press("Enter")