import pyautogui 
import time
import pyperclip
import tkinter as tk
from tkinter import ttk

#pyautogui.locateCenterOnScreen
pyautogui.PAUSE = 1 #tempo de pausa entre cada comando do pyautogui

empresa = 2063
filial = 1
ano = 2025
mes = "01"

empresafilial = f"{empresa}-{filial}"
empresafilial2 = f"{empresa}.{filial}"

pyautogui.click(713, 37)

# MUDAR A EMPRESA NO ATHENAS
pyautogui.press("F8")
pyautogui.click(34, 94)
pyautogui.click(324, 85)
pyautogui.write(empresafilial)
pyautogui.press("Enter")
time.sleep(1)

# MUDAR O MÊS E O ANO
pyautogui.click(159, 62) #vai para o pessoal
time.sleep(1)
pyautogui.click(460, 89)
time.sleep(1)
pyautogui.write("Janeiro")
pyautogui.press("Enter")
time.sleep(1)
pyautogui.click(457, 109)
pyautogui.write("2025")
pyautogui.press("Enter")

# IR PARA O "RELATÓRIOS FOLHA"
pyautogui.click(147, 112) #abre relat. folha
time.sleep(7)

# MUDAR O TIPO DO RELATÓRIO
pyautogui.click(368, 118)
pyautogui.hotkey("ctrl", "a")
pyautogui.press("5")

# VISUALIZAR
pyautogui.click(349, 709)
time.sleep(7)

pyautogui.click(42, 33) #salvar
pyautogui.click(974, 345)
time.sleep(2)

pyautogui.click(912, 59) #muda o diretorio

caminho = f"C:\\Users\\regina.santos\\Desktop\\Pessoal\\{empresafilial2}\\Provisões"
pyperclip.copy(caminho)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("Enter")

pyautogui.click(745, 577)
pyautogui.write(f"{empresafilial2}-{ano}-{mes}-PF")
pyautogui.press("Enter")
pyautogui.click(634, 397)
pyautogui.press("Enter")
time.sleep(3)

pyautogui.click(510, 37) #close

# MUDAR O TIPO DO RELATÓRIO
pyautogui.click(368, 118)
pyautogui.hotkey("ctrl", "a")
pyautogui.press("6")

# VISUALIZAR
pyautogui.click(349, 709)
time.sleep(7)

pyautogui.click(42, 33) #salvar
pyautogui.click(974, 345)
time.sleep(2)

pyautogui.click(912, 59) #muda o diretorio

caminho = f"C:\\Users\\regina.santos\\Desktop\\Pessoal\\{empresafilial2}\\Provisões"
pyperclip.copy(caminho)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("Enter")

pyautogui.click(745, 577)
pyautogui.write(f"{empresafilial2}-{ano}-{mes}-P13")
pyautogui.press("Enter")
pyautogui.click(634, 397)
pyautogui.press("Enter")
time.sleep(3)

pyautogui.click(510, 37) #close
pyautogui.click(1103, 706)