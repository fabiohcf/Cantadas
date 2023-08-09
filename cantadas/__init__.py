from time import sleep
import webbrowser
import pyautogui
import pyperclip
import openpyxl
from random import randint


# ler cantada do arquivo 'cantadas.xlsx' e sortear aleatoriamente
arq = openpyxl.load_workbook('cantadas.xlsx')  # arquivo.xlsx
plan = arq['aba1']  # aba
num = randint(1, 99)
cantadadodia = plan[f'A{num}'].value  # valor da celula[A1}
print(num)
print(cantadadodia)

# abrir gmail em nova aba
webbrowser.open('https://mail.google.com/')
sleep(5)
pyautogui.press('enter')

# atalho escrever email ('c')
pyautogui.press('c')
sleep(1)

# Atenção:
'''
Ativar os atalhos do teclado no Gmail:

No computador, acesse o Gmail.
No canto superior direito, clique em Configurações Configurações e Ver todas as configurações.
Clique em Configurações.
Role para baixo até a seção "Atalhos do teclado".
Selecione Atalhos do teclado ativados.
Na parte inferior da página, clique em Salvar alterações.
'''

# preencher destinatário
pyperclip.copy('destinatario@gmail.com') # insira o e-mail do destinatário
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
pyautogui.press('tab')
sleep(1)

# assunto
pyperclip.copy('Cantada do dia') # escolha o assunto
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

# corpo do email
pyperclip.copy(cantadadodia)
pyautogui.hotkey('ctrl', 'v')

# enviar email
sleep(2)
pyautogui.hotkey('ctrl', 'enter')

# fechar aba
sleep(5)
pyautogui.hotkey('ctrl', 'f4')

# finalizar
print('E-mail enviado com sucesso!')

