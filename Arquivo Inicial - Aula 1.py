import pyautogui as pa
import time as ti
import pyperclip as pc

# Passo 01: Entrar no sistema da empresa (no link drive)

pa.PAUSE = 1

pa.hotkey('ctrl', 't')
ti.sleep(5)
pc.copy('https://drive.google.com/drive/u/0/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pa.hotkey('ctrl', 'v')
pa.press('enter')

# Passo 02: Navegar até o local do relatório

ti.sleep(5)
pa.click(x=329, y=254)
pa.click(x=329, y=254, clicks=2)

# Passo 03: Exportar Relatório (Fazer Download)

ti.sleep(5)
pa.click(x=329, y=254)
pa.click(x=1159, y=154)
ti.sleep(3)
pa.click(x=913, y=592)
ti.sleep(10)

# Passo 04: Calcular os indicadores (faturamento e quantidade)

import pandas as pd

tabela = pd.read_excel(r'C:\Users\Lucy Minene\Downloads\Vendas - Dez.xlsx')

display(tabela)

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

print(faturamento)
print(quantidade)

# Passo 05: Enviar um email para diretoria

# Abrir uma nova aba e entrar no gmail

pa.hotkey('ctrl', 't')
pc.copy('https://mail.google.com/mail/u/0/#inbox')
pa.hotkey('ctrl', 'v')
pa.press('enter')

# Clicar no botão escrever

ti.sleep(15)
pa.click(x=48, y=151)


# Preencher as informações do email

# Destinatário
ti.sleep(5)
pa.write('brunopalhanog@gmail.com')
pa.press('tab')
pa.press('tab')

# Assunto

pc.copy('Relatório de Faturamento')
pa.hotkey('ctrl', 'v')
pa.press('tab')

# Corpo do email

ti.sleep(3)
texto = f"""

Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos vendida foi de: {quantidade:,}
    
Sds, Bruno Palhano

"""

pc.copy(texto)
pa.hotkey('ctrl', 'v')

# Enviar email
ti.sleep(3)
pa.hotkey('ctrl', 'enter')
