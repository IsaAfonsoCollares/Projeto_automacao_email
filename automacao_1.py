#importando bibliotecas
import pyautogui #ferramenta para automação
import time #para trabalhar com tempo, tempo de espera, etc
import pyperclip #permite copiar e colar usando python
import pandas as pd #ferramente para análise de dados

#usando o pyautogui para abrir o navegador automaticamente
pyautogui.PAUSE=1
pyautogui.press("winleft")#pressiona o botão windows no teclado
pyautogui.write("chrome")#escreve "chrome" na barra de pesquisa
pyautogui.press("enter")#pressiona enter
#tirar texto abaixo do comentário se quiser abrir uma nova guia
#pyautogui.alert("Vai começar, aperte ok e não mexa em nada") #cria um alerta na tela do computador
#pyautogui.hotkey("ctrl", "t")#executa um atalho de teclado

#abrir link da planilha no drive
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
#usamos o piperclip para armazenar o link no atalho "crtl c"
pyautogui.hotkey('ctrl','v')#utiliza o atalho "ctrl v"
pyautogui.press("enter")
time.sleep(5) 
#usamos time para que o programa aguarde 3 segundos antes de rodar a próxima linha de código, assim podemos carregar toda a página antes de continuar
pyautogui.click(378, 300, clicks =2)
time.sleep(5)
pyautogui.click(346, 480)
pyautogui.click(1169, 186)
pyautogui.click(907, 624)
time.sleep(5)
#usamos o pyautogui para clicar em posições específicas da tela 
#automatizando o processo de abrir, utilizar e fechar programas

#Ler e manusear o arquivo
df = pd.read_excel(r'C:\Users\isabe\Downloads\Vendas - Dez.xlsx') #utiliza o pandas para ler e armazenar o arquivo dentro da variável df
faturamento = df['Valor Final'].sum()#soma todos os dados da coluna Valor Final
qtde_produtos=df['Quantidade'].sum()#soma todos os dados da coluna Quantidade 
time.sleep(5)
#Enviar o relatório com as informações por email automaticamente
pyperclip.copy('isabela_afonso_collares@hotmail.com')
pyautogui.click(516, 748)
pyautogui.click(120, 334)
pyautogui.click(106, 109)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab', presses=3)
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')
texto=f"""
Prezados, bom dia
O faturamento de ontem for de: R${faturamento:,.2f}.
A quantidade de produtos foi de: {qtde_produtos:,}.

Abs,
Isabela Afonso Collares"""
pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
time.sleep(3)
pyautogui.click(1279,55)