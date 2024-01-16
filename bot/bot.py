from selenium import webdriver
from time import sleep                                                  
from selenium.webdriver.common.by import By                                
from selenium.webdriver.support.ui import Select
from botcity.web.parsers import table_to_dict
from botcity.plugins.excel import BotExcelPlugin
from envia_email import email_conclusao
from cofre import Credenciais


# Cria instância com o webdriver
bot = webdriver.Chrome(r'C:\acme\chromedriver-win64\chromedriver.exe')


# Cria instâncis com o excel e adiciona cabeçalho
excel = BotExcelPlugin()
excel.add_row(["WIID", "Description", "Type", "Status", "Date"])


# Contador e acumulador 
contador = 1
w1 = w2 = w3 = w4 = w5 = 0
lista_tabela = []


# 1. Criar um Acesso no sistema Acme System: https://acme-test.uipath.com/;
bot.get("https://acme-test.uipath.com/login")
bot.maximize_window()
bot.implicitly_wait(2)


# Efetua o login
bot.find_element(By.XPATH, "//input[@id='email']").send_keys(Credenciais.USUARIO)
bot.find_element(By.XPATH, "//input[@id='password']").send_keys(Credenciais.SENHA)
bot.find_element(By.XPATH, "/html/body/div/div[2]/div/div/div/form/button").click()
bot.implicitly_wait(2)


# 2) Clicar na Opção “work items”;
bot.find_element(By.XPATH, '//*[@id="dashmenu"]/div[2]/a/button').click()
bot.implicitly_wait(2)


# 3) Extrair os dados da Tabela (Plus: extrair os dados de todas as Abas);
while True:
    try:
        # Pular o click na primeira aba
        if contador > 1:
            bot.find_element(By.XPATH, "//a[@rel='next']").click()
            bot.implicitly_wait(2)

        # Captura a tabela
        tabela_dados = bot.find_element(By.CLASS_NAME, "table")
        tabela_dados = table_to_dict(table=tabela_dados)

        # Extende cada tabela
        lista_tabela.extend(tabela_dados) 
        
        contador+=1

    except Exception:
        print("Fim das páginas")
        break


# 4) Criar um arquivo do Excel com todos os dados da Tabela que foi extraída anteriormente;
    
for linha in lista_tabela:
    excel.add_row([linha['wiid'], linha['description'], linha['type'], linha['status'], linha['date']])

    # Contadores de tipo
    if linha['type'] == "WI1":
        w1+=1
    elif linha['type'] == "WI2":
        w2+=1
    elif linha['type'] == "WI3":
        w3+=1
    elif linha['type'] == "WI4":
        w4+=1
    elif linha['type'] == "WI5":
        w5+=1

excel.write(r"C:\acme\relatorio.xlsx")


# 5) Exibir uma mensagem com o total de Linhas da Tabela extraída
print("Total de Linhas da tabela extraída: ", len(lista_tabela))

# 6) Exibir o total de linhas baseado na coluna Type: WI1 = 50, WI2=20 e assim por diante...
print(f"Total de Linhas WI1: {w1} \n Total de Linhas WI2: {w2} \n Total de Linhas WI3: {w3} \n Total de Linhas WI4: {w4} \n Total de Linhas WI5: {w5}")

# 7) Enviar um email com a Planilha construída (preferencialmente para você mesmo), anexando o seu Resultado, com uma mensagem que deve ser
# informada de alguma forma (exemplo: por input de dados, por variável criada, por leitura de arquivo de configuração...).
email_conclusao(destino = "eltonpduarte@gmail.com")

# Fim
bot.close()