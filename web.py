from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import win32com.client as win32

#navegador selenium
nav = webdriver.Chrome()

# importar/visualizar a base de dados
tabela_produtos = pd.read_excel(r"buscas.xlsx")

#função de filtro que verifica a utilização de termos que devem ser excluídos
def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome:
            tem_termos_banidos = True
    return tem_termos_banidos


#função de filtro que verifica a existencia de todos os termos dentro da busca
def verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome):
    tem_todos_termos_produto = True
    for palavra in lista_termos_nome_produto:
        if palavra not in nome:
            tem_todos_termos_produto = False
    return tem_todos_termos_produto


#função de busca na loja do Google
def busca_google_shooping(nav, produto,termos_banidos,preco_minimo,preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    lista_produtos = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    #entrar no google
    nav.get("https://www.google.com/")
    nav.find_element(By.XPATH,'//*[@id="APjFqb"]').send_keys(produto, Keys.ENTER)

    #entrar na aba shopping
    nav.find_element(By.XPATH,'//*[@id="hdtb-sc"]/div/div/div[1]/div/div[2]/a/div').click()

    #pegar informações do produto
    lista_resultados = nav.find_elements(By.CLASS_NAME,'i0X6df')
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME,'tAxDx').text
        nome = nome.lower()

        #analisar se não tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        #analisar se tem TODOS os termos do nome do produto
        tem_todos_termos_produto = verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome)

        #selecionar apenas elementos corretos
        if not tem_termos_banidos and tem_todos_termos_produto:
            preco = resultado.find_element(By.CLASS_NAME,'a8Pemb').text
            preco = preco.replace("R$" , "").replace(" " , "").replace("." , "").replace("," , ".").replace("+impostos","")
            if preco != '':
                preco = float(preco)

                #se o preco está entre minimo e maximo
                if preco_minimo <= preco <= preco_maximo:
                    elemento_referencia = resultado.find_element(By.CLASS_NAME, 'bONr3b')
                    elemento_pai = elemento_referencia.find_element(By.XPATH,'..')
                    link = elemento_pai.get_attribute('href')
                    lista_produtos.append((nome, preco, link))
    return lista_produtos                   


#função de busca na loja do buscapé
def busca_buscape(nav, produto,termos_banidos,preco_minimo,preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    lista_produtos = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    nav.get("https://www.buscape.com.br/")
    nav.find_element(By.XPATH,
            '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)
    lista_resultados = nav.find_elements(By.CLASS_NAME,'ProductCard_ProductCard_Inner__gapsh')
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'ProductCard_ProductCard_Name__U_mUQ').text
        nome = nome.lower()

        #analisar se não tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        #analisar se tem TODOS os termos do nome do produto
        tem_todos_termos_produto = verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome)     

        if not tem_termos_banidos and tem_todos_termos_produto:
            preco = resultado.find_element(By.CLASS_NAME,'Text_MobileHeadingS__HEz7L').text
            preco = preco.replace("R$" , "").replace(" " , "").replace("." , "").replace("," , ".").replace("+impostos","")
            if preco != '':
                preco = float(preco)

                #se o preco está entre minimo e maximo
                if preco_minimo <= preco <= preco_maximo:
                    link = resultado.get_attribute('href')        
                    lista_produtos.append((nome,preco,link))
    return lista_produtos


tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    #pesquisar pelo produto
    produto = tabela_produtos.loc[linha,"Nome"]
    termos_banidos = tabela_produtos.loc[linha,"Termos banidos"]
    preco_minimo = tabela_produtos.loc[linha,"Preço mínimo"]
    preco_maximo = tabela_produtos.loc[linha,"Preço máximo"]

    lista_produtos_google_shopping = busca_google_shooping(nav, produto,termos_banidos,preco_minimo,preco_maximo)
    if lista_produtos_google_shopping:
        tabela_google_shooping = pd.DataFrame(lista_produtos_google_shopping, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas , tabela_google_shooping])
    else:
        tabela_google_shooping = None

    lista_produtos_buscape = busca_buscape(nav, produto,termos_banidos,preco_minimo,preco_maximo)
    if lista_produtos_buscape:
        tabela_buscape = pd.DataFrame(lista_produtos_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas , tabela_buscape])
    else:
        tabela_buscape = None

#salva as ofertas em um arquivo excel
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)

#enviar notificação para o email do responsável
if len(tabela_ofertas) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "@gmail.com"  #excluir o indice
    mail.Subject = f'Produto(s) encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f'''
    <p>Prezados, bom dia</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada</p>
    <p>Att.,</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Vinicius</p>
    '''
    # Enviar o email
    mail.Send()
    nav.quit()



