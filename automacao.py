from itertools import chain
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import re


# Define a URL do site
url = "https://sisat.sefaz.pi.gov.br/sisatweb-web/consultaPublica"


# Define as colunas do data frame
colunas = "Nº da NF-,Data Emissão,UF Origem,Dados Emitente,UF Destino,Dados Destinatário,Imposto,Situação da cobrança,Tipo de Serviço".split(
    ","
)

# Armazena Ids/xpath dos elementos

card_consulta_publica = (
    "j_id8:tabContent_tabItem1:formControlPanel:j_id-1671860232_7c2a1c3d"
)

id_0 = "tabView:tabContent_tabItem2:formConsultaCobranca:tableCobranca:0:dataTableNfesDetalheRetidas_data"
_id = "tabView:tabContent_tabItem2:formConsultaCobranca:tableCobranca:{}:dataTableNfesDetalheRetidas"
input_periodo_de = (
    "tabView:tabContent_tabItem2:formConsultaCobranca:inputPeriodoInicio_input"
)
input_periodo_ate = (
    "tabView:tabContent_tabItem2:formConsultaCobranca:inputPeriodoFim_input"
)

input_cnpj_destinatario = (
    "tabView:tabContent_tabItem2:formConsultaCobranca:contribuinteCpfCnpj"
)

script = """ var spans = document.querySelectorAll('tr span'); spans.forEach(function(span) { if (span.textContent.trim() === 'Ver NF-es') { span.click(); } }); """
btn_consultar = '//*[@id="tabView:tabContent_tabItem2:formConsultaCobranca:j_id-69522232_8289b9c"]/span[2]'

link_paginate_next = '//*[@id="tabView:tabContent_tabItem2:formConsultaCobranca:tableCobranca_paginator_bottom"]/a[3]'

# Instancia o driver do selenium
driver = webdriver.Chrome()

# Cria lista para armazenar os dados
dados = []

# Contador para controlar o número do elemento na página
count = 0

# Lista dos CNPJs que serão pesquisados
cnpjs = str(input('Digite a lista de CNPJS (Um CNPJ seguido do outro com espaço entre eles): ')).split()

data_inicio = str(input('Digite a data de Inicio: ')).strip()
data_fim = str(input('Digite a data de fim: ')).strip()


# Função para aguardar o elemento
def wait(element, time=20):
    return WebDriverWait(driver, time).until(EC.element_to_be_clickable(element))

# Função para aguardar os elementos
def wait_all(element, time=20):
    return WebDriverWait(driver, time).until(
        EC.presence_of_all_elements_located((By.ID, element))
    )

# Inicializa o navegador
driver.get(url)


def preencher_dados(cnpj, data_inicio, data_fim):
    element = wait((By.ID, input_cnpj_destinatario))
    element.clear()
    element.send_keys(cnpj)
    time.sleep(1)

    element = wait((By.ID, input_periodo_de))
    element.send_keys(Keys.BACKSPACE)
    time.sleep(1)
    element.send_keys(data_inicio)
    
    element = wait((By.ID, input_periodo_ate))
    element.send_keys(Keys.BACKSPACE)
    time.sleep(1)
    element.send_keys(data_fim)
    time.sleep(1)


    try:
        wait((By.XPATH, btn_consultar)).click()
    except ElementClickInterceptedException:
        time.sleep(2)
        wait((By.XPATH, btn_consultar)).click()
    time.sleep(2)


def abrir_span():
    print("Aguardando NF-es...")
    wait((By.XPATH, "//td//span[text()='Ver NF-es']"), 40)
    driver.execute_script(script)
    time.sleep(5)


def abrir_todos(count):
    temp_list = []
    print(count)
    nfs = wait((By.ID, _id.format(count)), 1).find_elements(By.TAG_NAME, "tr")
    for nf in nfs:
        if nf.text.startswith("Nº da NF-e"):
            continue
        temp = [text.text.strip() for text in nf.find_elements(By.TAG_NAME, "td")]
        temp_list.append(temp)
    return temp_list


def paginacao():
    wait((By.XPATH, link_paginate_next)).click()


def pegar_dados(count):
    while True:
        try:
            abrir_span()
        except TimeoutException:
            print("Sem Notas Fiscais")
            return 
        for _ in range(10):
            try:                
                result = abrir_todos(count)
                print(result)
                print("-" * 50)
                print("Content length: ", len(result))
                dados.append(result)
                count += 1
            except TimeoutException:
                break
        if "ui-state-disabled" in wait((By.XPATH, link_paginate_next)).get_attribute(
            "class"
        ):
            wait((By.XPATH, '//*[@id="tabView:tabContent_tabItem2:formConsultaCobranca:tableCobranca_paginator_bottom"]/a[1]')).click()
            time.sleep(1)
            print("Voltando para primeira pagina")
            break
        try:
            paginacao()
        except ElementClickInterceptedException:
            time.sleep(1)
            paginacao()


#data_inicio = "01/10/2023"
#data_fim = "31/12/2023"
wait((By.ID, card_consulta_publica)).click()
for cnpj in cnpjs:
    print("Pegando o cnpj: ", cnpj)
    count = 0
    preencher_dados(cnpj, data_inicio, data_fim)
    pegar_dados(count)


lista_aplanada = list(chain.from_iterable(dados))
try:
    result = pd.DataFrame(lista_aplanada, columns=colunas)
    result['CNPJ Emitente'] = result['Dados Emitente'].apply(lambda x: ''.join(filter(str.isdigit, x)))
    result['CNPJ Destinatário'] = result['Dados Destinatário'].apply(lambda x: ''.join(filter(str.isdigit, x)))
    result['Dados Emitente'] = result['Dados Emitente'].apply(lambda x: re.sub(r'\d+', '', x))
    result['Dados Destinatário'] = result['Dados Destinatário'].apply(lambda x: re.sub(r'\d+', '', x))
    result['Dados Emitente'] = result['Dados Emitente'].apply(lambda x: re.sub(r'[^\w\s]', '', x))
    result['Dados Destinatário'] = result['Dados Destinatário'].apply(lambda x: re.sub(r'[^\w\s]', '', x))
    result['Imposto'] = result['Imposto'].apply(lambda x: re.sub(r'^R\$', '', x))

    result = pd.DataFrame(result)
    novaOrdemColunas = ['Nº da NF-', 'Data Emissão', 'UF Origem', 'Dados Emitente', 'CNPJ Emitente', 'UF Destino', 'Dados Destinatário', 'CNPJ Destinatário', 'Imposto', 'Situação da cobrança', 'Tipo de Serviço']
    result = result.reindex(columns=novaOrdemColunas)

    print(result)
    result.to_excel('SISAT.xlsx', header=False, index=False)
except ValueError:
    breakpoint()
