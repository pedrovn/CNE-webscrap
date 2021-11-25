#### CNE WEB SCRAPING PARA ESTADOS ####

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from openpyxl import Workbook

### SET DRIVER SELENIUM ####

PATH = "C:/Users/Pedro/Desktop/pyscrapcne/chromedriver_win32/chromedriver.exe"
driver =  webdriver.Chrome(PATH)
driver.get("https://www2.cne.gob.ve/rm2021")


### ESPERAR QUE CARGUE LA PAGINA ANTES DE HACER CUALQUIER COSA #######

try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.card-header.d-flex.justify-content-end.paddingtb0"))
    )
except:
    print("PAGINA NO CARGA MALDITA SEA EL CNE")
finally:
    print("PAGINA CARGADA CORRECTAMENTE")

wb = Workbook()
dest_filename = 'ERM2021-ESTADOS.xlsx'
delsheet = wb['Sheet']
wb.remove(delsheet)

### CUANTOS ESTADOS #####

nestados = driver.find_elements_by_css_selector("div > div:nth-child(1) > select > option")
nest = len(nestados) - 3
print('Se encontraron ' + str(nest) + ' estados\n')

##### BEGIN LOOP DE LOS ESTADOS #### 

for x in range(3,len(nestados)):

    print('Consultando: ' + nestados[x].text + '(' + str(x) + ')')

    #### ESTADOS ####
    nomestado = driver.find_element_by_css_selector("div > div:nth-child(1) > select > option:nth-child(" + str(x+1) + ")")
    print(str(x))
    clickestado = nomestado.click()
    print('Se tocó el estado ' +  nomestado.text)    
    time.sleep(4)

    ### CREAR HOJA PARA ESTADO ####
    nombre_estado_hoja = nestados[x].text
    createws = wb.create_sheet(title=nombre_estado_hoja)
    ws = wb[nombre_estado_hoja]

    ### CLICK FILTRAR ####
    driver.find_element_by_css_selector('div > div.card-body.ng-star-inserted > filter-form > form > div > div:nth-child(6) > button > span').click()
    print ("Se clickeó para filtrar resultados de " + nomestado.text)
    time.sleep(4)

    ######## SCRAP DATOS DE CANDIDATOS ##############

    getdatacand = driver.find_elements_by_css_selector('tr.candidateRow')
    
    ##### LOOP PARA AGREGAR LOS DATOS AL EXCEL #######

    for x in range(len(getdatacand)):
        print(getdatacand[x].text)
        agre = getdatacand[x].text
        ws.append([agre])

    print("Se agregó la data de " + nomestado.text + ' correctamente al Excel' +'\n')

wb.save(filename = dest_filename)
driver.quit()
