from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from datetime import datetime 
import pandas as pd
import traceback
import json
import time
import os

def carregarParametros():
	with open("parametros.json", "r") as infile:
		parametros = json.load(infile)
	return parametros

def funcaoPrincipal():
	tabela_id_pacote_id_seller = pd.DataFrame({'ID Pacote':[],'ID Seller':[]})
	lista_de_ids = pd.read_clipboard()
	for id in lista_de_ids[lista_de_ids.columns[0]]:
		driver.get(f'https://tms.mercadolivre.com.br/packages/{id}/detail')
		for i in range(10):
			time.sleep(1)
			try:
				id_seller = driver.find_elements(By.CLASS_NAME, 'collapsible-info__value')[2].text
				break
			except:
				pass
		novalinha = pd.DataFrame({'ID Pacote':[id],'ID Seller':[id_seller]})
		tabela_id_pacote_id_seller = pd.concat([tabela_id_pacote_id_seller,novalinha])
	time_agora = time.strftime('%d_%m_%Y %H_%M_%S')
	tabela_id_pacote_id_seller['ID Seller'] = tabela_id_pacote_id_seller['ID Seller'].astype('int32')
	tabela_id_pacote_id_seller.to_excel(f'IDs de Sellers - {time_agora}.xlsx')
	
diretorio_robo = os.getcwd()
user_name = os.getlogin()
debug_mode = False

print('Abrindo driver Firefox')
# profile_path = carregarParametros()["perfilFirefox"]
options = Options()
# options.add_argument("-profile")
# options.add_argument(profile_path)
options.binary_location = carregarParametros()["caminhonavegador"]
driver = webdriver.Firefox(options=options)
driver.get('https://tms.mercadolivre.com.br')

input('Após logar no TMS, pressione ENTER para continuar...\n')

while True:
	try:
		funcaoPrincipal()
		break
	except:
		if debug_mode:
			print(traceback.format_exc())
		pass
