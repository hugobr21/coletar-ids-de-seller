from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from google_api_functions import *
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

def importarBasesDeRoteirizacao(nome_planilha,ciclo):
	try:
		try:
			planilhaBase = pd.read_excel([nomeDoArquivo for nomeDoArquivo in os.listdir() if nome_planilha in nomeDoArquivo][0],sheet_name='Planilha1')
		except:
			planilhaBase = pd.read_excel([nomeDoArquivo for nomeDoArquivo in os.listdir() if nome_planilha in nomeDoArquivo][0],sheet_name='Plan1')
		planilhaBase['Ciclo'] = ciclo
		planilhaBase['Shipment'] = planilhaBase['Shipment'].astype('str')
		return planilhaBase[['Shipment','Ciclo','Rota']]
	except:
		try:
			planilhaBase = pd.read_excel(nome_planilha + '.xlsm',sheet_name='Planilha1')
			planilhaBase['Ciclo'] = ciclo
			planilhaBase['Shipment'] = planilhaBase['Shipment'].astype('str')
			return planilhaBase[['Shipment','Ciclo','Rota']]
		except:
			return pd.DataFrame({'Shipment':[],'Rota':[],'Ciclo':[]})

def importarPlanification():
	try:
		planification = pd.read_csv([nomeDoArquivo for nomeDoArquivo in os.listdir() if 'planification' in nomeDoArquivo][0])
		planification['Shipment'] = planification['Shipment'].astype('str')
		return planification
	except:
		print('Um erro ocorreu ao importar planification.')

def atualizarBase(id_planilha,aba_range,dados_da_base):
	limpar_celulas(id_planilha,aba_range)
	update_values(id_planilha,aba_range, 'USER_ENTERED', dados_da_base)

def apagarCSVs():
	os.chdir(r'C:\\Users\\'+ user_name +'\\Downloads')
	try:
		nomesDosArquivos = [nomesDosArquivos for nomesDosArquivos in os.listdir() if ('.csv' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)]
		if debug_mode:
			print('Pasta de download limpa.')
		for arquivo in nomesDosArquivos:
			os.remove(arquivo)
		os.chdir(diretorio_robo)
	except IndexError:
		pass
	except Exception as e:
		time.sleep(1)

def baixar_planification():
	driver.get('https://envios.mercadolivre.com.br/logistics/routing/planification/download')
	nome_do_arquivo = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\' + 'planification_' + '_'.join([time.strftime("%d"),time.strftime("%m"),time.strftime("%Y")]) + '.csv'
	# print(nome_do_arquivo)
	contador = 0
	while True:
		time.sleep(5)
		try:
			driver.find_element(By.XPATH,'//*[@id="routing-downloads"]').click()
			botao_baixar_planification = driver.find_element(By.XPATH,'/html/body/main/div/div/div/div/div/div[7]/button')
			botao_baixar_planification.click()
			# print('clicar p baixar')
			break	
		except:
			if contador >= 10:
				print('#1 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/routing/planification/download")
				time.sleep(6)
				# funcao_principal()
			if debug_mode:
				print('erro ao baixar arquivo')
				print(traceback.format_exc())
			contador = contador + 1
			pass
	contador = 0
	while True:
		time.sleep(1)
		try:
			os.chdir(f'C:\\Users\\{user_name}\\Downloads')
			planification = pd.read_csv([nomesDosArquivos for nomesDosArquivos in os.listdir() if ('planification' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)][0], low_memory = False)
			os.chdir(diretorio_robo)
			planification['Shipment'] = pd.to_numeric(planification['Shipment'], errors='coerce')
			planification = planification.loc[~ (planification['Status'].isna())]
			planification = planification.loc[~ (planification['Shipment'].isna())]
			planification['Shipment'] = planification['Shipment'].astype('str').str[:11]
			planification = planification.fillna('')
			os.remove(nome_do_arquivo)
			if debug_mode:
				print('Arquivo carregado')
			return planification
		except Exception as e:
			if debug_mode:
				print(e)
				print(traceback.format_exc())
			contador = contador + 1
			if contador >=100:
				try:
					os.remove(nome_do_arquivo)
				except:
					pass
				print('#2 Reiniciando página e download')
				driver.get("https://envios.mercadolivre.com.br/logistics/routing/planification/download")
				time.sleep(6)
				raise KeyError
			# print('erro ao carregar arquivo')
			pass

def carregarBaseSorteadoEtiquetado(nome_base):
	data_agora = datetime(datetime.now().year,datetime.now().month,datetime.now().day,00,00,00).strftime("_%d_%m_%Y")
	try:
		base_sorteado_etiquetado = pd.read_excel(f'{nome_base}{data_agora}.xlsx')
	except:
		base_sorteado_etiquetado = pd.DataFrame({'Shipment':[],'Mudança de Status':[],'Hora':[]})
		base_sorteado_etiquetado.to_excel(f'{nome_base}{data_agora}.xlsx', index=False)
	base_sorteado_etiquetado['Shipment'] = base_sorteado_etiquetado['Shipment'].astype('str')
	return base_sorteado_etiquetado

def consolidarBaseSorteadoEtiquetado(planification,baseDeRoteirizacao):
	data_agora = datetime(datetime.now().year,datetime.now().month,datetime.now().day,00,00,00)
	planification_roteirizacao = planification.copy()
	colunas_planification = ['Shipment','Status','Facility de Origem','Promessa','Bairro', 'Cidade']
	planification_roteirizacao = planification_roteirizacao[colunas_planification]

	if debug_mode:
		print('Filtrando pacotes goleiro pela data')

	if debug_mode:
		print('Tratando bairro do planification')
	# Tratando bairro
	planification_roteirizacao['Bairro'] = planification_roteirizacao['Bairro'] + ', ' + planification_roteirizacao['Cidade']

	if debug_mode:
		print('Criando coluna de LeadTime')
	# Criando coluna de LeadTime
	planification_roteirizacao['Promessa'] = planification_roteirizacao['Promessa'].str.split('/').str[0].astype('datetime64')
	planification_roteirizacao.loc[planification_roteirizacao['Promessa'] == data_agora, 'Lead Time'] = 'On Time'
	planification_roteirizacao.loc[planification_roteirizacao['Promessa'] > data_agora, 'Lead Time'] = 'Early'
	planification_roteirizacao.loc[planification_roteirizacao['Promessa'] < data_agora, 'Lead Time'] = 'Delay'
	planification_roteirizacao['Shipment'] = planification_roteirizacao['Shipment'].astype('str')
	planification_roteirizacao = planification_roteirizacao.merge(baseDeRoteirizacao.copy(), how='left', on='Shipment')
	planification_roteirizacao['Cluster'] = planification_roteirizacao['Rota'].str.split('_').str[0]
	planification_roteirizacao = planification_roteirizacao.loc[~(pd.to_numeric(planification_roteirizacao['Shipment'],errors='coerce').isna()==True)]
	planification_roteirizacao = planification_roteirizacao.fillna('')
	planification_roteirizacao['Promessa'] = planification_roteirizacao['Promessa'].astype('str')
	planification_roteirizacao = planification_roteirizacao[planification_roteirizacao.columns].values.tolist()
	
	if debug_mode:
		print('Cruzando base de roteirização com planification')

	# Cruzando base de roteirização com planification
	baseDeRoteirizacaoConsolidada_tratada = baseDeRoteirizacao.copy().merge(planification.copy(), how='left', on='Shipment')
	baseDeRoteirizacaoConsolidada_tratada.loc[baseDeRoteirizacaoConsolidada_tratada['Status'].isna(),'Status'] = 'Sorteado'
	baseDeRoteirizacaoConsolidada_tratada.loc[baseDeRoteirizacaoConsolidada_tratada['Status'] == 'at_station|sorting', 'Status'] = 'Etiquetado'
	baseDeRoteirizacaoConsolidada_tratada.loc[baseDeRoteirizacaoConsolidada_tratada['Status'] == 'on_way', 'Status'] = 'On Way'
	baseDeRoteirizacaoConsolidada_tratada['Prefixo'] = baseDeRoteirizacaoConsolidada_tratada['Rota'].str.split('_').str[0]
	
	if debug_mode:
		print('Armazenando pacotes etiquetados')
	
	# Armazenando pacotes etiquetados
	base_etiquetado = carregarBaseSorteadoEtiquetado('mudanca_de_status_etiquetagem')

	planification_atstation = baseDeRoteirizacaoConsolidada_tratada.copy().merge(base_etiquetado, how='left', on='Shipment')
	planification_atstation = planification_atstation.loc[(planification_atstation['Status'] == 'Etiquetado') & (planification_atstation['Mudança de Status'].isna() == True)]
	planification_atstation['Mudança de Status'] = planification_atstation['Status']
	planification_atstation['Hora'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
	planification_atstation = planification_atstation[['Shipment','Mudança de Status', 'Hora']]
	base_etiquetado = pd.concat([base_etiquetado, planification_atstation])
	data_arquivo = data_agora.strftime("_%d_%m_%Y")
	base_etiquetado.to_excel(f'mudanca_de_status_etiquetagem{data_arquivo}.xlsx', index=False)

	if debug_mode:
		print('Armazenando pacotes sorteado')
	
	# Armazenando pacotes sorteado
	pacotes_sorteados = carregarBaseSorteadoEtiquetado('mudanca_de_status_sorteado')
	planification_sorteado = baseDeRoteirizacaoConsolidada_tratada.copy().merge(pacotes_sorteados, how='left', on='Shipment')
	# print(planification_sorteado)

	planification_sorteado = planification_sorteado.loc[(planification_sorteado['Status'] == 'Sorteado') & (planification_sorteado['Mudança de Status'].isna() == True)]
	planification_sorteado['Mudança de Status'] = planification_sorteado['Status']
	planification_sorteado['Hora'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
	planification_sorteado = planification_sorteado[['Shipment','Mudança de Status', 'Hora']]
	pacotes_sorteados = pd.concat([pacotes_sorteados, planification_sorteado])
	data_arquivo = data_agora.strftime("_%d_%m_%Y")
	pacotes_sorteados.to_excel(f'mudanca_de_status_sorteado{data_arquivo}.xlsx', index=False)

	sorteadoEtiquetadoConsolidado = pd.concat([pacotes_sorteados, base_etiquetado])
	sorteadoEtiquetadoConsolidado = sorteadoEtiquetadoConsolidado.merge(baseDeRoteirizacao.copy()[['Shipment', 'Rota', 'Ciclo']], how='left', on='Shipment')
	sorteadoEtiquetadoConsolidado = sorteadoEtiquetadoConsolidado[['Shipment','Mudança de Status', 'Hora', 'Rota','Ciclo']]
	sorteadoEtiquetadoConsolidado = sorteadoEtiquetadoConsolidado.fillna('')
	return sorteadoEtiquetadoConsolidado


def baixarMonitoramentoTerrestre():
	driver.get('https://envios.mercadolivre.com.br/logistics/line-haul/monitoring/routes')
	nome_do_arquivo = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\' + 'rotas.csv'
	# print(nome_do_arquivo)
	contador = 0
	# Baixa o arquivo
	while True:
		time.sleep(5)
		try:
			time.sleep(1)
			driver.find_element(By.XPATH,'//*[@id="button-download-csv"]').click()
			break	
		except:
			if contador >= 10:
				print('#1 Reiniciando página e download - baixarmonitoramentoTerrestre')
				driver.get("https://envios.mercadolivre.com.br/logistics/line-haul/monitoring/routes")
				time.sleep(6)
			print('erro ao baixar arquivo')
			print(traceback.format_exc())
			contador = contador + 1
			pass
	contador = 0
	# Carrega e trata o arquivo
	while True:
		time.sleep(1)
		try:
			os.chdir(f'C:\\Users\\{user_name}\\Downloads')
			nome_do_arquivo = [nomesDosArquivos for nomesDosArquivos in os.listdir() if ('rutas' in nomesDosArquivos or 'rotas' in nomesDosArquivos) and ('.part' not in nomesDosArquivos)][0]
			monitoramentoTerrestre = pd.read_csv(nome_do_arquivo)
			os.remove(nome_do_arquivo)
			os.chdir(diretorio_robo)
			
			# monitoramentoTerrestre = monitoramentoTerrestre.loc[monitoramentoTerrestre['Destino'].str.strip() == 'Srj1 (Rio Do Janeiro)']
			monitoramentoTerrestre = monitoramentoTerrestre.loc[monitoramentoTerrestre['Destino'].str.strip() == carregarParametros()["destinoLH"].strip()]
			limiteInferiorDataMonitoramento = datetime(datetime.now().year,datetime.now().month,datetime.now().day,00,00,00)
			limiteSuperiorDataMonitoramento = datetime(datetime.now().year,datetime.now().month,datetime.now().day,23,59,59)
			monitoramentoTerrestre = monitoramentoTerrestre.loc[
				(pd.to_datetime(monitoramentoTerrestre['Destino ETA']) >= limiteInferiorDataMonitoramento) &
			(pd.to_datetime(monitoramentoTerrestre['Destino ETA']) <= limiteSuperiorDataMonitoramento) &
			((monitoramentoTerrestre['Status'] == 'Em curso') | (monitoramentoTerrestre['Status'] == 'Finalizado'))
			]

			monitoramentoTerrestre = monitoramentoTerrestre.fillna('')
			print(f'Arquivo de rotas carregado - baixarmonitoramentoTerrestre')
			return monitoramentoTerrestre
		except Exception as e:
			if debug_mode:
				print(e)
				print(traceback.format_exc())
			contador = contador + 1
			if contador >=100:
				try:
					apagarCSVs()
				except:
					pass
				print('#2 Reiniciando página e download - baixarmonitoramentoTerrestre')
				driver.get("https://envios.mercadolivre.com.br/logistics/line-haul/monitoring/routes")
				time.sleep(6)
				# raise KeyError
			# print('erro ao carregar arquivo')
			pass

def gerarBaseDeRoteirizacao():
	print('Importando bases de roteirizacao')

	baseAM = importarBasesDeRoteirizacao('BASE AM','AM')
	basePM = importarBasesDeRoteirizacao('BASE PM','PM')

	try:
		baseDeRoteirizacaoConsolidada = pd.concat([baseAM,basePM])
	except:
		try:
			baseDeRoteirizacaoConsolidada = pd.concat([baseAM])
		except:
			baseDeRoteirizacaoConsolidada = pd.concat([basePM])

	return baseDeRoteirizacaoConsolidada

def funcaoPrincipal():
	baseDeRoteirizacao = gerarBaseDeRoteirizacao()
	while True:
		try:
			print('Atualizando hora/hora...')
			starttime = time.time()

			for i in range(100):
				apagarCSVs()
				planification = baixar_planification()
				etiquetagemESortingHoraHora = consolidarBaseSorteadoEtiquetado(planification=planification,baseDeRoteirizacao=baseDeRoteirizacao)
				deltatime = time.time() - starttime
				if debug_mode:
					print(f'{deltatime} segundos')
				if (deltatime/60) >= int(carregarParametros()["duracaoAtualizacaoHoraHora"]):
					break
			
			monitoramentoTerrestre = baixarMonitoramentoTerrestre()

			# etiquetagemFormsAM, etiquetagemFormsPM = importarEtiquetagemForms()

			print('Subindo bases para google sheets...')

			ID_PLANILHA_BASE_COCKPIT = carregarParametros()["ID_PLANILHA_BASE_COCKPIT"]
			ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH = carregarParametros()["ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH"]
			
			# ID_PLANILHA_BASE_COCKPIT = '1x3t-0JsNwN38FajdWNWlN9Z_cEbjz-BQqnchy-KjmWQ'
			# ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH = '1PG_xZsWDPJjjYHRDkxuycBiRIzLwohx7BRl1Ca006A0'

			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'PLANIFICATION VIVO!A2:AE')
			update_values(ID_PLANILHA_BASE_COCKPIT,'PLANIFICATION VIVO!A2','USER_ENTERED',planification.values.tolist())
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'MON. TERRESTE!A2:AS')
			update_values(ID_PLANILHA_BASE_COCKPIT,'MON. TERRESTE!A2','USER_ENTERED',monitoramentoTerrestre.values.tolist())
			
			limpar_celulas(ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH,'BASE AM!A2:E')
			update_values(ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH,'BASE AM!A2','USER_ENTERED',etiquetagemESortingHoraHora.values.tolist())

			# limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'INFORMAÇÕES OP!AA3:AC')
			# update_values(ID_PLANILHA_BASE_COCKPIT,'INFORMAÇÕES OP!AA3','USER_ENTERED',etiquetagemFormsAM.values.tolist())
			
			# limpar_celulas(ID_PLANILHA_BASE_COCKPIT,'INFORMAÇÕES OP!AD3:AF')
			# update_values(ID_PLANILHA_BASE_COCKPIT,'INFORMAÇÕES OP!AD3','USER_ENTERED',etiquetagemFormsPM.values.tolist())
			
			update_values(ID_PLANILHA_BASE_COCKPIT,'PLANIFICATION VIVO!AF2','USER_ENTERED',[[time.strftime("%d/%m/%Y %H:%M:%S")]])

		except Exception as e:
			if debug_mode:
				print(e)
				print(traceback.format_exc())
			pass

def importarEtiquetagemForms():
    try:
        inicioDoAM = get_values('1x3t-0JsNwN38FajdWNWlN9Z_cEbjz-BQqnchy-KjmWQ','INFORMAÇÕES OP!M3')[0][0]
        inicioDoPM = get_values('1x3t-0JsNwN38FajdWNWlN9Z_cEbjz-BQqnchy-KjmWQ','INFORMAÇÕES OP!U3')[0][0]
        tabelaHHDeReferenciaAM = pd.DataFrame({'Range de horas':pd.Series(pd.date_range(inicioDoAM,periods=8, freq="h")).dt.strftime('%H:%M'),\
			'Hora de Processamento':pd.Series(pd.date_range(inicioDoAM,periods=8, freq="h")).index})
        tabelaHHDeReferenciaPM = pd.DataFrame({'Range de horas':pd.Series(pd.date_range(inicioDoPM,periods=8, freq="h")).dt.strftime('%H:%M'),\
			'Hora de Processamento':pd.Series(pd.date_range(inicioDoAM,periods=8, freq="h")).index})
        dadosForms = get_values('15GKJ_Xa4m6J6bb7a59OnmUKgfRHB3cDmEFMN6PvBOmE','Respostas ao formulário 1!A1:F')
        #print(dadosForms)
        tabelaEtiquetagemForms = pd.DataFrame(dadosForms[1:],columns=dadosForms[0])
        tabelaEtiquetagemForms['Hora de Processamento'] = tabelaEtiquetagemForms['Hora de Processamento'].astype('int8')
        #print(tabelaEtiquetagemForms)
        tabelaEtiquetagemForms = tabelaEtiquetagemForms.sort_values(by='Carimbo de data/hora', ascending=False).drop_duplicates(subset=['Ciclo','Hora de Processamento','Bancada','Data'])
        tabelaEtiquetagemForms = tabelaEtiquetagemForms.sort_values(by=['Hora de Processamento','Bancada'], ascending=True)

        #print(tabelaEtiquetagemForms)
        colunasEtiquetagemForms = ['Range de horas','Bancada','Volume Etiquetado Por Estação']
        etiquetagemFormsAM = tabelaEtiquetagemForms.loc[(tabelaEtiquetagemForms['Ciclo'] == 'AM') & \
			(pd.to_datetime(tabelaEtiquetagemForms['Carimbo de data/hora']) >= datetime(datetime.now().year,datetime.now().month,datetime.now().day,00,00,00))].copy()
        etiquetagemFormsPM = tabelaEtiquetagemForms.loc[(tabelaEtiquetagemForms['Ciclo'] == 'PM') & \
			(pd.to_datetime(tabelaEtiquetagemForms['Carimbo de data/hora']) >= datetime(datetime.now().year,datetime.now().month,datetime.now().day,00,00,00))].copy()
        etiquetagemFormsAM = etiquetagemFormsAM.merge(tabelaHHDeReferenciaAM.copy(), how='left', on='Hora de Processamento')[colunasEtiquetagemForms]
        etiquetagemFormsPM = etiquetagemFormsPM.merge(tabelaHHDeReferenciaPM.copy(), how='left', on='Hora de Processamento')[colunasEtiquetagemForms]
        return etiquetagemFormsAM,etiquetagemFormsPM
    except:
        if debug_mode:
            print(traceback.format_exc())

def verificarProgresso():
	driver.find_element(By.CLASS_NAME,'downloadProgress').get_attribute('value')

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
driver.get('https://envios.mercadolivre.com.br/logistics/routing/planification/download')

input('Após logar no logistics, pressione ENTER para continuar...\n')

while True:
	try:
		funcaoPrincipal()
	except:
		pass
