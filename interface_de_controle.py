import tkinter as tk
from tkinter import messagebox
import json

def carregarParametros():
    with open("parametros.json", "r") as infile:
        parametros = json.load(infile)
    caminhonavegadorentry.insert(0,parametros["caminhonavegador"])
    destinoLHentry.insert(0,parametros["destinoLH"])
    duracaoAtualizacaoHoraHoraentry.insert(0,parametros["duracaoAtualizacaoHoraHora"])
    delayAcompanhamentoExpedicaoentry.insert(0,parametros["delayAcompanhamentoExpedicao"])
    ID_PLANILHA_BASE_COCKPITentry.insert(0,parametros["ID_PLANILHA_BASE_COCKPIT"])
    ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHHentry.insert(0,parametros["ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH"])
    perfilFirefoxentry.insert(0,parametros["perfilFirefox"])
    return parametros

def gravarParametros():

    try:
        #validação de campos numéricos

        int(duracaoAtualizacaoHoraHoraentry.get())
        int(delayAcompanhamentoExpedicaoentry.get())
        # int(delaypreclickwebentry.get())

        # validação de campo vazio

        if '' in [
            caminhonavegadorentry.get(),
            destinoLHentry.get(),
            duracaoAtualizacaoHoraHoraentry.get(),
            delayAcompanhamentoExpedicaoentry.get(),
            ID_PLANILHA_BASE_COCKPITentry.get(),
            ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHHentry.get(),
            perfilFirefoxentry.get()
            ]:
            messagebox.showinfo(title="Cuidado!", message="Há campos vazios!")
        else:
            parametros = {
                "caminhonavegador": caminhonavegadorentry.get(),
                "destinoLH": destinoLHentry.get(),
                "duracaoAtualizacaoHoraHora": duracaoAtualizacaoHoraHoraentry.get(),
                "delayAcompanhamentoExpedicao": delayAcompanhamentoExpedicaoentry.get(),
                "ID_PLANILHA_BASE_COCKPIT": ID_PLANILHA_BASE_COCKPITentry.get(),
                "ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH": ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHHentry.get(),
                "perfilFirefox": perfilFirefoxentry.get()
            }

            with open("parametros.json", "w") as outfile:
                json.dump(parametros, outfile)
            messagebox.showinfo(title="Feito!", message="Parâmetros gravados com sucesso!")
    except:
        messagebox.showinfo(title="Cuidado!", message="Valores de Delay inválidos")

def agendarPausa():
    while True:
        try:
            with open("pause.json", "r") as infile:
                parametros = json.load(infile)
            break
        except FileNotFoundError:
            parametros = {
            "statuspausa": False,
        }
            with open("pause.json", "w") as outfile:
                json.dump(parametros, outfile)

    if parametros["statuspausa"] == True:
        parametros["statuspausa"] = False
    else:
        parametros["statuspausa"] = True
    parametros = {
        "statuspausa": parametros["statuspausa"],
    }
    with open("pause.json", "w") as outfile:
        json.dump(parametros, outfile)
    messagebox.showinfo(title="Feito!", message="Pausa agendada com sucesso!")

window = tk.Tk()
window.geometry("600x280")
window.title("Parâmetros do robô")
window.resizable(False,False)
window.columnconfigure(1, weight=1)
window.columnconfigure(0, weight=3)

# labels da interface

caminhonavegador = tk.Label(text="Caminho de instalação Firefox/Chrome: ")
caminhonavegador.grid(column=0,row=0,sticky=tk.E)
destinoLH = tk.Label(text="Destino de LH (sua operação): ")
destinoLH.grid(column=0,row=1,sticky=tk.E)
duracaoAtualizacaoHoraHora = tk.Label(text="Duração(m*) de atualização de Etiquetagem e Sorting HH: ")
duracaoAtualizacaoHoraHora.grid(column=0,row=2,sticky=tk.E)
delayAcompanhamentoExpedicao = tk.Label(text="Tempo(m*) de pausa para atualização de Expedição (Sugerido: 10 minutos): ")
delayAcompanhamentoExpedicao.grid(column=0,row=3,sticky=tk.E)
ID_PLANILHA_BASE_COCKPIT = tk.Label(text="ID de planilha geral do Cockpit: ")
ID_PLANILHA_BASE_COCKPIT.grid(column=0,row=4,sticky=tk.E)
ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH = tk.Label(text="ID de planilha de etiquetagem HH: ")
ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHH.grid(column=0,row=5,sticky=tk.E)
perfilFirefox = tk.Label(text="Perfil do Firefox: ")
perfilFirefox.grid(column=0,row=6,sticky=tk.E)
observacoes = tk.Label(text="*Onde: s=Segundos, m=Minutos")
observacoes.grid(column=0,row=9,sticky=tk.E)

# campos de entrada

caminhonavegadorentry = tk.Entry(window)
caminhonavegadorentry.grid(column=1,row=0,sticky=tk.E)
destinoLHentry = tk.Entry()
destinoLHentry.grid(column=1,row=1,sticky=tk.E)
duracaoAtualizacaoHoraHoraentry = tk.Entry()
duracaoAtualizacaoHoraHoraentry.grid(column=1,row=2,sticky=tk.E)
delayAcompanhamentoExpedicaoentry = tk.Entry()
delayAcompanhamentoExpedicaoentry.grid(column=1,row=3,sticky=tk.E)
ID_PLANILHA_BASE_COCKPITentry = tk.Entry()
ID_PLANILHA_BASE_COCKPITentry.grid(column=1,row=4,sticky=tk.E)
ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHHentry = tk.Entry()
ID_PLANILHA_BASE_COCKPIT_ETIQUETAGEMHHentry.grid(column=1,row=5,sticky=tk.E)
perfilFirefoxentry = tk.Entry()
perfilFirefoxentry.grid(column=1,row=6,sticky=tk.E)

# botões de agendamento

atualizarparametros = tk.Button(window,text='1 - Carregar parâmetros', command=carregarParametros)
atualizarparametros.grid(column=0,row=8, columnspan=2,padx=5, pady=5,sticky=tk.E)
atualizarparametros = tk.Button(window,text='2 - Atualizar parâmetros', command=gravarParametros)
atualizarparametros.grid(column=0,row=9, columnspan=2,padx=5, pady=5,sticky=tk.E)

window.mainloop()