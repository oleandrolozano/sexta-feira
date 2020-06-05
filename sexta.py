import os
import configparser
import sys, win32com.client
import requests
import time
from datetime import datetime, timedelta
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
from colorama import Fore
from colorama import Style

from colorama import init
from termcolor import colored

shell = win32com.client.Dispatch("WScript.Shell")
# use Colorama to make Termcolor work on Windows too
init()


def criarPastaRotina():
    pastaRotina = os.path.expanduser('~') + "\\Documents\\rotina"
    pastaLog = pastaRotina + "\\logs"
    pastaCodigos = pastaRotina + "\\codigos"
    pastaTemp = pastaRotina + "\\temp"
    if not os.path.exists(pastaRotina):
        os.makedirs(pastaRotina)
        print(colored("Sexta-Feira: [" + pastaRotina + "] criado com sucesso...", 'cyan'))
    if not os.path.exists(pastaLog):
        os.makedirs(pastaLog)
        print(colored("Sexta-Feira: [" + pastaLog + "] criado com sucesso...", 'cyan'))
    if not os.path.exists(pastaCodigos):
        os.makedirs(pastaCodigos)
        print(colored("Sexta-Feira: [" + pastaCodigos + "] criado com sucesso...", 'cyan'))
    if not os.path.exists(pastaTemp):
        os.makedirs(pastaTemp)
        print(colored("Sexta-Feira: [" + pastaTemp + "] criado com sucesso...", 'cyan'))

def criarArquivoConfiguracao():
    nomeArquivo = os.path.expanduser('~') + "\\Documents\\rotina\\config.ini"
    Config = configparser.ConfigParser()
    if not os.path.exists(nomeArquivo):
        cfgfile = open(nomeArquivo,'w')
        Config.write(cfgfile)
        cfgfile.close()
        print(colored("Sexta-Feira: [" + nomeArquivo + "] criado com sucesso...", 'cyan'))

def criarArquivoModeloVBS():
    modeloScript = os.path.expanduser('~') + "\\Documents\\rotina\\temp\\modeloScritp.vbs"
    if not os.path.exists(modeloScript):
        url = "https://raw.githubusercontent.com/condkai/sexta-feira/master/modeloScritp.vbs"
        r = requests.get(url)
        f = open(modeloScript, "w")
        f.write(r.text)
        f.close()
        print(colored("Sexta-Feira: [" + modeloScript + "] criado com sucesso...", 'cyan'))

def executarProgramaSAS(localProgramaSAS, nomeDoArquivo):
    if os.path.exists(localProgramaSAS):
        print(colored("Programa: [" + localProgramaSAS + "] em execução...", 'green'))
        #print("Programa: [" + localProgramaSAS + "] em execução...")   
        localArquivosLog = os.path.expanduser('~') + "\\Documents\\rotina\\logs\\" + nomeDoArquivo + ".log"
        modeloScript = os.path.expanduser('~') + "\\Documents\\rotina\\temp\\modeloScritp.vbs"
        arquivoTemporarioDeExecucao =  os.path.expanduser('~') + "\\Documents\\rotina\\temp\\tempVB.vbs"


        try:
            os.remove(arquivoTemporarioDeExecucao)
        except OSError as e:  ## if failed, report it back to the user ##
            print ("Error: %s - %s." % (e.filename, e.strerror))

        # Read in the file
        with open(modeloScript, 'r') as file :
            filedata = file.read()

        # Replace the target string
            filedata = filedata.replace("[local_arquivo_sas]", localProgramaSAS)
            filedata = filedata.replace("[local_log_sas]", localArquivosLog)

        # Write the file out again
        with open(arquivoTemporarioDeExecucao, 'w') as file:
            file.write(filedata)

        shell.Run(r"cmd /K C:\Windows\SysWOW64\cscript.exe " + arquivoTemporarioDeExecucao )
    else:
        # Caso o arquivo sas nao exista
        print(colored("Programa: [" + localProgramaSAS + "] não localizado...", 'red'))
        #print("Programa: [" + localProgramaSAS + "] não localizado...")   
        


def executarProgramasCadastrados():
    
    arquivoConfig = os.path.expanduser('~') + "\\Documents\\rotina\\config.ini"
    config = configparser.ConfigParser()
    config.read(arquivoConfig)

    for each_section in config.sections():
        recorrencia = ""
        execucao_habilitada = ""
        local_arquivo = ""
        tempo = ""
        data_execucao = ""
        status_execucao = ""
        data_inicio_execucao = ""
        for (each_key, each_val) in config.items(each_section):
            if each_key == "proxima_execucao":
                data_execucao = datetime.strptime(each_val, r"%d/%m/%Y, %H:%M:%S")
            if each_key == "local_arquivo":
                local_arquivo = each_val
            if each_key == "execucao_habilitada":
                execucao_habilitada = each_val 
            if each_key == "recorrencia":
                recorrencia = each_val
            if each_key == "tempo":
                tempo = each_val 
            if each_key == "status_execucao":
                status_execucao = each_val 
            if each_key == "data_inicio_execucao":
                data_inicio_execucao = datetime.strptime(each_val, r"%d/%m/%Y, %H:%M:%S")

                
            
            if status_execucao == "0":#Dormindo 
                if data_execucao != ""  and local_arquivo != "" and execucao_habilitada != "" and recorrencia != "" and tempo != "" and data_inicio_execucao != "":
                    if isinstance(data_execucao, datetime):
                        agora = datetime.now()
                        if agora > data_execucao and execucao_habilitada == "S":                                          
                            executarProgramaSAS(local_arquivo, each_section)
                            atualizarProximaExecucao(config, arquivoConfig,each_section,data_execucao,recorrencia, tempo)   

            elif status_execucao == "1":#Em Execução
                if data_inicio_execucao != "" and local_arquivo != "":
                    atualizarStatusExecucao(config, arquivoConfig,each_section,data_inicio_execucao, local_arquivo)
 
 
def atualizarStatusExecucao(config, localArquivo_ini, secao_ini, data_inicio_execucao, local_arquivo):
    
    pastaArquivoLog = os.path.expanduser('~') + "\\Documents\\rotina\\logs\\"
    localArquivoLog = pastaArquivoLog + secao_ini + ".log"

    if os.path.exists(localArquivoLog) and isinstance(data_inicio_execucao, datetime):
        if modification_date(localArquivoLog) > data_inicio_execucao:
            with open(localArquivo_ini, 'w') as configfile:  
                
                if existeErro(localArquivoLog):
                    now = datetime.now()
                    print(colored("Programa: [" + secao_ini + "] finalizado com erro...Executando novamente...", 'red'))
                    executarProgramaSAS(local_arquivo, secao_ini)
                    config.set(secao_ini, "data_inicio_execucao", now.strftime(r"%d/%m/%Y, %H:%M:%S"))
                    config.set(secao_ini, "status_execucao", "1")
                else:
                    print(colored("Programa: [" + secao_ini + "] finalizado com sucesso...", 'green'))
                    config.set(secao_ini, "status_execucao", "0")
                config.write(configfile)

def atualizarProximaExecucao(config, localArquivo_ini, secao_ini, data_atual, tipo_recorrencia, tempo):
    now = datetime.now()
    proxima_execucao = data_atual
    proxima_execucao_str = "1"
    if tipo_recorrencia == "dia":
        proxima_execucao = data_atual + timedelta(days=int(tempo))
        proxima_execucao_str = proxima_execucao.strftime(r"%d/%m/%Y, %H:%M:%S")
    if tipo_recorrencia == "minuto":
        proxima_execucao = now + timedelta(minutes=int(tempo))
        proxima_execucao_str = proxima_execucao.strftime(r"%d/%m/%Y, %H:%M:%S")
    
    print(colored("Rotina: [" + secao_ini +  "] atualizada para: '" + proxima_execucao_str + "'", 'cyan'))
    
    with open(localArquivo_ini, 'w') as configfile:        
        config.set(secao_ini, "proxima_execucao", proxima_execucao_str)
        config.set(secao_ini, "data_inicio_execucao", now.strftime(r"%d/%m/%Y, %H:%M:%S"))
        config.set(secao_ini, "status_execucao", "1")
        config.write(configfile)
        


def criarArquivoConfiguracaoTeste():
    nomeArquivo = os.path.expanduser('~') + "\\Documents\\rotina\\config.ini"
    Config = configparser.ConfigParser()
    cfgfile = open(nomeArquivo,'a')
    Config.add_section('p_emissao')
    Config.set('p_emissao','execucao_habilitada', 'S')
    Config.set('p_emissao','local_arquivo',r'C:\Users\llozano\Desktop\asd.txt')
    Config.set('p_emissao','proxima_execucao',r'30/05/2020, 12:00:00')        
    Config.set('p_emissao','recorrencia',"dia")        
    Config.set('p_emissao','tempo','10')        
    Config.write(cfgfile)
    cfgfile.close()

def cadastrarRotina(nomePrograma, localArquivo, horario_execucao, recorrencia, tempo):
    now = datetime.now()
    dt_string = now.strftime(r"%d/%m/%Y, ")
    nomeArquivo = os.path.expanduser('~') + "\\Documents\\rotina\\config.ini"
    Config = configparser.ConfigParser()
    cfgfile = open(nomeArquivo,'a')
    Config.add_section(nomePrograma)
    Config.set(nomePrograma,'execucao_habilitada', 'S')
    Config.set(nomePrograma,'local_arquivo',localArquivo)
    Config.set(nomePrograma,'proxima_execucao',dt_string + horario_execucao)        
    Config.set(nomePrograma,'recorrencia',recorrencia)        
    Config.set(nomePrograma,'tempo',tempo)
    Config.set(nomePrograma,'data_inicio_execucao',dt_string + horario_execucao)
    Config.set(nomePrograma,'status_execucao',"0")
    Config.write(cfgfile)
    cfgfile.close()
    print(colored("Rotina: [" + nomePrograma +  "] cadastrado com sucesso...", 'cyan'))
    time.sleep(3)
    
    



def isTimeFormat(input):
    try:
        time.strptime(input, '%H:%M:%S')
        return True
    except ValueError:
        return False

def main():
    os.system("cls") # Windows
    criarPastaRotina() 
    criarArquivoConfiguracao()
    criarArquivoModeloVBS()
    
    print("Sexta-Feira: Em execução...")
    choice ='0'
    while choice =='0':
        print("[1] Executar Rotinas Automáticas")
        print("[2] Cadastrar uma nova rotina")
        print("[9] Sair")
        
        choice = input ("Escolha uma opção: ")

        if choice == "1":
            #print("Executando rotinas...Pressione 9 para sair ao Menu...")
            print(colored("Executando rotinas...Pressione [CTRL + C] no terminal para sair...", 'blue'))
            try:                
                
                while True:
                    criarPastaRotina() 
                    criarArquivoConfiguracao()
                    criarArquivoModeloVBS()
                    executarProgramasCadastrados()
                    time.sleep(3)
            except KeyboardInterrupt:
                    print(colored("Adeus!. By Sexta.", 'yellow'))
            
        elif choice == "2":
            menu_criar_programa()
        elif choice == "9":
            break
        else:
            print("Opção inválida...")
            main()

def menu_criar_programa():
    os.system("cls") # Windows
    print("Casdastro de nova rotina")
    choice ='0'
    while choice =='0':
        choice = input ("Digite 9 para voltar ou Enter para continuar... ")
        if choice == "9":
            main()
            break

        local_arquivo = input ("Insira o nome do arquivo *.sas:")

        if not os.path.exists(local_arquivo):
            print(colored("Local do arquivo inválido...", 'red'))
            main()
            break

        nome_programa = input ("De um nome ao programa: ")

        horario_execucao = input ("Hora da execução: (05:00:00) por exemplo ")
        if not isTimeFormat(horario_execucao):
            print(colored("Horário invalido...", 'red'))
            main()
            break
        tempo = ""
        recorrencia = input ("[1] a cada X dias ou [2] a cada X minutos: (Escolha uma opção) ")
        if recorrencia == "1":
            tempo = input ("De quantos em quantos dias: ")
            if not tempo.isdigit():
                print(colored("Dias invalidos...", 'red'))
                main()
                break
            recorrencia = "dia"
        elif recorrencia == "2":
            tempo = input ("De quantos em quantos minutos: ")
            if not tempo.isdigit():
                print(colored("Numeros invalidos...", 'red'))
                main()
                break
            recorrencia = "minuto"
        else:
            print(colored("Recorrencia invalida...", 'red'))
            main()
            break

        cadastrarRotina(nome_programa, local_arquivo, horario_execucao, recorrencia, tempo)
        main()
        break

def menu_sair():
    print("Saindo...")


def modification_date(filename):
    t = os.path.getmtime(filename)
    horarioModificacao = datetime.fromtimestamp(t)
    return horarioModificacao

def existeErro(filename):
    f = open(filename, "r")
    words = ""
    for line in f:    
        words = line.split()
        if len(words) > 0:
            if words[0] == 'ERROR:':
                return True
                

def listarArquivosPorExtencao(folderDir, extencao):
    arquivos = []
    for file in os.listdir(folderDir):
        if file.endswith("." + extencao):
            arquivos.append(os.path.join(r"C:\Users\llozano\Documents", file))

    return arquivos



main()
#existeErro(r"C:\Users\llozano\Documents\p_especiais_721.log")





