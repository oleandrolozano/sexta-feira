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
    if not os.path.exists(pastaLog):
        os.makedirs(pastaLog)
    if not os.path.exists(pastaCodigos):
        os.makedirs(pastaCodigos)
    if not os.path.exists(pastaTemp):
        os.makedirs(pastaTemp)

def criarArquivoConfiguracao():
    nomeArquivo = os.path.expanduser('~') + "\\Documents\\rotina\\config.ini"
    Config = configparser.ConfigParser()
    if not os.path.exists(nomeArquivo):
        cfgfile = open(nomeArquivo,'w')
        Config.write(cfgfile)
        cfgfile.close()

def criarArquivoModeloVBS():
    modeloScript = os.path.expanduser('~') + "\\Documents\\rotina\\temp\\modeloScritp.vbs"
    if not os.path.exists(modeloScript):
        url = "https://raw.githubusercontent.com/condkai/sexta-feira/master/modeloScritp.vbs"
        r = requests.get(url)
        f = open(modeloScript, "w")
        f.write(r.text)
        f.close()

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
            
            if data_execucao != ""  and local_arquivo != "" and execucao_habilitada != "" and recorrencia != "" and tempo != "":
                if isinstance(data_execucao, datetime):
                    agora = datetime.now()
                    if agora > data_execucao and execucao_habilitada == "S":                                          
                        executarProgramaSAS(local_arquivo, each_section)
                        atualizarProximaExecucao(config, arquivoConfig,each_section,data_execucao,recorrencia, tempo)     
        
 
def atualizarProximaExecucao(config, localArquivo_ini, secao_ini, data_atual, tipo_recorrencia, tempo):
       
    proxima_execucao = data_atual
    proxima_execucao_str = "1"
    if tipo_recorrencia == "dia":
        proxima_execucao = data_atual + timedelta(days=int(tempo))
        proxima_execucao_str = proxima_execucao.strftime(r"%d/%m/%Y, %H:%M:%S")
    if tipo_recorrencia == "minuto":
        proxima_execucao = data_atual + timedelta(minutes=int(tempo))
        proxima_execucao_str = proxima_execucao.strftime(r"%d/%m/%Y, %H:%M:%S")
    
    print(colored("Rotina: [" + secao_ini +  "] atualizada para: '" + proxima_execucao_str + "'", 'cyan'))
    #print("Rotina: [" + secao_ini +  "] atualizado para: '" + proxima_execucao_str + "'")
    
    with open(localArquivo_ini, 'w') as configfile:        
        config.set(secao_ini, "proxima_execucao", proxima_execucao_str)
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
    Config.write(cfgfile)
    cfgfile.close()
    print(colored("Rotina: [" + nomePrograma +  "] cadastrado com sucesso...", 'cyan'))
    time.sleep(3)
    
    

def lerArquivoConfig():
    nomeArquivo = os.path.expanduser('~') + "\\Documents\\rotina\\config.ini"
    settings = configparser.ConfigParser()
    settings._interpolation = configparser.ExtendedInterpolation()
    settings.read(nomeArquivo)
    settings.sections()
    age = settings.get('p_emissao', 'local_arquivo')
    print(age)

def existeErroArquivoLog(localArquivo):
    file = open(localArquivo)
    print(file.read())
    search_word = input("ERROR:")
    if(search_word == file):
        return True
    else:
        return False

def isTimeFormat(input):
    try:
        time.strptime(input, '%H:%M:%S')
        return True
    except ValueError:
        return False

def main():
    os.system("cls") # Windows
    print("Sexta-Feira: Em execução...")
    choice ='0'
    while choice =='0':
        print("[1] Executar Rotinas Automáticas")
        print("[2] Cadastrar uma nova rotina")
        print("[9] Sair")
        
        choice = input ("Escolha uma opção: ")

        if choice == "1":
            #print("Executando rotinas...Pressione 9 para sair ao Menu...")
            print(colored("Executando rotinas...Pressione [9] para sair...", 'blue'))
            try:                
                
                while True:
                    criarPastaRotina() 
                    criarArquivoConfiguracao()
                    #lerArquivoConfig()
                    criarArquivoModeloVBS()
                    #executarProgramaSAS(r"C:\Users\llozano\Desktop\asd.txt")
                    #criarArquivoConfiguracaoTeste()
                    executarProgramasCadastrados()
                    #shell.Run(r"cmd /K C:\Windows\SysWOW64\cscript.exe C:\Users\llozano\Desktop\vb.vbs")

                    time.sleep(1)
            except KeyboardInterrupt:
                    print("Adeus!. By Sexta.")
            
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





main()

