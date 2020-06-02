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

shell = win32com.client.Dispatch("WScript.Shell")

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


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
        #Config.add_section('SEXTA_FEIRA')
        #Config.set('p_emissao','execucao_habilitada', 'S')
        #Config.set('p_emissao','local_arquivo',r'C:\Users\llozano\Desktop\asd.txt')
        #Config.set('p_emissao','proxima_execucao',r'30/05/2020 12:00:00')
        
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

def executarProgramaSAS(localProgramaSAS):
    print("Programa: [" + localProgramaSAS + "] em execução...")   
    localArquivosLog = os.path.expanduser('~') + "\\Documents\\rotina\\logs"
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
                        executarProgramaSAS(local_arquivo)
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
    

    print("Programa: [" + secao_ini +  "] atualizado para: '" + proxima_execucao_str + "'")
    
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


def lerArquivoConfig():
    nomeArquivo = os.path.expanduser('~') + "\\Documents\\rotina\\config.ini"
    settings = configparser.ConfigParser()
    settings._interpolation = configparser.ExtendedInterpolation()
    settings.read(nomeArquivo)
    settings.sections()
    age = settings.get('p_emissao', 'local_arquivo')
    print(age)

def on_created(event):
    print(f"hey, {event.src_path} has been created!")


def on_deleted(event):
    print(f"what the f**k! Someone deleted {event.src_path}!")

def on_modified(event):
    print(f"hey buddy, {event.src_path} has been modified")

def on_moved(event):
    print(f"ok ok ok, someone moved {event.src_path} to {event.dest_path}")

def existeErroArquivoLog(localArquivo):
    file = open(localArquivo)
    print(file.read())
    search_word = input("ERROR:")
    if(search_word == file):
        return True
    else:
        return False

if __name__ == "__main__":
    patterns = "*"
    ignore_patterns = ""
    ignore_directories = False
    case_sensitive = True
    my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)

my_event_handler.on_created = on_created
my_event_handler.on_deleted = on_deleted
my_event_handler.on_modified = on_modified
my_event_handler.on_moved = on_moved


path = "."
go_recursively = True
my_observer = Observer()
my_observer.schedule(my_event_handler, path, recursive=go_recursively)
my_observer.start()

'''
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    my_observer.stop()
    my_observer.join()
'''



try:
    
    print("Sexta-Feira: Em execução...")
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
        my_observer.stop()
        my_observer.join()
        print("Adeus!. By Sexta.")

