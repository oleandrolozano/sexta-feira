import os
import configparser
import sys, win32com.client
import requests

shell = win32com.client.Dispatch("WScript.Shell")

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
        Config.add_section('p_emissao')
        Config.set('p_emissao','execucao_habilitada', 'S')
        Config.set('p_emissao','local_arquivo',r'C:\Users\llozano\Desktop\asd.txt')
        Config.set('p_emissao','proxima_execucao',r'30/05/2020 12:00:00')
        
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

def criarArquivoModeloVBS():
    modeloScript = os.path.expanduser('~') + "\\Documents\\rotina\\temp\\modeloScritp.vbs"
    if not os.path.exists(modeloScript):
        url = "https://raw.githubusercontent.com/condkai/sexta-feira/master/modeloScritp.vbs"
        r = requests.get(url)
        f = open(modeloScript, "w")
        f.write(r.text)
        f.close()

def executarProgramaSAS(localProgramaSAS):
    modeloScript = os.path.expanduser('~') + "\\Documents\\rotina\\temp\\modeloScritp.vbs"
    arquivoTemporarioDeExecucao =  os.path.expanduser('~') + "\\Documents\\rotina\\tempVB.vbs"

    try:
        os.remove(arquivoTemporarioDeExecucao)
    except OSError as e:  ## if failed, report it back to the user ##
        print ("Error: %s - %s." % (e.filename, e.strerror))

    # Read in the file
    with open(modeloScript, 'r') as file :
        filedata = file.read()

    # Replace the target string
        filedata = filedata.replace("[local_arquivo_sas]", localProgramaSAS)

    # Write the file out again
    with open(arquivoTemporarioDeExecucao, 'w') as file:
        file.write(filedata)

    shell.Run(r"cmd /K C:\Windows\SysWOW64\cscript.exe " + arquivoTemporarioDeExecucao )

#def executarProgramasCadastrados():


criarPastaRotina()
criarArquivoConfiguracao()
lerArquivoConfig()
criarArquivoModeloVBS()
executarProgramaSAS(r"C:\Users\llozano\Desktop\asd.txt")
#shell.Run(r"cmd /K C:\Windows\SysWOW64\cscript.exe C:\Users\llozano\Desktop\vb.vbs")