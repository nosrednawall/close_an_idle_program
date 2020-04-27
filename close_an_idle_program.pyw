"""
This program close an application after time stipulated for Windows
"""
# imports apis windows
import win32api
import wmi

# imports de tipos de data e tempo
import time
from datetime import datetime as dt
from datetime import timedelta

# imports de gui windows
import winsound
import PySimpleGUI as sg

# nome e path dos programas
PROGRAMA_UM_NOME = "notepad.exe"
PROGRAMA_UM_PATH = "C:\\Windows\\system32\\notepad.exe"
PROGRAMA_UM_MENSAGEM = "Notepad"

# Tempo Ocioso
min_ocioso = 40
min_entre_checagens = 10

### Main Script here
# Collection time information
t_now = dt.now()                                    # Current time for reference
t_ocioso = min_ocioso*60                            # Tempo ocioso
t_delta_ocioso = timedelta(0,t_ocioso)              # Time Delta in mins

# import windowns management
c = wmi.WMI ()
    
# Inicio das funcoes

# Pega o tempo inativo do computador em segundos
def getIdleTime():
    return (win32api.GetTickCount() - win32api.GetLastInputInfo()) / 1000.0

# funcão principal
def buscaPrograma(nome,path,mensagemPersonalizada):
    # Entra dentro do OctopusLeclair8
    if process.Name == nome and process.ExecutablePath == path:
        #Recebe a data do json de process e remove a parte desnecessária
        data_cru = process.CreationDate.split('.')

        # Hora de agora
        data_agora = dt.now()

        # Hora criado em formato objeto datetime
        data_criado = dt.strptime(data_cru[0], '%Y%m%d%H%M%S')
        
        # Calcula a diferença
        diferenca = timedelta()
        diferenca = data_agora - data_criado

        #print(process.ProcessId, process.Name, process.ExecutablePath)
        #print(diferenca)

        # Verifica se o valor da diferença do horario de inicialização do programa é maior que 40 minutos e também o tempo ocioso do computador é maior que 40 minutos
        if diferenca > t_delta_ocioso and getIdleTime() > t_ocioso:
            winsound.Beep((37*100), 500)
            sg.popup_timed( "Computador Ocioso a mais de "+ repr(min_ocioso) +" minutos\nFechando "+ mensagemPersonalizada +" em 1 minuto" ,auto_close_duration=60)

            if getIdleTime() > t_ocioso:
                resultado = process.Terminate()
                mensagem = "Encerrado " if resultado != 0 else "Ocorreu um erro ao encerrar "
                winsound.Beep((37*100), 500)
                sg.Popup(mensagem+mensagemPersonalizada)
                

# Loop principal
while True:
    
    # Busca os octopus em todos os programas rodando no computador
    for process in c.Win32_Process():
        # Adicione quantos programas desejar
        buscaPrograma(PROGRAMA_UM_NOME,PROGRAMA_UM_PATH,PROGRAMA_UM_MENSAGEM)
# Para não sobrecarregar o CPU
time.sleep(min_entre_checagens * 60)