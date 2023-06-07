from testaTicket import ler_indicadores
from comparaCodigo2 import compara_codigo
from testRequest import auto_request
from aceitaDesconto import aceita_auto
from requestTabela import manu_desconto
import PySimpleGUI as sg
from PySimpleGUI import Column, VSeparator
import re

sg.theme('SystemDefault')

layout_esquerdo = [
    [sg.Image('./assets/logo.png')]
]

layout_direito = [
    [sg.Push(), sg.Text('Informe a função desejada', font=("Helvetica", 10, "bold")), sg.Push()],
    [sg.Radio('Gerar planilha de desconto automático', "RADIO1", default=False, key="-IN1-"), sg.Push()],
    [sg.Radio('Aplicar desconto automático', "RADIO1", default=False, key="-IN2-"), sg.Push()],
    [sg.Radio('Desconto manual', "RADIO1", default=False, key="-IN3-"), sg.Push()],
    [sg.Push(), sg.Button('Proximo', size=(10,1), button_color=('white', '#28478E'),border_width=0)]
]

layout = [
    [
        Column(layout_esquerdo),
        VSeparator(),
        Column(layout_direito)
    ]
]

layout_planilha = [
    [sg.Text('Gerar planilha de desconto automático', font=("Helvetica", 10, "bold"))],
    [sg.Text('Informe o mês da planilha de Indiacadores'), sg.Push(),sg.Input(key='-PLANILHA_MES-')],
    [sg.Text('Digite a data de entrada no formato dia/mes/ano'), sg.Push(), sg.Input(key='-PLANILHA_INICIO-')],
    [sg.Text('Digite a ultima data de entrada no formato dia/mes/ano'), sg.Push(), sg.Input(key='-PLANILHA_FIM-')],
    [sg.Text('Nome para a planilha'), sg.Push(), sg.Input(key='-PLANILHA_NOME-')],
    [sg.Push(), sg.Button('Gerar', size=(20,1), button_color=('white', '#28478E'),border_width=0)]
]

layout_auto = [
    [sg.Text('Aplicar Desconto Automático', font=("Helvetica", 10, "bold"))],
    [sg.Text('Informe o nome da planilha com os descontos'), sg.Push(),sg.Input(key='-AUTO_PLANILHA-')],
    [sg.Push(), sg.Button('Aplicar', size=(20,1), button_color=('white', '#28478E'),border_width=0)]
]

layout_manu = [
    [sg.Text('Aplicar Desconto Manual', font=("Helvetica", 10, "bold"))],
    [sg.Push()],
    [sg.Text('Escolha a empresa da planilha'), sg.Push() ,sg.Radio('PADTEC', "RADIO1", default=True, key="-MANU_PAD-"), sg.Push(),sg.Radio('RADIANTE', "RADIO1", default=False, key="-MANU_RAD-")],
    [sg.Text('Informe a linha que inicia os novos descontos'), sg.Push(),sg.Input(key='-MANU_LINHA-', size=(36, 1))],
    [sg.Push(), sg.Button('Aplicar', size=(20,1), button_color=('white', '#28478E'),border_width=0)]
]

window = sg.Window(
    'Tela de seleção',
    layout=layout,
    element_justification='c',
    icon='./assets/icontelebras_resized.ico'
)

window_planilha = sg.Window(
    'Gerar planilha de desconto automático',
    layout=layout_planilha,
    element_justification='c',
    icon='./assets/icontelebras_resized.ico'
)

window_auto = sg.Window(
    'Aplicar Desconto Automático',
    layout=layout_auto,
    element_justification='c',
    icon='./assets/icontelebras_resized.ico'
)

window_manu = sg.Window(
    'Aplicar Desconto Manual',
    layout=layout_manu,
    element_justification='c',
    icon='./assets/icontelebras_resized.ico'
)

while True:
    evento, valores = window.read()

    if evento == sg.WINDOW_CLOSED:
        break
    elif valores['-IN1-'] == True:
        window.close()
        while True:
            evento_planilha, valores_planilha = window_planilha.read()
        
            if evento_planilha == sg.WINDOW_CLOSED:
                break
            elif valores_planilha['-PLANILHA_MES-'] != '' and valores_planilha['-PLANILHA_INICIO-'] != '' and valores_planilha['-PLANILHA_FIM-'] != '' and valores_planilha['-PLANILHA_NOME-'] != '':
                
                pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
                if pattern.match(valores_planilha['-PLANILHA_INICIO-']) and pattern.match(valores_planilha['-PLANILHA_FIM-']):
                    ler_indicadores(valores_planilha['-PLANILHA_MES-'], valores_planilha['-PLANILHA_INICIO-'], valores_planilha['-PLANILHA_FIM-'])
                    compara_codigo(valores_planilha['-PLANILHA_NOME-'], valores_planilha['-PLANILHA_MES-'])
                    window_planilha.close()
                    break
                else:
                    sg.popup("A data não está no formato dia/mês/ano")
            else:
                sg.popup("Prencha todos os campos")
       
    elif valores['-IN2-'] == True:
        
        window.close()
        while True:
            evento_auto, valores_auto = window_auto.read()
        
            if evento_auto == sg.WINDOW_CLOSED:
                break
            elif valores_auto['-AUTO_PLANILHA-'] != '':
                auto_request(valores_auto['-AUTO_PLANILHA-'])
                aceita_auto()
                window_auto.close()
                break
            else:
                sg.popup("Prencha o campo com a planilha")

    elif valores['-IN3-'] == True:
        
        window.close()
        while True:
            evento_manu, valores_manu = window_manu.read()
            empresa = 0
            if evento_manu == sg.WINDOW_CLOSED:
                break
            elif valores_manu['-MANU_LINHA-'] != '':
                if valores_manu['-MANU_LINHA-'][-1] not in ('0123456789'):
                    sg.popup("Apenas números permitido")
                else:
                    if valores_manu['-MANU_PAD-'] == True:
                        empresa = 1
                    else:
                        empresa = 2
                    manu_desconto(empresa, valores_manu['-MANU_LINHA-'])
                    aceita_auto()
                    window_manu.close()
                    break
            else:
                sg.popup("Prencha o campo linha de inicio")

    window.close()