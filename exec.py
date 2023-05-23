from testaTicket import ler_indicadores
from comparaCodigo2 import compara_codigo
from testRequest import auto_request
from aceitaDesconto import aceita_auto
from requestTabela import manu_desconto

case = int(input('Informe a função desejada\n1 - Gerar planilha de desconto automático\n2 - Aplicar desconto automático\n3 - Desconto manual\n'))

if case == 1:
    ler_indicadores()
    compara_codigo()
elif case == 2:
    auto_request()
    aceita_auto()
elif case == 3:
    manu_desconto()
    aceita_auto()
