from ldap3 import Connection, Server, ALL, SUBTREE
import socket
import win32print
import os
import ctypes



################################################# CONFIGURAÇÃO ##################################################
printer_mapping = {
    'TIC': 'IMP_TIC',   #Tecnologia da Informação
    'FAT': 'IMP_FAT',   #Faturamento
    'CEN': 'IMP_CC',    #Centro Cirurgico
    'SEC': 'IMP_SEC',   #Secretaria
    'COM': 'IMP_COMP',  #Compras
    'AUD': 'IMP_AUD',   #Auditoria
    'CCIH': 'IMP_CCIH', #CCIH
    'UTI': 'IMP_UTI',   #UTI
    'PRO0069': 'IMP_UE',  #PS - PC Medicos
    'UEE': 'IMP_UE',   #PS - PC Enfermagem
    'CON': 'IMP_CONT',  #Contabilidade
    'CST': 'IMP_CUSTOS',#Custos
    'FIN': 'IMP_FIN',   #Financeiro
    'REC0121': 'IMP_INTER',  #PC - Controle de Acesso
    'REC0021': 'IMP_INTER',  #PC - Recepcao Internamento
    'REC0066': 'IMP_INTER',  #PC - Recepcao Internamento 2
    'REC0025': 'IMP_REC_UE',  #PC Meio - Recepção PS
    'REC0027': 'IMP_REC_UE',  #PC Direita - Recepcão PS
    'REC0042': 'IMP_REC_UE',  #PC Esquerda - Recepcão PS
    'DEP': 'IMP_DP',    #Departamento Pessoal
    'WIN10-TREINO1': 'IMP_TIC_COLOR',  #TESTE
}

AD_SERVER = os.getenv('AD_SERVER')
AD_USER = os.getenv('AD_USER')
AD_PASSWORD = os.getenv('AD_PASSWORD')
AD_SEARCH_BASE= os.getenv('AD_SEARCH_BASE')
################################################# CONFIGURAÇÃO ##################################################



################################################# FUNÇÕES ##################################################
# Função para garantir que todas as variáveis de ambiente necessárias estão definidas
def get_env_variable():
    required_vars = ['AD_SERVER', 'AD_USER', 'AD_PASSWORD', 'AD_SEARCH_BASE']
    for value in required_vars:
        if value is None:
            raise EnvironmentError(f'Variável de ambiente ausente: {value}')
    return True

# Função para obter informações do Active Directory
def get_ad_info(username):
    server = Server(AD_SERVER)
    conn = Connection(server, AD_USER, AD_PASSWORD) 
    conn.bind()
    
    searchFilter = f'(sAMAccountName={username})'
    conn.search(search_base=AD_SEARCH_BASE, search_filter=searchFilter, search_scope=SUBTREE, attributes=['cn', 'memberOf'])
    
    user_info = {}
    for entry in conn.entries:
        if entry.cn == username:
            user_info['username'] = entry.cn
            user_info['groups'] = [group for group in entry.memberOf if group.startswith('CN=')]
            break
    
    conn.unbind()
    return user_info

# Função para obter o hostname do PC
def get_hostname():
    return socket.gethostname()

# Função para listar impressoras instaladas
def list_printers():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS)
    #itera sobre cada tupla na lista printers e extrai o terceiro elemento (printer[2]), que é o nome da impressora
    return [printer[2] for printer in printers]

# Função para definir a impressora padrão
def set_default_printer(printer_name):
    try:
        win32print.SetDefaultPrinter(printer_name)
        print(f'Impressora padrão definida para: {printer_name}')
    except Exception as e:
        print(f'Erro ao definir a impressora padrão: {e}')

################################################# FUNÇÕES ##################################################



# Função principal
def main():
    hasEnvNames = get_env_variable()
    
    if hasEnvNames:
        hostname = get_hostname()
        print(f'Hostname: {hostname}')
        
        username = os.getenv('USERNAME')
        print(f'Username: {username}')
        
        printers = list_printers()
        print("Impressoras disponíveis:")
        for printer in printers:
            print(f'- {printer}')
        
        # Lógica para definir impressora padrão baseada no hostname
        default_printer = None
        for prefixHost, printerKey in printer_mapping.items():
            if hostname.startswith(prefixHost):
                for printer in printers:
                    if printerKey in printer:
                        default_printer = printer
                        break
                break

        if default_printer:
            set_default_printer(default_printer)

if __name__ == '__main__':
    main()
