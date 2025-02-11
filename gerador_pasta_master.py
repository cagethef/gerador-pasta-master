import os
import pandas as pd
import re
import xlrd
import winshell

def criar_atalho(caminho, codigo):
    codigo = str(codigo).strip("A")
    for keys, values in pasta_tipos.items():
        if str(codigo).startswith(keys):
            link_filepath = os.path.join(
                diretorio_master, pasta_a_ser_criada, values, codigo + ".lnk")
            break
    with winshell.shortcut(link_filepath) as link:
        link.path = os.path.join(caminho, codigo)
    return

def criar_atalho_teste(caminho, codigo):
    link_filepath = os.path.join(
        diretorio_master, pasta_a_ser_criada, "TESTE", codigo + ".lnk")
    with winshell.shortcut(link_filepath) as link:
        link.path = os.path.join(caminho, pasta_a_ser_criada)
    return

def criar_atalho_smd(caminho, codigo):
    link_filepath = os.path.join(
        diretorio_master, pasta_a_ser_criada, "PROGRAMA SMD", codigo + ".lnk")
    with winshell.shortcut(link_filepath) as link:
        link.path = os.path.join(caminho, pasta_a_ser_criada)
    return

def mapeamento(tipo, letra):
    print(f"o servidor {tipo} não está mapeado na letra {letra} !")
    print("CERTIFIQUE-SE DE QUE OS MAPEAMENTOS DO SERVIDOR ESTÃO CONFORME ABAIXO: ")
    print("desenv -> H:\ninstrucoes -> S:\nate300 -> F:")
    a = input("Aperte enter para continuar...")
    return

diretorio_master = r"S:\PASTA MASTER"

lista_pastas = ["ESQUEMÁTICOS E PLACAS",
                "INSTRUÇÕES",
                "IMAGENS",
                "MAGNÉTICOS",
                "MECÂNICOS",
                "OAC",
                "OUTROS",
                "TESTE",
                "PROGRAMA SMD",
                "ESPECIFICAÇÃO DE CLIENTE",
                "QUALIDADE"]

ppap_pastas = ["01 Registro de Projeto incluindo dados de IMDS",
               "02 Documentos de Alteraçao de Engenharia se houver",
               "03 Aprovaçao de Engenharia do cliente",
               "04 FMEA de Projeto",
               "05 Diagrama de Fluxo de Processo",
               "06 FMEA de Processo",
               "07 Plano de Controle da Produçao",
               "08 Estudo e Analise do Sistema de Mediçao",
               "09 Resultados Dimensionais, com desenho marcado (boleado ou quadrantes)",
               "10 Resultados de Ensaios de Material Desempenho",
               "11 Estudos Iniciais do Processo (Capabilidade do Processo)",
               "12 Documentaçao de Laboratorio Qualificado",
               "13 Relatorio de Aprovaçao de Aparencia (RAA), se aplicavel",
               "14 Amostra de Produto",
               "15 Amostra Padrao",
               "16 Auxilio de Verificaçao",
               "17 Registros de Conformidade com os Requisitos Especificos do Cliente",
               "18 Certificado de Submissao de Peça (PSW)"]

tipos = {"M-": r"\\servidor\instrucoes-leitura\MECANICOS",
         "MTSM-": r"\\servidor\instrucoes-leitura\MECANICOS",
         "T-": r"\\servidor\instrucoes-leitura\MECANICOS\COMPONENTES\CABOS",
         "F-": r"\\servidor\instrucoes-leitura\MAGNÉTICOS",
         "J-": r"\\servidor\desenv\DOCUMENTOS\LISTA PLACAS SIGE",
         "E-": r"\\servidor\instrucoes-leitura\ETIQUETAS"}

pasta_tipos = {"M-": "MECÂNICOS",
               "MTSM-": "MECÂNICOS",
               "T-": "MECÂNICOS",
               "F-": "MAGNÉTICOS",
               "J-": "ESQUEMÁTICOS E PLACAS",
               "E-": "MECÂNICOS"}

while True:
    if not os.path.exists("S:\\PASTA MASTER"):
        mapeamento("instrucoes", "S:")
    elif not os.path.exists(r"H:\DOCUMENTOS"):
        mapeamento("desenv", "H:")
    elif not os.path.exists(r"F:\FA-4200ATE"):
        mapeamento("ate300", "F:")
    else:
        break

while True:
    pasta_a_ser_criada = input("Digite o grupo para criar a pasta: ").upper()
    if not re.fullmatch(r"G[0DK][0-9]{4,5}", pasta_a_ser_criada):
        print("Entrada inválida!")
    else:
        break

while True:
    caminho_dataframe = input(
        "Digite o caminho da planilha da lista de peças: ").strip('"')
    try:
        book = xlrd.open_workbook(caminho_dataframe, encoding_override="ansi")
        df = pd.read_excel(book, engine="xlrd", dtype="object")
        break
    except:
        print("planilha inválida!")

caminho_pasta_a_ser_criada = os.path.join(diretorio_master, pasta_a_ser_criada)

if not os.path.exists(caminho_pasta_a_ser_criada):
    os.makedirs(caminho_pasta_a_ser_criada)
for pastas in lista_pastas:
    subpasta = os.path.join(caminho_pasta_a_ser_criada, pastas)
    if not os.path.exists(subpasta):
        os.makedirs(subpasta)
if not os.path.exists(os.path.join(caminho_pasta_a_ser_criada, "OUTROS", "CHECK-LISTS ANTIGAS")):
    os.makedirs(os.path.join(caminho_pasta_a_ser_criada,
                "OUTROS", "CHECK-LISTS ANTIGAS"))

# Criação das pastas PPAP dentro de QUALIDADE
ppap_path = os.path.join(caminho_pasta_a_ser_criada, "QUALIDADE", "PPAP")
if not os.path.exists(ppap_path):
    os.makedirs(ppap_path)
for ppap_subpasta in ppap_pastas:
    subpasta = os.path.join(ppap_path, ppap_subpasta)
    if not os.path.exists(subpasta):
        os.makedirs(subpasta)

for i in range(len(df.columns)):
    for k in range(len(df)):
        celula = df.iat[k, i]
        if pd.isnull(celula):
            continue
        for itens in tipos.keys():
            if str(celula).startswith(itens):
                if str(celula).startswith("J-"):
                    if re.fullmatch(r"J-[0-9]{4}", str(celula)):
                        criar_atalho(tipos[itens], celula)
                elif str(celula).startswith("F-"):
                    if not re.fullmatch(r"F-[0-3][0-9]{2}A?", str(celula)):
                        criar_atalho(tipos[itens], celula)
                elif str(celula).startswith("T-"):
                    if not re.fullmatch(r"T-[0-4][0-9]{2}A?", str(celula)):
                        criar_atalho(tipos[itens], celula)
                    else:
                        continue
                else:
                    criar_atalho(tipos[itens], celula)
        else:
            continue

for pasta_sat in os.listdir("F:"):
    if pasta_sat == pasta_a_ser_criada:
        criar_atalho_teste(r"\\servidor\ate300", pasta_sat + " - SAT")
        break
else:
    for pasta_manual in os.listdir(r"S:\PLANILHAS DE TESTE MANUAL\G0"):
        if pasta_manual == pasta_a_ser_criada:
            criar_atalho_teste(
                r"\\servidor\instrucoes-leitura\PLANILHAS DE TESTE MANUAL\G0", pasta_manual + " - Planilha")
            break

for pasta_fotos in os.listdir(r"S:\FOTOGRAFIAS & IM'S\SIGE"):
    if pasta_fotos == pasta_a_ser_criada:
        link_filepath = os.path.join(
            diretorio_master, pasta_a_ser_criada, "INSTRUÇÕES E IMAGENS", pasta_a_ser_criada + " - Site.lnk")
        with winshell.shortcut(link_filepath) as link:
            link.path = os.path.join(
                r"\\servidor\instrucoes-leitura\SITE\FOTOS\FOTOS PRODUTOS", pasta_a_ser_criada)

for pasta_smd in os.listdir("S:\PROGRAMAS SMD"):
    if pasta_smd == pasta_a_ser_criada:
        criar_atalho_smd(r"\\servidor\instrucoes-leitura\PROGRAMAS SMD", pasta_smd)
        #link_filepath = os.path.join(
           # diretorio_master, pasta_a_ser_criada, "PROGRAMA SMD", pasta_a_ser_criada + " -
