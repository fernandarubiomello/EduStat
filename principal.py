# EduStat
# Desenvolvido por Maria Olívia Meca e Fernanda Rubio

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
# p tirar esse aviso: C:\Users\55169\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\site-packages\openpyxl\worksheet\_reader.py:329: UserWarning: Unknown extension is not supported and will be removed warn(msg) e outros mais que poderiam aparecer

import glob

import uteis.funcoes as fn
from dados import *

# link para os caracteres especiais: https://pt.piliapp.com/symbol/
# fonte que uso na planilha: Aptos Narrow

print("\nBem-vindo ao EduStat ★ Análise Inteligente de Desempenho Acadêmico")
print("""
Por favor, siga as instruções:

✔ 1. Coloque seu arquivo .csv ou xlsx/.xlsm dentro da pasta /dados.
   • O EduStat aceita arquivos .csv (UTF-8, separado por virgula) ou .xlsx/.xlsm.
   • Se houver mais de um arquivo, você poderá escolher qual abrir.

✔ 2. O arquivo deve seguir o formato correto:
   • A primeira coluna (A) deve conter APENAS os nomes das matérias e a célula A1 deve conter  a palavra 'Matéria'.
   • A primeira linha deve conter os nomes das avaliações (ex: Nota 1, Nota 2...).
   • As notas devem ser numéricas (ex: 7.5, 8, 9.3).
   • Se houver espaços em branco entre as notas, será contado como erro.

✔ 3. Evite formatações avançadas no Excel:
   • Não use cores, bordas, filtros ou estilos desnecessários.

✔ 4. O EduStat irá:
   • Validar automaticamente sua planilha ou CSV.
   • Calcular média aritmética, desvio padrão e variância por matéria e total.
   • Inserir o histograma e o boxplot diretamente na planilha - não compatível em todas as plataformas de planilha, mas Microsoft Excel, por exemplo, aceita.
   • Criar uma versão .xlsx caso você envie um CSV.

✔ 5. Caso apareçam erros:
   • Leia atentamente a mensagem: ela indicará exatamente onde está o problema
     (ex: “Valor inválido em C5: deve ser um número”).
   • Corrija a planilha e execute o EduStat novamente.
""")

input("\nPressione ENTER para começar...")

dados = {}

arquivos_csv = glob.glob('dados/*.csv')
arquivos_xls = glob.glob('dados/*.xls*')

if arquivos_csv and arquivos_xls:

    print("→ Foram encontrados os seguintes arquivos:")
    print(f"CSV : {arquivos_csv}")
    print(f"XLS : {arquivos_xls}")

    tipo = input("Qual deseja abrir? (csv/xls) -> ").lower().strip() 
    while(tipo != "csv" and tipo != "xls"):
        tipo = input("Não consegui identificar o tipo do arquivo, digite novamente... Qual deseja abrir? (csv/xls) -> ")
    
    if (tipo == "csv"):

        quant1 = len(arquivos_csv)

        if (quant1 > 1):
            n = input(f"Qual da lista você deseja abrir? (1 a {quant1}) -> ")
            while not n.isdigit():
                n = input(f"Valor inválido. Por favor, digite um número (1 a {quant1})-> ")
            n = int(n)
            while (n > quant1 or n < 1):
                n = input(f"Número digitado fora do intervalo, digite novamente... Qual da lista você deseja abrir? (1 a {quant1}) -> ")
                while not n.isdigit():
                    n = input(f"Valor inválido. Por favor, digite um número (1 a {quant1})-> ")
                n = int(n)
            fn.identificarCaminhoCSV(arquivos_csv, n, dados)

        else:
            fn.identificarCaminhoCSV(arquivos_csv, 1, dados)

    
    elif (tipo == "xls"):

        quant2 = len(arquivos_xls)

        if (quant2 > 1):
            n = input(f"Qual da lista você deseja abrir? (1 a {quant2}) -> ")
            while not n.isdigit():
                n = input(f"Valor inválido. Por favor, digite um número (1 a {quant2})-> ")
            n = int(n)
            while (n > quant2 or n < 1):
                n = input(f"Número digitado fora do intervalo, digite novamente... Qual da lista você deseja abrir? (1 a {quant2}) -> ")
                while not n.isdigit():
                    n = input(f"Valor inválido. Por favor, digite um número (1 a {quant2})-> ")
                n = int(n)
            fn.identificarCaminhoXLS(arquivos_xls, n, dados)
        else:
            fn.identificarCaminhoXLS(arquivos_xls, 1, dados)

elif arquivos_csv:
    fn.identificarCaminhoCSV(arquivos_csv, 1, dados)

elif arquivos_xls: 
    fn.identificarCaminhoXLS(arquivos_xls, 1, dados)


else:
    print("✘ Nenhum arquivo CSV ou Excel encontrado em /dados. Leia as instruções novamente.")
    print("\nAgradeçemos o uso do EduStat ★ \n")
    print("Até mais...")
