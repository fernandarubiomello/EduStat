# Sistema_EduStat_Console

O EduStat é um sistema em Python que realiza análise automática de notas acadêmicas a partir de arquivos CSV ou Excel.

O programa:
- Valida rigorosamente os dados fornecidos pelo usuário
- Calcula média, desvio padrão e variância por matéria e no geral
- Indica a situação do aluno (Aprovado / Recuperação / Reprovado)
- Gera histograma e boxplot
- Insere automaticamente os resultados e gráficos em uma planilha Excel final
  
O objetivo é facilitar a análise estatística de desempenho acadêmico de forma clara, automática e segura.

📂 Estrutura de Pastas

Projeto/
│
├── principal.py
│
├── uteis/
│   └── funcoes.py
│
├── dados/
│   ├── (arquivos .csv ou .xlsx/.xlsm inseridos pelo usuário)
│   ├── (arquivo .xlsx gerado caso o input seja CSV)
│   ├── histograma.png
│   └── boxplot.png

📄 Formatos de Arquivos Aceitos

CSV
- Deve estar em UTF-8
- Separado por vírgulas

Excel
- .xlsx ou .xlsm
- Preferencialmente sem formatações avançadas

📐 Regras de Formatação dos Dados

Cabeçalho:
- A1 deve conter um texto, por exemplo: "Matéria"
- A primeira linha deve conter texto, por exemplo: os nomes das avaliações
  Nota 1, Nota 2, Nota Prova, etc.)

Coluna A:
- Deve conter apenas texto, por exemplo: os nomes das matérias
- Não pode haver células vazias no meio dos dados

Notas:
- Devem ser numéricas
- Inteiros ou decimais (ex: 7, 8.5, 9,3)
- Não é permitido deixar espaços vazios entre notas
- Cada matéria deve possuir ao menos uma nota

Observações Importantes:
- Linhas completamente vazias são removidas automaticamente
- Evite cores, bordas, filtros ou estilos no Excel
- O EduStat informa exatamente onde está o erro, se houver

🛠️ Bibliotecas Utilizadas

Bibliotecas padrão (já vêm com o Python):
- os
- glob
- csv
- statistics
- warnings
- sys

Bibliotecas externas (necessitam instalação):
- openpyxl
- matplotlib

▶️ Como Executar o Projeto

- Clone ou baixe este repositório
- Coloque seu arquivo .csv ou .xlsx/.xlsm dentro da pasta /dados
- Abra o terminal na pasta do projeto
- Execute apenas o arquivo: python principal.py
- Siga as instruções exibidas no terminal
- Ao final, abra a planilha gerada para visualizar os resultados

Caso o arquivo não esteja no formato correto, o EduStat não encerra abruptamente, ele exibe mensagens de erro claras, basta corrigir o erro indicado e executar novamente.

📝 Observações Finais

- A pasta /dados do projeto contém exemplos para teste
- Os gráficos gerados também ficam salvos como imagens na pasta /dado
