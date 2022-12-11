# Importa a biblioteca openpyxl
import openpyxl
# Importa a biblioteca datetime
from datetime import datetime
# Importa a função mean do módulo statistics
from statistics import mean

# Cria uma lista vazia para armazenar as linhas do diário
diary = []

# Define o formato de data a ser utilizado
date_format = '%d/%m/%Y'

# Pede ao usuário para inserir as informações de cada linha do diário
while True:
    # Pergunta ao usuário se deseja adicionar uma nova linha
    add_new = input('Deseja adicionar uma nova linha ao diário? (s/n)')

    # Se o usuário não desejar adicionar uma nova linha, sai do loop
    if add_new.lower() == 'n':
        break

    # Pede ao usuário para inserir as informações da linha
    date_str = input('Insira a data no formato DD/MM/AAAA:')

    # Valida a data inserida pelo usuário
    try:
        date = datetime.strptime(date_str, date_format)
    except ValueError:
        print('Data inválida. Tente novamente.')
        continue

    activity = input('Insira a atividade realizada:')
    grade = input('Insira a nota (se não houver, insira "-"):')
    falta = input('Insira a falta: ')

    # Cria um dicionário com as informações da linha e o adiciona à lista
    diary.append({'date': date, 'activity': activity, 'grade': grade, 'falta': falta})

# Pede ao usuário para inserir o nome do arquivo
file_name = input('Insira o nome do arquivo (com a extensão .xlsx):')

# Cria um novo arquivo Excel
wb = openpyxl.Workbook()

# Cria uma nova planilha
sheet = wb.active

# Adiciona os títulos das colunas na primeira linha da planilha
sheet.append(['Data', 'Atividade', 'Nota', 'Falta'])

# Percorre a lista de linhas do diário
for row in diary:
    # Converte a string de notas em uma lista de números
    grades = [float(x) for x in row['grade'].split(',')]

    # Calcula a média das notas
    average = mean(grades)

    # Adiciona uma nova linha com os dados da atividade, incluindo a média
    sheet.append([row['date'].strftime(date_format), row['activity'], row['grade'], row['falta'], average])

# Salva o arquivo Excel
wb.save(file_name)
