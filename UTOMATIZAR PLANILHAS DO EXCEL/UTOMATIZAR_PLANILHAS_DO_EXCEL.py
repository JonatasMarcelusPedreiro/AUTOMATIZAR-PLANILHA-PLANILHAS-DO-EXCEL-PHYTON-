#Para começar, precisamos instalar a biblioteca openpyxl, que é a principal biblioteca para manipulação de arquivos Excel em Python.

#Para instalar, basta executar o seguinte comando:


#pip install openpyxl
#Com a biblioteca instalada, podemos começar a escrever código para automatizar tarefas no Excel.


#Vamos começar criando uma nova planilha em branco com o nome “Exemplo.xlsx”.

#Para fazer isso, utilizaremos a seguinte sequência de comandos em Python:


from openpyxl import Workbook

# Cria uma nova planilha em branco
wb = Workbook()

# Seleciona a planilha ativa
planilha = wb.active

# Define o valor da célula A1
planilha['A1'] = 'Olá, Excel com Python!'

# Salva o arquivo
wb.save('Exemplo.xlsx')


#Neste exemplo, importamos a classe Workbook da biblioteca openpyxl.
#Em seguida, criamos um objeto wb do tipo Workbook, que representa a planilha em branco.
#Em seguida, selecionamos a planilha ativa usando o atributo active.
#Depois disso, definimos o valor da célula A1 da planilha usando a notação de índice 'A1'.
#No exemplo, atribuímos o texto Olá, Excel com Python! para a célula A1.
#Por fim, salvamos o arquivo com o nome “Exemplo.xlsx” usando o método save do objeto wb.


#Além de criar uma nova planilha, também podemos usar o Python para ler dados de uma planilha existente.

#Para fazer isso, utilizamos a seguinte sequência de comandos:

from openpyxl import load_workbook

# Carrega o arquivo existente
wb = load_workbook('Exemplo.xlsx')

# Seleciona a planilha ativa
planilha = wb.active

# Lê o valor da célula A1
valor_celula = planilha['A1'].value

# Imprime o valor na tela
print(valor_celula)


#Neste exemplo, importamos a função load_workbook da biblioteca openpyxl.
#Em seguida, carregamos o arquivo “Exemplo.xlsx” usando a função load_workbook e atribuímos o objeto resultante à variável wb.
#Depois disso, selecionamos a planilha ativa utilizando o atributo active.
#Em seguida, lemos o valor da célula A1, utilizando a notação de índice 'A1', e atribuímos o valor à variável valor_celula.
#Por fim, imprimimos o valor da célula A1 na tela utilizando a função print.






#Atualizando dados em uma planilha existente
#Além de ler dados de uma planilha existente, também podemos atualizar os dados usando o Python.

#Para isso, utilizamos a seguinte sequência de comandos:


from openpyxl import load_workbook

# Carrega o arquivo existente
wb = load_workbook('Exemplo.xlsx')

# Seleciona a planilha ativa
planilha = wb.active

# Atualiza o valor da célula A1
planilha['A1'] = 'Novo valor'

# Salva o arquivo
wb.save('Exemplo.xlsx')

#Neste exemplo, importamos a função load_workbook da biblioteca openpyxl.
#Em seguida, carregamos o arquivo “Exemplo.xlsx” usando a função load_workbook e atribuímos o objeto resultante à variável wb.
#Depois disso, selecionamos a planilha ativa utilizando o atributo active.
#Em seguida, atualizamos o valor da célula A1, utilizando a notação de índice 'A1' e atribuímos o novo valor.
#Por fim, salvamos o arquivo usando o método save do objeto wb.
