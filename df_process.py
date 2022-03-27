import pandas as pd

print("Carregando arquivo de dados")
df_vendas = pd.read_excel("df_desafio.xlsx", sheet_name="Dados - Questão 1", skiprows=0)

print("Configurando colunas do dateframe")
df_vendas = df_vendas.drop(['Nº Compra', 'Usuario', 'Data da Compra'], axis='columns')
df_vendas = df_vendas.rename(columns={'Nome':'nome_vendedor',
                          'Tipo de Mercadoria':'nome_produto',
                          'Filial':'nome_filial',
                          'Valor da Compra':'valor_da_compra',
                          'Imposto':'imposto',
                          'CPF Na Nota':'cpf',
                          'Produto Devolvido':'devolucoes',
                          'Motivo Devolução':'motivo_devolucao'})

print("Calculando a filial com maior renda")
df_filial = df_vendas[['nome_filial', 'valor_da_compra']]
df_filial = df_filial.groupby(['nome_filial']).agg({'valor_da_compra' : 'sum'}).reset_index()
df_filial = df_filial.sort_values(by="valor_da_compra", ascending=False).head().reset_index()
filial_maior_venda = df_filial.nome_filial[0]
print(df_filial)
print(filial_maior_venda)

print("Calculando o percentual relativo ao cpf")
relacao_cpf = df_vendas["cpf"].value_counts()
relacao_cpf["percentual_sim"] = (relacao_cpf['Sim']/relacao_cpf.sum())*100
relacao_cpf["percentual_nao"] = (relacao_cpf['Não']/relacao_cpf.sum())*100
relacao_cpf = relacao_cpf.round(2)
## Corrigindo margen de erro percentual
relacao_cpf.percentual_nao += (100 - (relacao_cpf.percentual_nao + relacao_cpf.percentual_sim))
print(relacao_cpf)

print("Calculando percentual de imposto relativo ao valor total de compras")
valor_total_imposto = df_vendas.imposto.sum().round(3)
valor_total_compra = df_vendas.valor_da_compra.sum()
percentual_imposto = ((valor_total_imposto/valor_total_compra)*100).round(2)
print("Percentual de imposto: %", percentual_imposto)

print("Calculando percentual de devolucoes")
relacao_devolucoes = (df_vendas['devolucoes'].value_counts()).round(2)
percentual_devolucao = ((relacao_devolucoes['Sim']/relacao_devolucoes.sum())*100).round(2)
print("Percentual de devolucoes: %", percentual_devolucao)




