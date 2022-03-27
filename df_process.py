import pandas as pd

df_vendas = pd.read_excel("df_desafio.xlsx", sheet_name="Dados - Questão 1", skiprows=0)

df_vendas = df_vendas.drop(['Nº Compra', 'Usuario', 'Data da Compra'], axis='columns')

df_vendas = df_vendas.rename(columns={'Nome':'nome_vendedor',
                          'Tipo de Mercadoria':'nome_produto',
                          'Filial':'nome_filial',
                          'Valor da Compra':'valor_da_compra',
                          'Imposto':'imposto',
                          'CPF Na Nota':'cpf',
                          'Produto Devolvido':'devolucoes',
                          'Motivo Devolução':'motivo_devolucao'})

df_filial = df_vendas[['nome_filial', 'valor_da_compra']]
df_filial = df_filial.groupby(['nome_filial']).agg({'valor_da_compra' : 'sum'}).reset_index()
df_filial = df_filial.sort_values(by="valor_da_compra", ascending=False).head().reset_index()
filial_maior_venda = df_filial.nome_filial[0]
relacao_cpf = df_vendas["cpf"].value_counts()


print(relacao_cpf)
print(filial_maior_venda)
print(df_filial)
print(df_vendas.cpf)