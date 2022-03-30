import pandas as pd

print("Carregando arquivo de dados")
df_vendas = pd.read_excel("df_desafio.xlsx", sheet_name="Dados - Questão 1", skiprows=0)
df_compras = pd.read_excel("df_desafio.xlsx", sheet_name="Dados - Questão 2", skiprows=0)

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
df_compras = df_compras[['Valor Total', 'Dinheiro de Volta (Aplicado direto no total)']]
df_compras = df_compras.rename(columns={'Valor Total':'valor_total',
                                        'Dinheiro de Volta (Aplicado direto no total)':'retorno_financeiro'})

print("Calculando a filial com maior renda")
df_filial = df_vendas[['nome_filial', 'valor_da_compra']]
df_filial = df_filial.groupby(['nome_filial']).agg({'valor_da_compra' : 'sum'}).round(2)
df_filial = df_filial.sort_values(by="valor_da_compra", ascending=False).reset_index()
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
percentual_uso_cpf = relacao_cpf.percentual_sim
print(percentual_uso_cpf)

print("Calculando percentual de imposto relativo ao valor total de compras")
valor_total_imposto = df_vendas.imposto.sum().round(3)
valor_total_compra = df_vendas.valor_da_compra.sum()
percentual_imposto = ((valor_total_imposto/valor_total_compra)*100).round(2)
print("Percentual de imposto: %", percentual_imposto)

print("Calculando percentual de devolucoes")
relacao_devolucoes = (df_vendas['devolucoes'].value_counts()).round(2)
df_devolucoes = pd.DataFrame(relacao_devolucoes).reset_index()
df_devolucoes = df_devolucoes.rename(columns={'index': 'escolha', 'devolucoes': 'quantidade'})
percentual_devolucao = ((relacao_devolucoes['Sim']/relacao_devolucoes.sum())*100).round(2)
print("Percentual de devolucoes: %", percentual_devolucao)


print('Analisando motivos das devoluções')
motivos_devolucoes = df_vendas['motivo_devolucao'].value_counts()
dict_motivos ={
        "Situacao" : [
            "defeito_produto",
            "insatisfacao_produto",
            "problema_entrega",
            "antecipacao_troca",
            "insastifacao_atendimento"
        ],
        "Quantidade" : [
            motivos_devolucoes[0],
            motivos_devolucoes[1],
            motivos_devolucoes[2],
            motivos_devolucoes[3],
            motivos_devolucoes[4],
        ],
        "Percentual" : [
            ((motivos_devolucoes[0]/relacao_devolucoes['Sim'].sum())*100).round(2),
            ((motivos_devolucoes[1]/relacao_devolucoes['Sim'].sum())*100).round(2),
            ((motivos_devolucoes[2]/relacao_devolucoes['Sim'].sum())*100).round(2),
            ((motivos_devolucoes[3]/relacao_devolucoes['Sim'].sum())*100).round(2),
            ((motivos_devolucoes[4]/relacao_devolucoes['Sim'].sum())*100).round(2)
        ]
    }
df_motivos = pd.DataFrame(dict_motivos)
print("Motivos das devolucoes \n", df_motivos)

print("Analisando a situação dos vendedores")
df_vendedor = df_vendas[['nome_vendedor', 'valor_da_compra', 'devolucoes']]
df_vendedor = df_vendedor.groupby(['devolucoes', 'nome_vendedor']).agg({'valor_da_compra': 'sum'}).reset_index()
index_devolvidos = df_vendedor[df_vendedor["devolucoes"]=="Sim"].index
df_vendedor = df_vendedor.drop(index_devolvidos)
df_vendedor = df_vendedor.sort_values(by="valor_da_compra", ascending=False).head().reset_index()
df_vendedor = df_vendedor.drop("index", axis=1)
vendedor_que_mais_vendeu = df_vendedor.nome_vendedor[0]
print("Os melhores vendedores foram: \n", df_vendedor)


print("Analizando os produto mais vendidos ")
relacao_de_produtos = df_vendas["nome_produto"].value_counts().sort_values(ascending=False).head(10)
print("Os Top produtos mais vendidos foram ", relacao_de_produtos)
df_produtos = pd.DataFrame(relacao_de_produtos).reset_index()
df_produtos =  df_produtos.rename(columns={'index': 'nome_produto','nome_produto': 'quantidade'})

print("Analisando o retorno financeiro ")
percentual_de_retorno = df_compras.agg({'valor_total':'sum', 'retorno_financeiro':'sum'})
percentual_de_retorno = ((percentual_de_retorno.retorno_financeiro / percentual_de_retorno.valor_total) * 100).round(2)
print("O retorno financeiro das compras e de %", percentual_de_retorno)

dict_percentuais = {
    'percentual' : [
        'uso do cpf',
        'imposto',
        'retorno financeiro'
    ],
    'valor (%)' : [
        percentual_uso_cpf,
        percentual_imposto,
        percentual_de_retorno
    ]
}
df_percentuais = pd.DataFrame(dict_percentuais)
print("Tabela de percentuais \n", df_percentuais)
#saída?!

#df.to_excel("nome_planilha.xlsx", sheet_name="nome_aba", index=False)
#with pd.ExcelWriter("path_to_file.xlsx") as writer:
#    df1.to_excel(writer, sheet_name="Sheet1")
#    df2.to_excel(writer, sheet_name="Sheet2")
print("Salvando arquivos em uma planilha")
with pd.ExcelWriter("Dados processados.xlsx") as writer:
    df_filial.to_excel(writer, sheet_name="df_filial", index=False)
    df_percentuais.to_excel(writer, sheet_name="df_percentuais", index=False)
    df_devolucoes.to_excel(writer, sheet_name="df_devolucoes", index=False)
    df_vendedor.to_excel(writer, sheet_name="df_vendedor", index=False)
    df_produtos.to_excel(writer, sheet_name="df_produtos", index=False)
    df_motivos.to_excel(writer, sheet_name="df_motivos_devolucao", index=False)