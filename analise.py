import pandas as pd

tb_vendedores = pd.read_excel("database-py.xlsx", sheet_name="vendedores")
tb_clientes = pd.read_excel("database-py.xlsx", sheet_name="clientes")
tb_categorias = pd.read_excel("database-py.xlsx", sheet_name="categorias")
tb_produtos = pd.read_excel("database-py.xlsx", sheet_name="produtos")
tb_vendas = pd.read_excel("database-py.xlsx", sheet_name="vendas")

# adicionar a categoria, o preço e o custo unitario do produto a tabela de vendas
tb_vendas = tb_vendas.merge(tb_produtos)

# adicionar o imposto do produto a tabela de vendas
tb_vendas = tb_vendas.merge(tb_categorias)

# adicionar o vendedor e o percentual de frete a tabela de vendas
tb_vendas = tb_vendas.merge(tb_clientes)

# adicionar o percentual de comissao a tabela de vendas
tb_vendas = tb_vendas.merge(tb_vendedores)

# calcular o valor total vendido por produto
tb_vendas['valor_total'] = tb_vendas['qtde'] * tb_vendas['pr_unit']

# calcular o valor desconto por produto
tb_vendas['valor_desconto'] = tb_vendas['valor_total'] * (tb_vendas['desconto']/100)

# calcular o valor liquido da venda por produto
tb_vendas['valor_liquido'] = tb_vendas['valor_total'] - tb_vendas['valor_desconto']

# calcular o valor da comissao por venda
tb_vendas['valor_comissao'] = round((tb_vendas['valor_liquido'] * (tb_vendas['comissao']/100)), 2)

# calcular o valor do frete por venda
tb_vendas['valor_frete'] = round((tb_vendas['valor_liquido'] * (tb_vendas['frete']/100)), 2)

# calcular o valor do imposto por venda
tb_vendas['valor_imposto'] = round((tb_vendas['valor_liquido'] * (tb_vendas['imposto']/100)), 2)

# calcular o valor total do custo por produto
tb_vendas['custo_total'] = tb_vendas['qtde'] * tb_vendas['custo_unit']

# calcular o total dos abatimentos por produto
tb_vendas['valor_abatimentos'] = tb_vendas['valor_comissao'] + tb_vendas['valor_frete'] +tb_vendas['valor_imposto'] + tb_vendas['custo_total']

# calcular o resultado operacional por produto
tb_vendas['valor_margem'] = tb_vendas['valor_liquido'] - tb_vendas['valor_abatimentos']

# calcular a margem bruta por produto
tb_vendas['perc_margem'] = round((tb_vendas['valor_margem'] / tb_vendas['valor_liquido'] * 100), 2)
print(tb_vendas)

# faturamento por produto
fat_produto = tb_vendas[['nome_produto', 'qtde', 'valor_liquido']].groupby('nome_produto').sum()
fat_produto['pr_med'] = round((fat_produto['valor_liquido'] / fat_produto['qtde']), 2)
print('\nFaturamento por Produto')
print(fat_produto)

# faturamento por categoria
fat_categoria = tb_vendas[['nome_categoria', 'qtde', 'valor_liquido']].groupby('nome_categoria').sum()
fat_categoria['pr_med'] = round((fat_categoria['valor_liquido'] / fat_categoria['qtde']), 2)
print('\nFaturamento por Categoria')
print(fat_categoria)

# faturamento por cliente
fat_cliente = tb_vendas[['nome_cliente', 'qtde', 'valor_liquido']].groupby('nome_cliente').sum()
fat_cliente['pr_med'] = round((fat_cliente['valor_liquido'] / fat_cliente['qtde']), 2)
print('\nFaturamento por Cliente')
print(fat_cliente)

# faturamento por vendedor
fat_vendedor = tb_vendas[['nome_vendedor', 'qtde', 'valor_liquido']].groupby('nome_vendedor').sum()
fat_vendedor['pr_med'] = round((fat_vendedor['valor_liquido'] / fat_vendedor['qtde']), 2)
print('\nFaturamento por Vendedor')
print(fat_vendedor)

