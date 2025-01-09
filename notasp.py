import pandas as pd
import xmltodict
import os

caminho_pasta = r'C:\Users\Nicolas\Desktop\Notas'

arquivos = os.listdir(caminho_pasta)

lista_arquivos = []
for arquivo in arquivos:
    if 'xml' in arquivo and 'DANFE' in arquivo:
        lista_arquivos.append((arquivo))

def ler_xml_danfe(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

        dict_notafiscal = documento['nfeProc']['NFe']['infNFe']

        data_emissao = dict_notafiscal['ide']['dhEmi']
        nome_fornecedor = dict_notafiscal['emit']['xNome']

        produtos = dict_notafiscal['det']
        if isinstance(produtos, dict):
            produtos = [produtos]  
        lista_produtos = []
        for produto in produtos:
            quantidade = float(produto['prod']['qCom'])  # Convertendo para float
            descricao = produto['prod']['xProd']
            valor_unitario = float(produto['prod']['vUnCom'])  # Convertendo para float
            lista_produtos.append((quantidade, descricao, valor_unitario))

        cnpj_fornecedor = documento['nfeProc']['NFe']['infNFe']['emit']['CNPJ']
        numero_nota = documento['nfeProc']['NFe']['infNFe']['ide']['cNF']

        resposta = {
            'data_emissao': data_emissao,
            'nome_fornecedor': nome_fornecedor,
            'lista_produtos': lista_produtos,
            'cnpj_fornecedor': cnpj_fornecedor,
            'numero_nota': numero_nota,
        }
        return resposta

df_final = pd.DataFrame()
for arquivo in lista_arquivos:
    df = pd.DataFrame.from_dict(ler_xml_danfe(os.path.join(caminho_pasta, arquivo)))
    df_final = pd.concat([df_final, df], ignore_index=True)

df_final[['quantidade', 'descricao', 'valor_unitario']] = pd.DataFrame(df_final['lista_produtos'].tolist(), index=df_final.index)

df_final.drop(columns=['lista_produtos'], inplace=True)

df_final.rename(columns={'data_emissao': 'DATA DE EMISSAO', 'nome_fornecedor': 'NOME DO FORNECEDOR'}, inplace=True)

df_final['valor_unitario'] = df_final['valor_unitario'].astype(float)

df_final['valor_unitario'] = df_final['valor_unitario'].astype(str).str.replace('.', ',')

df_final['quantidade'] = df_final['quantidade'].astype(float).round().astype(int)

df_final['valor_unitario'] = df_final['valor_unitario'].str.replace(',', '.').astype(float)

df_final['VALOR TOTAL'] = df_final['quantidade'] * df_final['valor_unitario']


df_final['DATA DE EMISSAO'] = pd.to_datetime(df_final['DATA DE EMISSAO']).dt.strftime('%d/%m/%Y')

df_final['CNPJ DO FORNECEDOR'] = df_final['cnpj_fornecedor']
df_final['NUMERO DA NOTA'] = df_final['numero_nota']

df_final['LINHA ITENS'] = df_final.groupby('NUMERO DA NOTA').cumcount() + 1

df_final['ID FORNECEDOR'] = ''
df_final['CENTRO DE CUSTO FATURA'] = ''
df_final['FORMA DE PAGAMENTO (NEXXERA)'] = ''
df_final['ITEM'] = ''
df_final['CENTRO DE CUSTO ITEM'] = ''
df_final['MEMO'] = ''
df_final['TIPO DOCUMENTO'] = 'Fatura'
df_final['TIPO DE NOTA'] = 'NF-e'
df_final['CONDICAO DE PAGAMENTO'] = 'A Vista'
df_final['TIPO DE FORNECEDOR'] = '2'

df_final.rename(columns={'NUMERO DA NOTA': 'N DE REFERENCIA'}, inplace=True)

df_final = df_final[['N DE REFERENCIA', 'NOME DO FORNECEDOR', 'ID FORNECEDOR', 'DATA DE EMISSAO', 'CENTRO DE CUSTO FATURA',
                     'FORMA DE PAGAMENTO (NEXXERA)', 'TIPO DOCUMENTO', 'TIPO DE NOTA', 'ITEM', 'quantidade',
                     'descricao', 'valor_unitario', 'VALOR TOTAL', 'CENTRO DE CUSTO ITEM', 'CONDICAO DE PAGAMENTO',
                     'TIPO DE FORNECEDOR', 'MEMO', 'LINHA ITENS', 'CNPJ DO FORNECEDOR']]


# Salve o DataFrame em um arquivo Excel
excel_file_path = caminho desejado
df_final.to_excel(excel_file_path, index=False, engine='openpyxl')

