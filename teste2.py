from tabula import read_pdf # Windows / MacOS: pip install tabula-py
                            # Linux: pip3 install tabula-py

import pandas as pd # Windows / MacOS: pip install pandas
                    # Linux: pip3 install pandas

import zipfile as z # Windows / MacOS: pip install zipfile38
                    # Linux: pip3 install zipfile38
import sys
import os


def extrairQuadro30(arquivo, paginas, nomeOut):
    listaTabelas = read_pdf(arquivo, encoding='utf-8', pages=paginas, silent=True)
    df = listaTabelas[0]
    df.columns = df.iloc[0]
    df[[0,1]] = df['Código Descrição da categoria'].str.split(n=1, expand=True)
    df = df[0:]
    df.columns = df.iloc[0]
    df = df[1:]
    df = df.drop('Código Descrição da categoria', axis=1)
    df.to_csv((nomeOut+'.csv'), encoding='utf-8')


def extrairQuadro31(arquivo, paginas, nomeOut):
    listaTabelas = read_pdf(arquivo, encoding='utf-8', pages=paginas, silent=True)
    df = listaTabelas[0]
    df.columns = df.iloc[0]
    df = df[1:]
    for tabela in listaTabelas[1:-1]:
        df = pd.concat([df, tabela], axis=0)
    df.to_csv((nomeOut+'.csv'), encoding='utf-8')


def extrairQuadro32(arquivo, paginas, nomeOut):
    listaTabelas = read_pdf(arquivo, encoding='utf-8', pages=paginas, silent=True)

    df = listaTabelas[1]
    df = df[1:]
    df = df.dropna()
    df.columns = df.iloc[0]
    df[['Categoria','Descrição da Categoria']] = df['Descrição da categoria'].str.split(n=1, expand=True)
    df = df[1:]
    df = df.drop('Descrição da categoria', axis=1)
    df = df.rename(index={2: 1, 3: 2, 5: 3})
    df.to_csv((nomeOut+'.csv'), encoding='utf-8')


os.system('cls')

print('\n -----------------------------')
print('| EXTRAIR QUADROS 30, 31 E 32 |')
print(' -----------------------------')

if os.path.exists('./padrao_tiss_componente_organizacional_202108.pdf'):
    print('\nArquivo referente ao Componente Organizacional do Padrão TISS localizado\n')
    arquivo = 'padrao_tiss_componente_organizacional_202108.pdf'
else:
    print('\nNão foi possível localizar o arquivo solicitado')
    input('Pressione qualquer tecla para sair...\n')
    sys.exit()

zipCsvs = z.ZipFile('Teste_Intuitive_Care_Rodriggo_Mendes.zip', 'w')
comSucesso = ''

extrairQuadro30(arquivo, '108', 'quadro_30')
if os.path.exists('./quadro_30.csv'):
    print('Quadro 30 extraído com sucesso')
    zipCsvs.write('quadro_30.csv')
    comSucesso += '30, '
else:
    print('Falha ao extrair Quadro 30')

extrairQuadro31(arquivo, '109-114', 'quadro_31')
if os.path.exists('./quadro_31.csv'):
    print('Quadro 31 extraído com sucesso')
    zipCsvs.write('quadro_31.csv')
    comSucesso += '31, '
else:
    print('Falha ao extrair Quadro 31')

extrairQuadro32(arquivo, 114, 'quadro_32')
if os.path.exists('./quadro_32.csv'):
    print('Quadro 32 extraído com sucesso')
    zipCsvs.write('quadro_32.csv')
    comSucesso += '32 '
else:
    print('Falha ao extrair Quadro 32')

zipCsvs.close()

os.remove("quadro_30.csv")
os.remove("quadro_31.csv")
os.remove("quadro_32.csv")

print('\nQuadros ' + comSucesso  + 'zipados em "Teste_Intuitive_Care_Rodriggo_Mendes.zip"')
input('Pressione qualquer tecla para sair...\n')