import json
import os
import pandas as pd

import unicodedata


def remove_acentos(text):
    return "".join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')


def analisa_planilha(arquivo_planilha: str, categorias: list[str], palavras: list[str], pasta_textos=None):
    if pasta_textos is None:
        folder = os.getcwd() + '/Documentos/Indenizáveis/Textos/'
    else:
        folder = pasta_textos

    if not os.path.exists(folder):
        raise ValueError("Não foi possível encontrar pasta com os arquivos de texto")

    planilha_folder, planilha_file = os.path.split(arquivo_planilha)
    dfs = pd.read_excel(arquivo_planilha, header=2)
    for index, row in dfs.iterrows():
        cat = row['COD categoria']
        if cat not in categorias:
            continue
        documento = row['COD documento']
        file = str(documento) + ".txt"
        file_path = os.path.join(folder, file)
        try:
            with open(file_path, encoding="utf-8") as arquivo:
                doc_text = arquivo.read()
                doc_text = remove_acentos(doc_text.lower())
                for palavra in palavras:
                    encontrado = ''
                    if remove_acentos(palavra.lower()) in doc_text:
                        encontrado = "SIM"
                    else:
                        encontrado = "NÃO"
                    dfs.loc[index, palavra] = encontrado
        except:
            print(f"Arquivo {file} não encontrado")
            dfs.loc[index, palavras[0]] = f'Arquivo {file} não Encontrado'

    dfs.to_excel(os.path.join(planilha_folder, "resultado.xlsx"), index=False)
