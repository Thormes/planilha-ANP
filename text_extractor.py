import json
import time

from requests import ReadTimeout
from tika import parser
import os

from Logger.logger import get_logger

loger = get_logger("Extrator", "extrator.log")

def extract_text_from_pdfs_recursively(dir):
    text_files = os.listdir(dir + "/Textos")
    processados = []
    for text_file in text_files:
        if text_file.endswith(".txt"):
            processados.append(text_file)

    for root, dirs, files in os.walk(dir):
        total = len(files)
        count = 0
        for file in files:
            count += 1
            path_to_pdf = os.path.join(root, file)
            [stem, ext] = os.path.splitext(path_to_pdf)
            if ext.lower() == '.pdf':
                text_file = file.lower().replace(".pdf", ".txt")
                if text_file in processados:
                    continue
                size = os.stat(path_to_pdf).st_size
                print("Processando " + path_to_pdf, f"({size / 1024:.2f} KB)", count, "/", total)
                loger.info(f"Processando {path_to_pdf} ({size / 1024:.2f} KB) {count}/{total}")
                try:
                    start = time.time()
                    pdf_contents = parser.from_file(path_to_pdf, headers={"X-Tika-PDFocrStrategy": "auto", "X-Tika-Timeout-Millis": "1200000"}, requestOptions={"timeout": 1200})
                    end = time.time()
                    print("Arquivo processado em ", end - start, "segundos")
                    loger.info(f"Arquivo processado em {end - start} segundos")
                    if pdf_contents['content'] is None:
                        raise ConnectionError("Erro de conexão")
                    yield os.path.join(root, "Textos", text_file), str(pdf_contents['content'])
                except (ReadTimeout, ConnectionError) as out:
                    print("Demorou muito ao converter", file, "em texto:", repr(out))
                    loger.error(f"Demorou muito ao converter o conteúdo do arquivo {file} em texto: {repr(out)}")
                    time.sleep(5)
                    continue
                except Exception as ex:
                    print("Erro ao converter o conteúdo do arquivo", file, "em texto:", repr(ex))
                    loger.error(f"Erro ao converter o conteúdo do arquivo {file} em texto: {repr(ex)}")
                    continue


if __name__ == "__main__":
    for file, content in extract_text_from_pdfs_recursively(os.getcwd() + '/Documentos/Indenizáveis/'):
        loger.info(f"Anotando {len(content) / (1024):.2f}KB de conteúdo texto no arquivo {file}")
        print("Anotando conteúdo no arquivo", file)
        document = open(file, 'w', encoding='utf-8')
        document.write(content)
        document.close()
