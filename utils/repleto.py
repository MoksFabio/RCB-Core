import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox

def converter_txt_para_xlsx():
    Tk().withdraw()  # Esconde a janela principal do Tkinter

    # Selecionar vários arquivos TXT
    arquivos_txt = filedialog.askopenfilenames(
        title="Selecione os arquivos TXT",
        filetypes=[("Arquivos TXT", "*.txt")]
    )

    if not arquivos_txt:
        messagebox.showinfo("Aviso", "Nenhum arquivo selecionado.")
        return

    # Selecionar pasta de saída
    pasta_saida = filedialog.askdirectory(title="Selecione a pasta de destino")
    if not pasta_saida:
        messagebox.showinfo("Aviso", "Nenhuma pasta de destino selecionada.")
        return

    for arquivo in arquivos_txt:
        try:
            linhas = []
            with open(arquivo, "r", encoding="latin1") as f:
                for linha in f:
                    # Quebra a linha pelo TAB "\t"
                    partes = linha.rstrip("\n").split("\t")
                    linhas.append(partes)

            # Cria DataFrame sem cabeçalhos extras
            df = pd.DataFrame(linhas)

            # Substitui 0,0 e 0,00 por 0 (em todo o DataFrame)
            df = df.replace({"0,0": "0", "0,00": "0", "0,000": "0"})

            # Nome do arquivo de saída
            nome_saida = os.path.splitext(os.path.basename(arquivo))[0] + ".xlsx"
            caminho_saida = os.path.join(pasta_saida, nome_saida)

            # Salva em XLSX
            df.to_excel(caminho_saida, index=False, header=False, engine="openpyxl")
            print(f"Convertido: {caminho_saida}")

        except Exception as e:
            print(f"Erro ao processar {arquivo}: {e}")

    messagebox.showinfo("Concluído", "Conversão finalizada com sucesso!")

if __name__ == "__main__":
    converter_txt_para_xlsx()
