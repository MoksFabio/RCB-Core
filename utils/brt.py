import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def gerar_relatorio_brt_final_robusto():
    """
    Script final e robusto que lida com diferentes formatos de linha nos arquivos TXT.
    """
    print("Iniciando o gerador de relatório de passageiros BRT...")
    
    root = tk.Tk()
    root.withdraw()
    caminhos_arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos TXT originais",
        filetypes=(("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*"))
    )

    if not caminhos_arquivos:
        print("Nenhum arquivo selecionado. O programa será encerrado.")
        return

    dados_compilados = []
    mapa_meses = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    
    for caminho_arquivo in caminhos_arquivos:
        nome_arquivo = os.path.basename(caminho_arquivo)
        try:
            print(f"Processando arquivo: {nome_arquivo}...")

            partes_nome = nome_arquivo.split('-')
            data_str = partes_nome[1]
            
            ano = 2000 + int(data_str[0:2])
            mes_num = int(data_str[2:4])
            dia = int(data_str[4:6])
            quinzena = 1 if dia <= 15 else 2
            mes_nome = mapa_meses[mes_num]

            with open(caminho_arquivo, 'r', encoding='latin-1') as f:
                linhas = f.readlines()

            linha_alvo_passageiro_total = None
            for linha in linhas:
                if "Passageiro Total" in linha:
                    linha_alvo_passageiro_total = linha
                    break
            
            if linha_alvo_passageiro_total is None:
                print(f"  -> AVISO: Não foi encontrada a linha 'Passageiro Total' no arquivo. Pulando.")
                continue

            partes = linha_alvo_passageiro_total.split()
            
            valor_total = None
            linha_numero = None

            # Verifica se o primeiro item é um número (Formato 1)
            if partes[0].isdigit():
                valor_total = int(partes[0])
                linha_numero = partes[5]
            # Se não for número, é o Formato 2
            else:
                # No Formato 2, o TOTAL está na primeira linha de dados (linha do cabeçalho + 1)
                # E a linha está na 5a posição (índice 4)
                primeira_linha_dados = linhas[1].split()
                valor_total = int(primeira_linha_dados[0])
                linha_numero = partes[4]

            dados_arquivo = {
                'Linha': linha_numero,
                'TOTAL': valor_total,
                'Ano': ano,
                'Mês': mes_nome,
                'Quinzena': quinzena
            }
            dados_compilados.append(dados_arquivo)
            print(f"  -> Processado com sucesso. Linha: {linha_numero}, Total: {valor_total}")

        except Exception as e:
            print(f"  -> ERRO ao processar o arquivo {nome_arquivo}: {e}")

    if not dados_compilados:
        print("\nNenhum dado válido foi processado. Verifique seus arquivos.")
        return

    df_consolidado = pd.DataFrame(dados_compilados)
    
    ordem_meses = list(mapa_meses.values())
    df_consolidado['Mês'] = pd.Categorical(df_consolidado['Mês'], categories=ordem_meses, ordered=True)

    tabela_final = pd.pivot_table(
        df_consolidado,
        values='TOTAL',
        index='Linha',
        columns=['Mês', 'Quinzena']
    )
    
    tabela_final = tabela_final.sort_index(axis=1, level=['Mês', 'Quinzena'])

    nome_arquivo_saida = 'Relatorio_Passageiros_BRT_Final.xlsx'
    tabela_final.to_excel(nome_arquivo_saida)

    print(f"\nProcesso concluído! O relatório final foi salvo como '{nome_arquivo_saida}'.")

if __name__ == "__main__":
    gerar_relatorio_brt_final_robusto()