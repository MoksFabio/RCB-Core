import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

import csv
from simpledbf import Dbf5

def processar_arquivo():
    root = tk.Tk()
    root.withdraw()

    # 1. Selecionar arquivo
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo CSV ou DBF",
        filetypes=[("Arquivos de Dados", "*.csv;*.dbf"), ("Arquivos CSV", "*.csv"), ("Arquivos DBF", "*.dbf"), ("Todos os arquivos", "*.*")]
    )

    if not file_path:
        return

    try:
        # 2. Leitura
        _, ext = os.path.splitext(file_path)
        if ext.lower() == '.dbf':
            try:
                dbf = Dbf5(file_path, codec='latin-1')
                df = dbf.to_dataframe()
            except Exception as e:
                # Tenta codec utf-8 se latin-1 falhar, ou sobe erro
                try:
                    dbf = Dbf5(file_path, codec='utf-8')
                    df = dbf.to_dataframe()
                except:
                   raise Exception(f"Erro ao ler arquivo DBF: {e}")
        else:
            try:
                df = pd.read_csv(file_path, sep=';', encoding='utf-8-sig', low_memory=False)
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, sep=';', encoding='latin1', low_memory=False)

        # Limpeza de nomes de colunas
        df.columns = df.columns.str.strip()

        # --- CORREÇÃO: FILTRAR APENAS VIAGENS NORMAIS (NOR) ---
        # Isso remove viagens de garagem/manutenção que o sistema antigo não mostra
        if 'Tipo Viagem' in df.columns:
            df = df[df['Tipo Viagem'] == 'NOR']
        
        
        # Mapping para nomes curtos (DBF) se necessário
        mapa_colunas = {
            'OPERADORA': 'Operadora',
            'DATA': 'Data Coleta',
            'LINHA': 'Linha',
            'INI_VIAG': 'DataHora Inicio Operação',  # Exemplo hipotético, ajuste conforme real
            'FIM_VIAG': 'DataHora Final Operação',   # Exemplo hipotético, ajuste conforme real
            'PTO_RET': 'Data Hora Ponto Retorno'     # Exemplo hipotético, ajuste conforme real
        }
        # Tenta renomear se encontrar as chaves
        df = df.rename(columns=mapa_colunas)

        # Definição das colunas
        coluna_inicio = 'DataHora Inicio Operação' 
        coluna_fim = 'DataHora Final Operação'
        coluna_retorno = 'Data Hora Ponto Retorno'
        
        colunas_desejadas = [
            'Operadora', 
            'Data Coleta', 
            'Linha', 
            coluna_inicio, 
            coluna_fim, 
            coluna_retorno
        ]
        
        # Verifica colunas
        missing_cols = [c for c in colunas_desejadas if c not in df.columns]
        if missing_cols:
            colunas_encontradas = list(df.columns)
            msg_erro = f"Colunas faltando: {missing_cols}\n\nColunas encontradas no arquivo:\n{colunas_encontradas}"
            # Tenta ser prestativo e sugere mapeamento se for DBF
            if any(len(c) <= 10 for c in colunas_encontradas):
                msg_erro += "\n\nDica: Arquivos DBF costumam ter nomes de colunas abreviados (max 10 chars). Verifique o mapeamento."
            
            messagebox.showerror("Erro de Colunas", msg_erro)
            return

        df = df[colunas_desejadas].copy()

        # Remove duplicatas exatas
        df = df.drop_duplicates(subset=['Operadora', 'Linha', coluna_inicio], keep='first')

        # 3. FILTRAGEM DE OPERADORAS
        operadoras_validas = ['BOA', 'CAX', 'CSR', 'EME', 'GLO', 'SJT', 'VML']
        df = df[df['Operadora'].isin(operadoras_validas)]

        if df.empty:
            messagebox.showwarning("Aviso", "Nenhuma linha encontrada com as operadoras selecionadas.")
            return

        # 4. LIMPEZA DE LINHA 0
        df['Linha_Sort'] = pd.to_numeric(df['Linha'], errors='coerce')
        df = df[df['Linha_Sort'] != 0] 
        df = df.dropna(subset=['Linha_Sort'])
        df['Linha_Sort'] = df['Linha_Sort'].fillna(df['Linha'])

        # 5. ORDENAÇÃO
        df['Operadora'] = pd.Categorical(df['Operadora'], categories=operadoras_validas, ordered=True)
        
        for col in [coluna_inicio, coluna_fim, coluna_retorno]:
            # Alteração: Removido dayfirst=True pois o usuário relatou que as datas estão invertidas (Mês/Dia).
            # Ao remover, o pandas tentará inferir ou usar o padrão (Mês primeiro).
            df[col] = pd.to_datetime(df[col], errors='coerce')

        df.dropna(subset=[coluna_inicio, coluna_fim], inplace=True)

        df = df.sort_values(by=['Operadora', 'Linha_Sort', coluna_inicio], ascending=[True, True, True])

        # 6. CÁLCULOS
        df['Duracao_Segundos'] = (df[coluna_fim] - df[coluna_inicio]).dt.total_seconds()
        df['Hora_Inicio'] = df[coluna_inicio].dt.hour
        
        # Média agrupada por LINHA e HORA
        df['Media_Hora_Linha'] = df.groupby(['Linha', 'Hora_Inicio'])['Duracao_Segundos'].transform('mean')

        # Função auxiliar para formatar tempo (segundos) em HH:MM:SS
        def formatar_segundos(total_seconds):
             if pd.isna(total_seconds): return "00:00:00"
             total_seconds = int(total_seconds)
             hours = total_seconds // 3600
             minutes = (total_seconds % 3600) // 60
             seconds = total_seconds % 60
             return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

        def processar_linha(row):
            duracao = row['Duracao_Segundos']
            media = row['Media_Hora_Linha']
            
            # Defaults
            valor = 1.0
            cod_erro = ""
            aviso = ""
            
            if pd.isna(media) or media == 0:
                pass
            else:
                # Regra 1: Se Tempo Viagem TOTAL < 30% de (Tempo Viagem Médio / 2)
                meia_viagem_segundos = media / 2
                limite_total = meia_viagem_segundos * 0.30
                
                if duracao < limite_total:
                    valor = 0.5
                    meia_viagem_str = formatar_segundos(meia_viagem_segundos)
                    cod_erro = f"TP VIAG <30% TP MED/2 ({meia_viagem_str})"
                    aviso = "MEIA VIAGEM"
                
                # Regra 2: Se existe Ponto de Retorno, verifica o trecho Retorno -> Fim
                # Se (Fim - Retorno) < 30% de (Tempo Viagem Médio / 2) -> Meia Viagem
                elif not pd.isna(row['Data Hora Ponto Retorno']):
                    try:
                        dt_fim = pd.to_datetime(row['DataHora Final Operação'])
                        dt_ret = pd.to_datetime(row['Data Hora Ponto Retorno'])
                        
                        if not pd.isna(dt_fim) and not pd.isna(dt_ret):
                            duracao_segunda_perna = (dt_fim - dt_ret).total_seconds()
                            # Limite da Perna: 30% da Meia Viagem Média
                            limite_perna = meia_viagem_segundos * 0.30
                            
                            if duracao_segunda_perna < limite_perna:
                                valor = 0.5
                                cod_erro = "RETORNO < 30% MEIA VIAGEM"
                                aviso = "MEIA VIAGEM"
                    except:
                        pass # Erro de data, ignora


            # Tempo Viagem Formatado
            tempo_viagem_str = formatar_segundos(duracao)
            
            return pd.Series([valor, cod_erro, aviso, tempo_viagem_str])

        df[['Valor', 'Código de Erro', 'Aviso', 'Tempo Viagem']] = df.apply(processar_linha, axis=1)


        # 7. LIMPEZA FINAL
        df_final = df.drop(columns=['Duracao_Segundos', 'Hora_Inicio', 'Media_Hora_Linha', 'Linha_Sort'])
        df_final = df_final.rename(columns={'Operadora': 'Operador'})

        # 8. SALVAR
        save_path = filedialog.asksaveasfilename(
            title="Salvar arquivo processado",
            defaultextension=".csv",
            filetypes=[("Arquivos CSV", "*.csv")]
        )

        if save_path:
            for col in [coluna_inicio, coluna_fim, coluna_retorno]:
                 if col in df_final.columns:
                    df_final[col] = df_final[col].dt.strftime('%d/%m/%Y %H:%M:%S')

            # Salva com decimal vírgula e ASPAS EM TUDO (QUOTE_ALL)
            # Isso garante que o Excel entenda "0,5" como um único valor, mesmo que ele ache que a vírgula é separador
            try:
                df_final.to_csv(
                    save_path, 
                    index=False, 
                    sep=';', 
                    float_format='%.1f', 
                    encoding='utf-8-sig', 
                    decimal=',', 
                    quoting=csv.QUOTE_ALL
                )
            except Exception as e_save:
                messagebox.showerror("Erro", f"Erro ao salvar arquivo:\n{e_save}")
                return
            
            # Mensagem com contagem final
            messagebox.showinfo("Sucesso", f"Processamento concluído!\nLinhas geradas: {len(df_final)}\n(Filtro 'Tipo Viagem=NOR' aplicado)")

    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Detalhe do erro: {str(e)}")

if __name__ == "__main__":
    processar_arquivo()