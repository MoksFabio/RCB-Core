import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def processar_arquivo():
    root = tk.Tk()
    root.withdraw()

    # 1. Selecionar arquivo
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo CSV",
        filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
    )

    if not file_path:
        return

    try:
        # 2. Leitura
        try:
            df = pd.read_csv(file_path, sep=';', encoding='utf-8-sig')
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, sep=';', encoding='latin1')

        # Limpeza de nomes de colunas
        df.columns = df.columns.str.strip()

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
            messagebox.showerror("Erro", f"Colunas faltando: {missing_cols}")
            return

        df = df[colunas_desejadas].copy()

        # 3. FILTRAGEM DE OPERADORAS
        operadoras_validas = ['BOA', 'CAX', 'CSR', 'EME', 'GLO', 'SJT', 'VML']
        df = df[df['Operadora'].isin(operadoras_validas)]

        if df.empty:
            messagebox.showwarning("Aviso", "Nenhuma linha encontrada com as operadoras selecionadas.")
            return

        # 4. LIMPEZA DE LINHA 0 (NOVO PASSO)
        # Converte para número para identificar o 0 com precisão
        df['Linha_Sort'] = pd.to_numeric(df['Linha'], errors='coerce')
        
        # --- AQUI ESTÁ A CORREÇÃO ---
        # Remove linhas onde a Linha é 0 ou nula (NaN)
        df = df[df['Linha_Sort'] != 0]
        df = df.dropna(subset=['Linha_Sort']) 
        
        # Preenche letras originais (caso existam linhas como 11A) para ordenação
        df['Linha_Sort'] = df['Linha_Sort'].fillna(df['Linha'])

        # 5. ORDENAÇÃO
        # Ordena Operadora (ordem fixa) -> Linha (número) -> Hora
        df['Operadora'] = pd.Categorical(df['Operadora'], categories=operadoras_validas, ordered=True)
        
        # Conversão de datas
        for col in [coluna_inicio, coluna_fim, coluna_retorno]:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

        df.dropna(subset=[coluna_inicio, coluna_fim], inplace=True)

        # Ordenação final
        df = df.sort_values(by=['Operadora', 'Linha_Sort', coluna_inicio], ascending=[True, True, True])

        # 6. CÁLCULOS
        df['Duracao_Segundos'] = (df[coluna_fim] - df[coluna_inicio]).dt.total_seconds()
        df['Hora_Inicio'] = df[coluna_inicio].dt.hour
        
        # Média agrupada por LINHA e HORA
        df['Media_Hora_Linha'] = df.groupby(['Linha', 'Hora_Inicio'])['Duracao_Segundos'].transform('mean')

        def definir_valor(row):
            if pd.isna(row['Media_Hora_Linha']) or row['Media_Hora_Linha'] == 0:
                return 1.0
            
            limite = row['Media_Hora_Linha'] * 0.30
            
            if row['Duracao_Segundos'] < limite:
                return 0.5
            return 1.0

        df['Valor'] = df.apply(definir_valor, axis=1)

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

            df_final.to_csv(save_path, index=False, sep=';', float_format='%.1f', encoding='utf-8-sig')
            messagebox.showinfo("Sucesso", "Linha 0 removida e arquivo gerado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Detalhe do erro: {str(e)}")

if __name__ == "__main__":
    processar_arquivo()