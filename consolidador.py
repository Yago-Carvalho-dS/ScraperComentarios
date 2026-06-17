import pandas as pd
import os
import glob
import sys

def process_excel_files(directory, output_file="Consolidado_Comentarios.xlsx"):
    """
    Lê os arquivos Excel baixados, identifica colunas de comentários e consolida em um único arquivo.
    """
    print(f"\n--- Iniciando Consolidação em: {directory} ---")
    
    # Busca todos os arquivos .xlsx na pasta
    files = glob.glob(os.path.join(directory, "*.xlsx"))
    
    # Ignora o arquivo de modelo e o próprio arquivo de saída se ele já existir na pasta
    files = [f for f in files if "Teste.xlsx" not in f and os.path.basename(f) != output_file]
    
    if not files:
        print(f"[Erro] Nenhum arquivo .xlsx encontrado na pasta: {directory}")
        return

    print(f"Encontrados {len(files)} arquivos para processar.")
    
    all_rows = []
    
    # Palavras-chave para identificar colunas de comentários
    keywords = ['comente', 'comentário', 'comentario', 'describe', 'descreva', 'ponto', 'fraco', 'forte']
    
    for file in files:
        try:
            # Carrega o Excel
            df = pd.read_excel(file)
            if df.empty or len(df) < 1:
                continue
                
            # Identifica colunas de comentário verificando o cabeçalho e a primeira linha (texto da pergunta)
            question_texts = df.iloc[0]
            comment_cols = []
            
            for col in df.columns:
                col_name_lower = str(col).lower()
                first_row_val_lower = str(question_texts[col]).lower()
                
                is_comment = any(kw in col_name_lower or kw in first_row_val_lower for kw in keywords)
                
                if is_comment:
                    comment_cols.append(col)
            
            if comment_cols:
                # Tenta identificar onde começam os dados reais
                # Row 0=IDs (headers), Row 1=Text (index 0), Row 2=ImportID (index 1) OU Data (index 1)
                data_start_idx = 1
                if len(df) > 1:
                    # Verifica se a linha index 1 parece ser metadados (ImportID)
                    val_check = str(df.iloc[1, 0])
                    if '{"importid"' in val_check.lower():
                        data_start_idx = 2
                
                data_df = df.iloc[data_start_idx:].copy()
                
                # Pega o nome do formulário (primeira parte do nome do arquivo)
                form_name = os.path.basename(file).split("_")[0]
                
                rows_with_comments = 0
                for _, row in data_df.iterrows():
                    new_entry = {"Formulário": form_name}
                    found_comment = False
                    for i, col in enumerate(comment_cols):
                        val = row[col]
                        if pd.notna(val) and str(val).strip() != "":
                            new_entry[f"Comentario{i+1}"] = val
                            found_comment = True
                    
                    if found_comment:
                        all_rows.append(new_entry)
                        rows_with_comments += 1
                
                if rows_with_comments > 0:
                    print(f"  [OK] {os.path.basename(file)}: {rows_with_comments} comentários encontrados.")
                else:
                    print(f"  [-] {os.path.basename(file)}: Nenhum comentário preenchido.")
                    
        except Exception as e:
            print(f"  [Erro] Falha ao processar {os.path.basename(file)}: {e}")

    if all_rows:
        final_df = pd.DataFrame(all_rows)
        # Garante a ordem das colunas: Formulário, Comentario1, Comentario2...
        cols = ["Formulário"] + [c for c in final_df.columns if c.startswith("Comentario")]
        final_df = final_df[cols]
        
        final_df.to_excel(output_file, index=False)
        print(f"\n==========================================")
        print(f"[SUCESSO] Arquivo gerado: {output_file}")
        print(f"Total de comentários consolidados: {len(all_rows)}")
        print(f"==========================================")
    else:
        print("\n[Aviso] Nenhum comentário foi encontrado em nenhum dos arquivos.")

if __name__ == "__main__":
    # Se passar um caminho via linha de comando, usa ele. Senão usa a pasta atual.
    target_folder = sys.argv[1] if len(sys.argv) > 1 else "."
    
    # Se a pasta "ExcelComent" existir e não foi passado argumento, foca nela por padrão
    if len(sys.argv) == 1 and os.path.exists("ExcelComent"):
        target_folder = "ExcelComent"
        
    process_excel_files(target_folder)
    input("\nPressione ENTER para fechar...")
