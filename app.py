import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================================
# CONFIGURAÇÃO DA PÁGINA E DESIGN POPPINS
# =========================================================================
st.set_page_config(page_title="Cherry Picking - Natura", page_icon="💵", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');

    html, body, [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Poppins', sans-serif !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("💵 Agente de Sourcing: Cherry Picking")
st.write("Suba as planilhas oficiais ou CSVs dos fornecedores para gerar a consolidação automática.")

# Caixa de upload: Apenas planilhas Excel e CSV
arquivos_carregados = st.file_uploader(
    "Arraste as cotações aqui (Formatos aceitos: Excel .xlsx ou .csv):", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True
)

def limpar_valor(val):
    if pd.isna(val) or val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            return float(val)
        texto = str(val).replace("R$", "").replace(" ", "").strip()
        if "." in texto and "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        return float(texto)
    except ValueError:
        return None

# =========================================================================
# PROCESSAMENTO INTELIGENTE DE MULTI-ABAS E CUSTO TOTAL
# =========================================================================
def processar_arquivo_cotacao(file_obj, nome_arquivo):
    dfs_extraidos = []
    
    if nome_arquivo.endswith(".csv"):
        df_raw = pd.read_csv(file_obj)
        dfs_extraidos.append((nome_arquivo, df_raw))
    else:
        xl = pd.ExcelFile(file_obj)
        for sheet in xl.sheet_names:
            df_preview = pd.read_excel(file_obj, sheet_name=sheet, header=None)
            
            # Localiza a linha exata onde estão os cabeçalhos reais (Região, Turnos, etc.)
            header_row = 0
            for r_idx in range(min(6, len(df_preview))):
                vals = [str(v).strip().lower() for v in df_preview.iloc[r_idx].values if pd.notna(v)]
                if any("regi" in v for v in vals) and any("turn" in v for v in vals):
                    header_row = r_idx
                    break
                    
            df_sheet = pd.read_excel(file_obj, sheet_name=sheet, header=header_row)
            
            # Rotula o fornecedor/modalidade considerando o nome da aba se houver mais de uma
            label = nome_arquivo
            if len(xl.sheet_names) > 1:
                aba_limpa = sheet.replace("Planilha de Cotação", "").replace("Cotação", "").strip()
                label = f"{nome_arquivo} ({aba_limpa})" if aba_limpa else f"{nome_arquivo} ({sheet})"
                
            dfs_extraidos.append((label, df_sheet))

    dfs_normalizados = []
    for label, df in dfs_extraidos:
        col_reg, col_cargo, col_turno, col_valor = None, None, None, None
        
        for col in df.columns:
            c_lower = str(col).strip().lower()
            if any(k in c_lower for k in ["regi", "cd", "local"]):
                col_reg = col
            elif any(k in c_lower for k in ["carg", "funç", "func"]):
                col_cargo = col
            elif any(k in c_lower for k in ["turn", "horar"]):
                col_turno = col

        # Prioridade máxima para "Valor Total" para capturar o custo final com taxas/impostos
        for col in df.columns:
            c_lower = str(col).strip().lower()
            if "valor total" in c_lower or "total" in c_lower:
                col_valor = col
                break
                
        if not col_valor:
            for col in df.columns:
                c_lower = str(col).strip().lower()
                if any(g in c_lower for g in ["diaria", "diária", "preço", "preco", "valor"]) and not any(b in c_lower for bad in ["taxa", "imposto", "%", "pis", "iss", "ir"] for b in [bad]):
                    col_valor = col
                    break

        if col_reg and col_turno and col_valor:
            df_sub = pd.DataFrame()
            df_sub["Região"] = df[col_reg].astype(str).str.strip()
            df_sub["Cargo"] = df[col_cargo].astype(str).str.strip() if col_cargo else "Ajudante Picking"
            df_sub["Turnos"] = df[col_turno].astype(str).str.strip()
            df_sub[label] = df[col_valor].apply(limpar_valor)
            
            # Limpeza de linhas vazias e cabeçalhos residuais
            df_sub = df_sub.dropna(subset=["Região", label])
            df_sub = df_sub[~df_sub["Região"].str.lower().isin(["nan", "região", "region"])]
            dfs_normalizados.append(df_sub)
            
    return dfs_normalizados

# =========================================================================
# FUNÇÃO DE FORMATAÇÃO E EXPORTAÇÃO EXCEL
# =========================================================================
def estilizar_planilha_excel(df, fornecedores):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cherry Picking"
    
    colunas_completas = ["Região", "Cargo", "Turnos"] + fornecedores + ["Melhor Preço", "Fornecedor Vencedor"]
    ws.append(colunas_completas)
    
    idx_min_col = 4 + len(fornecedores)
    idx_winner_col = 5 + len(fornecedores)
    total_cols = idx_winner_col
    
    header_fill = PatternFill(start_color="9E472A", end_color="9E472A", fill_type="solid") # Terracota Natura
    header_font = Font(name="Poppins", size=11, bold=True, color="FFFFFF")
    border_cinza = Border(
        left=Side(style='thin', color='E0D8D3'), right=Side(style='thin', color='E0D8D3'),
        top=Side(style='thin', color='E0D8D3'), bottom=Side(style='thin', color='E0D8D3')
    )
    
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for col_idx in range(1, total_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border_cinza

    row_num = 2
    current_region = None
    color_index = 0
    fallback_colors = ["FDF2EE", "FAF4F0", "EBF2EE", "FDF5EC", "F5F2EB", "FCF9F2", "FDF2F4", "F4F5F6"]
    assigned_colors = {}

    for _, row in df.iterrows():
        regiao_original = row.get("Região", "")
        regiao = str(regiao_original).strip()
        
        if current_region is not None and regiao != current_region:
            row_num += 1
            
        current_region = regiao
        
        if regiao not in assigned_colors:
            assigned_colors[regiao] = fallback_colors[color_index % len(fallback_colors)]
            color_index += 1
            
        cor_fundo_regiao = assigned_colors[regiao]
        fill_linha = PatternFill(start_color=cor_fundo_regiao, end_color=cor_fundo_regiao, fill_type="solid")
        
        ws.cell(row=row_num, column=1, value=regiao_original)
        ws.cell(row=row_num, column=2, value=row.get("Cargo", ""))
        ws.cell(row=row_num, column=3, value=row.get("Turnos", ""))
        
        for col_idx, forn in enumerate(fornecedores, start=4):
            valor = limpar_valor(row.get(forn, None))
            ws.cell(row=row_num, column=col_idx, value=valor)
        
        col_let_forn_start = get_column_letter(4)
        col_let_forn_end = get_column_letter(3 + len(fornecedores))
        col_let_min = get_column_letter(idx_min_col)
        
        # Fórmulas de Excel gravadas para execução no programa desktop
        ws.cell(row=row_num, column=idx_min_col, value=f"=MIN({col_let_forn_start}{row_num}:{col_let_forn_end}{row_num})")
        ws.cell(row=row_num, column=idx_winner_col, value=f'=_xlfn.XLOOKUP({col_let_min}{row_num}, {col_let_forn_start}{row_num}:{col_let_forn_end}{row_num}, ${col_let_forn_start}$1:${col_let_forn_end}$1)')
        
        valores_linha = {col_idx: limpar_valor(row.get(forn, None)) for col_idx, forn in enumerate(fornecedores, start=4)}
        valores_validos = {col: val for col, val in valores_linha.items() if val is not None}
        coluna_vencedora = min(valores_validos, key=valores_validos.get) if valores_validos else None
        
        winner_fill = PatternFill(start_color="E6F0EA", end_color="E6F0EA", fill_type="solid")
        winner_font = Font(name="Poppins", size=10, bold=True, color="1E3D2F")
        normal_font = Font(name="Poppins", size=10, color="000000")
        
        for col_idx in range(1, total_cols + 1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.border = border_cinza
            cell.font = normal_font
            
            if col_idx == coluna_vencedora:
                cell.fill = winner_fill
                cell.font = winner_font
            else:
                cell.fill = fill_linha
                
            if col_idx in [1, 2, 3]:
                cell.alignment = left_align
            else:
                cell.alignment = center_align
                
            if 4 <= col_idx <= idx_min_col:
                cell.number_format = 'R$ #,##0.00'
                
        row_num += 1

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    ws.views.sheetView[0].showGridLines = True
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================================
# EXECUÇÃO DA INTERFACE
# =========================================================================
if arquivos_carregados:
    dfs_processados = []

    for arquivo in arquivos_carregados:
        nome_clean = arquivo.name.split(".")[0].replace("Cotações M.O.xlsx -", "").replace("Cotações M.O. -", "").strip()
        try:
            sub_dfs = processar_arquivo_cotacao(arquivo, nome_clean)
            dfs_processados.extend(sub_dfs)
        except Exception as e:
            st.error(f"Erro ao ler arquivo {arquivo.name}: {e}")

    if dfs_processados:
        colunas_fornecedores = [list(d.columns)[-1] for d in dfs_processados]
        st.success(f"🤖 Agente: {len(colunas_fornecedores)} coluna(s) de cotação identificada(s): {', '.join(colunas_fornecedores)}")

        if st.button("🚀 Gerar Cherry Picking Mestre"):
            with st.spinner("⚡ Unificando planilhas e calculando os menores preços..."):
                try:
                    df_consolidado = dfs_processados[0]
                    for df_prox in dfs_processados[1:]:
                        df_consolidado = pd.merge(df_consolidado, df_prox, on=["Região", "Cargo", "Turnos"], how="outer")

                    df_consolidado = df_consolidado.sort_values(by=["Região", "Turnos"]).reset_index(drop=True)

                    # Cálculo explícito em Python para que a prévia da tela mostre os resultados reais
                    def calc_melhor_preco(row):
                        vals = [limpar_valor(row[c]) for c in colunas_fornecedores if pd.notna(row[c])]
                        vals_validos = [v for v in vals if v is not None]
                        return min(vals_validos) if vals_validos else None

                    def calc_fornecedor_vencedor(row):
                        best = calc_melhor_preco(row)
                        if best is None:
                            return None
                        for c in colunas_fornecedores:
                            if limpar_valor(row[c]) == best:
                                return c
                        return None

                    df_consolidado["Melhor Preço"] = df_consolidado.apply(calc_melhor_preco, axis=1)
                    df_consolidado["Fornecedor Vencedor"] = df_consolidado.apply(calc_fornecedor_vencedor, axis=1)

                    buffer_excel = estilizar_planilha_excel(df_consolidado, colunas_fornecedores)

                    st.balloons()
                    st.success("✨ Processo concluído com sucesso!")

                    st.write("📊 Prévia da Tabela Consolidada:")
                    st.dataframe(df_consolidado, use_container_width=True)

                    st.download_button(
                        label="📥 Clique aqui para baixar a Planilha Excel (.xlsx)",
                        data=buffer_excel,
                        file_name="Cherry_Picking_Consolidado_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ Erro ao consolidar planilhas: {e}")
