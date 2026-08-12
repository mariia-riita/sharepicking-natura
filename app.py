import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import json
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

# =========================================================================
# SEGURANÇA: Configuração da Chave API oculta nos segredos do Streamlit
# =========================================================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=API_KEY)
    chave_configurada = True
except Exception:
    chave_configurada = False

# Caixa de upload limpa: APENAS planilhas Excel e CSV
arquivos_carregados = st.file_uploader(
    "Arraste as cotações aqui (Formatos aceitos: Excel .xlsx ou .csv):", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True
)

# Limpeza de segurança para garantir o recebimento de um JSON puro da IA
def limpar_json_retornado(texto):
    texto = texto.strip()
    if texto.startswith("```"):
        linhas = texto.split("\n")
        if lines := [l for l in linhas if not l.startswith("```")]:
            texto = "\n".join(lines).strip()
    return texto

# Extração e conversão de strings de dinheiro para float puro
def limpar_valor(val):
    if pd.isna(val) or val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            return float(val)
        texto = str(val).replace("R$", "").replace(" ", "")
        if "." in texto and "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        return float(texto.strip())
    except ValueError:
        return None

# =========================================================================
# FUNÇÃO DE FORMATAÇÃO DO EXCEL (Design Poppins + Pulo de Linha por Região)
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
        
        # Salta uma linha física no Excel ao mudar de região
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
# OPERAÇÃO DA INTERFACE
# =========================================================================
if arquivos_carregados:
    if not chave_configurada:
        st.error("❌ Chave API não configurada no secrets.toml!")
    else:
        fornecedores_detectados = []
        contexto_planilhas = ""
        
        for arquivo in arquivos_carregados:
            nome_fornecedor = arquivo.name.split(".")[0].replace("Cotações M.O.xlsx -", "").replace("Cotações M.O. -", "").strip()
            fornecedores_detectados.append(nome_fornecedor)
            
            try:
                if arquivo.name.endswith(".xlsx"):
                    df_temp = pd.read_excel(arquivo)
                else:
                    df_temp = pd.read_csv(arquivo)
                
                contexto_planilhas += f"\n--- PROPOSTA DO FORNECEDOR: {nome_fornecedor} ---\n"
                contexto_planilhas += df_temp.to_csv(index=False) + "\n"
            except Exception as e:
                st.error(f"Erro ao ler {arquivo.name}: {e}")
            
        st.success(f"🤖 Agente: {len(arquivos_carregados)} fornecedores prontos: {', '.join(fornecedores_detectados)}")

        if st.button("🚀 Gerar Cherry Picking Mestre"):
            with st.spinner("🧠 IA unificando as planilhas e aplicando o padrão Natura..."):
                
                # Configuração segura com o modelo estável gemini-1.5-pro
                config_segura_json = {"temperature": 0.1, "response_mime_type": "application/json"}
                model = genai.GenerativeModel(model_name="gemini-1.5-pro")
                
                fornecedores_str = ", ".join([f'"{f}"' for f in fornecedores_detectados])
                
                prompt_unico = f"""
                Você é um analista especialista em suprimentos da Natura.
                Sua tarefa é consolidar os dados das planilhas de TODOS os fornecedores fornecidos abaixo em UMA ÚNICA TABELA CONSOLIDADA DE CHERRY PICKING.

                Fornecedores a incluir como colunas de valores: {fornecedores_str}

                Conteúdo das propostas:
                {contexto_planilhas}

                Regras Rígidas de Consolidação:
                1. Alinhe exatamente as mesmas linhas cruzando: Região, Cargo e Turnos.
                2. Padronize os nomes de "Região", "Cargo" e "Turnos" de forma idêntica para todos os fornecedores (ex: use o formato completo de turnos como "1º turno (Segunda à Sábado) - 06:00 - 14:00").
                3. Para cada linha, crie campos específicos para os valores numéricos das diárias de CADA fornecedor listado.
                4. O valor dos preços das diárias devem ser numéricos puramente (Ex: 229.82). Se algum fornecedor não tiver cotação para uma linha, envie null.

                Retorne APENAS uma lista de objetos JSON onde cada objeto representa uma linha com as chaves:
                "Região", "Cargo", "Turnos", {fornecedores_str}
                """
                
                try:
                    resposta = model.generate_content(prompt_unico, generation_config=config_segura_json)
                    texto_json = limpar_json_retornado(resposta.text)
                    
                    dados_json = json.loads(texto_json)
                    df_consolidado = pd.DataFrame(dados_json)
                    
                    # Organiza por Região e Turnos para garantir o layout em blocos
                    if "Região" in df_consolidado.columns:
                        df_consolidado = df_consolidado.sort_values(by=["Região", "Turnos"]).reset_index(drop=True)
                    
                    buffer_excel = estilizar_planilha_excel(df_consolidado, fornecedores_detectados)
                    
                    st.balloons()
                    st.success("✨ Processo concluído com Sucesso!")
                    
                    st.write("📊 Prévia da Tabela Consolidada:")
                    st.dataframe(df_consolidado, use_container_width=True)
                    
                    st.download_button(
                        label="📥 Clique aqui para baixar a Planilha Excel (.xlsx)",
                        data=buffer_excel,
                        file_name="Cherry_Picking_Consolidado_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Erro ao consolidar propostas: {e}")
