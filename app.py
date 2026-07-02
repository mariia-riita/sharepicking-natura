import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import json
import openpyxl
import pypdf
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================================
# CONFIGURAÇÃO DE DESIGN DA PLATAFORMA (Padrão Natura & Co com emoji 💵)
# =========================================================================
st.set_page_config(page_title="Share Picking - Natura", page_icon="💵", layout="wide")

st.title("💵 Agente de Sourcing: Share Picking")
st.write("Suba as planilhas oficiais, CSVs ou propostas em PDF dos fornecedores para gerar a consolidação automática.")

# =========================================================================
# SEGURANÇA: Configuração da Chave API oculta nos segredos do Streamlit
# =========================================================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=API_KEY)
    chave_configurada = True
except Exception:
    chave_configurada = False

# Caixa de upload de arquivos (Aceita múltiplos formatos)
arquivos_carregados = st.file_uploader(
    "Arraste as cotações aqui (Formatos aceitos: Excel .xlsx, .csv ou propostas em .pdf):", 
    type=["xlsx", "csv", "pdf"], 
    accept_multiple_files=True
)

# Limpeza de segurança para garantir o recebimento de um JSON puro da IA
def limpar_json_retornado(texto):
    texto = texto.strip()
    if texto.startswith("```"):
        linhas = texto.split("\n")
        if linhas[0].startswith("```"):
            linhas = linhas[1:]
        if linhas[-1].startswith("```"):
            linhas = linhas[:-1]
        texto = "\n".join(linhas).strip()
    return texto

# Extração e conversão de strings de dinheiro para float puro
def limpar_valor(val):
    if pd.isna(val) or val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            return float(val)
        texto = str(val).replace("R$", "").replace(" ", "")
        # Lida com formatos europeus/brasileiros de milhar e decimal
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
    ws.title = "Share Picking"
    
    # Monta o cabeçalho dinâmico baseado nos arquivos reais processados
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
        
        # PULO DE LINHA ESTILIZADO: Se mudou de região, salta uma linha física no Excel
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
        
        # Injeção das fórmulas vivas que o Excel traduz nativamente para MÍN e PROCX
        ws.cell(row=row_num, column=idx_min_col, value=f"=MIN({col_let_forn_start}{row_num}:{col_let_forn_end}{row_num})")
        ws.cell(row=row_num, column=idx_winner_col, value=f'=_xlfn.XLOOKUP({col_let_min}{row_num}, {col_let_forn_start}{row_num}:{col_let_forn_end}{row_num}, ${col_let_forn_start}$1:${col_let_forn_end}$1)')
        
        valores_linha = {col_idx: limpar_valor(row.get(forn, None)) for col_idx, forn in enumerate(fornecedores, start=4)}
        valores_validos = {col: val for col, val in valores_linha.items() if val is not None}
        coluna_vencedora = min(valores_validos, key=valores_validos.get) if valores_validos else None
        
        winner_fill = PatternFill(start_color="E6F0EA", end_color="E6F0EA", fill_type="solid") # Soft Mint
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
# PROCESSAMENTO DE ARQUIVOS EM LOOP INDIVIDUAL
# =========================================================================
if arquivos_carregados:
    if not chave_configurada:
        st.error("❌ Chave API não configurada no secrets.toml!")
    else:
        dfs_fornecedores = []
        fornecedores_detectados = []
        
        # Loop individual por arquivo para garantir que a IA não misture os dados
        for arquivo in arquivos_carregados:
            nome_fornecedor = arquivo.name.split(".")[0]
            nome_fornecedor = nome_fornecedor.replace("Cotações M.O.xlsx -", "")
            nome_fornecedor = nome_fornecedor.replace("Cotações M.O. -", "")
            nome_fornecedor = nome_fornecedor.replace("Cópia de Proposta_Técnica___Comercial_", "")
            nome_fornecedor = nome_fornecedor.replace("Proposta Comercial ", "")
            nome_fornecedor = nome_fornecedor.replace("Prestador Serviço", "")
            nome_fornecedor = nome_fornecedor.strip()
            
            fornecedores_detectados.append(nome_fornecedor)
            
            # Extração de texto isolado
            if arquivo.name.endswith(".pdf"):
                pdf_reader = pypdf.PdfReader(io.BytesIO(arquivo.read()))
                conteudo_texto = ""
                for pagina in pdf_reader.pages:
                    texto_pagina = pagina.extract_text()
                    if texto_pagina: conteudo_texto += texto_pagina + "\n"
            elif arquivo.name.endswith(".xlsx"):
                df_temp = pd.read_excel(arquivo)
                conteudo_texto = df_temp.to_csv(index=False)
            else:
                df_temp = pd.read_csv(arquivo)
                conteudo_texto = df_temp.to_csv(index=False)
                
            st.write(f"🔍 Extraindo dados de: **{nome_fornecedor}**...")
            
            # Prompt cirúrgico focado em UM único fornecedor por vez
            prompt_individual = f"""
            Você é um analista de dados especialista em suprimentos. 
            Sua tarefa é extrair os valores de diárias de mão de obra para o fornecedor '{nome_fornecedor}'.
            
            Aqui está o conteúdo do arquivo enviado por ele:
            {conteudo_texto}
            
            Gere um objeto JSON contendo uma lista de dicionários com as chaves exatas:
            "Região", "Cargo", "Turnos", "{nome_fornecedor}"
            
            Regras de padronização:
            1. "Região" deve conter o nome do local limpo (Ex: "Murici / AL", "São Paulo/Cajamar", "Uberlândia / MG").
            2. "Cargo" deve ser mapeado como "Ajudante Picking".
            3. "Turnos" deve descrever o turno ou horário de trabalho de forma clara.
            4. "{nome_fornecedor}" deve conter o valor numérico puro da diária (Ex: 265.00). Não coloque texto ou R$.
            
            Retorne APENAS o JSON bruto, sem formatações markdown.
            """
            
            model = genai.GenerativeModel(model_name="gemini-3.1-flash-lite-preview")
            resposta = model.generate_content(prompt_individual, generation_config={"temperature": 0.1})
            
            texto_json = limpar_json_retornado(resposta.text)
            
            try:
                dados_json = json.loads(texto_json)
                df_individual = pd.DataFrame(dados_json)
                
                # Garante padronização básica de colunas
                df_individual.columns = [str(c).strip() for c in df_individual.columns]
                df_individual = df_individual.drop_duplicates(subset=["Região", "Cargo", "Turnos"])
                
                dfs_fornecedores.append(df_individual)
            except Exception as e:
                st.error(f"Erro ao interpretar dados de {nome_fornecedor}: {e}")

        # =========================================================================
        # CONSOLIDAÇÃO INTELIGENTE VIA PYTHON (OUTER JOIN)
        # =========================================================================
        if dfs_fornecedores:
            if st.button("🚀 Gerar Share Picking Mestre"):
                with st.spinner("🧠 Unindo planilhas e aplicando a identidade da Natura..."):
                    
                    # Usa o poder do Pandas para fazer o cruzamento perfeito das tabelas
                    df_consolidado = dfs_fornecedores[0]
                    for df_proximo in dfs_fornecedores[1:]:
                        df_consolidado = pd.merge(
                            df_consolidado, df_proximo, 
                            on=["Região", "Cargo", "Turnos"], 
                            how="outer"
                        )
                    
                    # Organiza as linhas por região para o layout em blocos ficar perfeito
                    df_consolidado = df_consolidado.sort_values(by=["Região", "Turnos"]).reset_index(drop=True)
                    
                    # Cria o arquivo final estilizado
                    buffer_excel = estilizar_planilha_excel(df_consolidado, fornecedores_detectados)
                    
                    st.balloons()
                    st.success("✨ Processo concluído! Os PDFs e Planilhas foram unificados.")
                    
                    st.write("📊 Prévia Consolidada:")
                    st.dataframe(df_consolidado, use_container_width=True)
                    
                    st.download_button(
                        label="📥 Clique aqui para baixar a Planilha Excel (.xlsx)",
                        data=buffer_excel,
                        file_name="Share_Picking_Consolidado_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )