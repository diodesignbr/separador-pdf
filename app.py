import os
import re
import unicodedata
import zipfile
import io
import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject
import pymupdf

# Configuração da página
st.set_page_config(page_title="Separador de Aulas PDF",
                   page_icon="📚", layout="wide")

st.title("📚 Separador Inteligente de Aulas")
st.markdown("Carregue o sumário (`.txt`), o PDF da apostila e, opcionalmente, a planilha de nomenclatura (`.xlsx`).")

# ==========================================
# FUNÇÕES AUXILIARES E NORMALIZAÇÃO
# ==========================================


def limpar_nome_arquivo(nome):
    nome_limpo = re.sub(r'[\\/*?:"<>|]', '_', str(nome)).strip()
    return re.sub(r'^_+', '', nome_limpo)


def remover_acentos(txt):
    return ''.join(c for c in unicodedata.normalize('NFD', str(txt)) if unicodedata.category(c) != 'Mn')


def normalizar_texto_base(txt):
    return remover_acentos(str(txt)).lower().strip()


def obter_sigla_disciplina(materia_txt):
    m_norm = normalizar_texto_base(materia_txt)
    if "gramatica" in m_norm:
        return "GRA"
    if "interpretacao" in m_norm or "inter" in m_norm:
        return "ITX"
    if "literatura" in m_norm:
        return "LIT"
    if "lingua portuguesa" in m_norm or "portuguesa" in m_norm:
        return "POR"
    if "ingles" in m_norm:
        return "ING"
    if "espanhol" in m_norm:
        return "ESP"
    if "matematica" in m_norm:
        return "MAT"
    if "historia" in m_norm:
        return "HIS"
    if "geografia" in m_norm:
        return "GEO"
    if "biologia" in m_norm:
        return "BIO"
    if "fisica" in m_norm:
        return "FIS"
    if "quimica" in m_norm:
        return "QUI"
    return "GER"


def normalizar_subcomponente(materia_txt):
    m_norm = normalizar_texto_base(materia_txt)
    if "gramatica" in m_norm:
        return "lingua portuguesa [gramatica]"
    if "interpretacao" in m_norm or "inter" in m_norm:
        return "lingua portuguesa [inter. de texto]"
    if "literatura" in m_norm:
        return "literatura"
    if "ingles" in m_norm:
        return "lingua inglesa"
    if "espanhol" in m_norm:
        return "lingua espanhola"
    if "matematica" in m_norm:
        return "matematica"
    if "historia" in m_norm:
        return "historia"
    if "geografia" in m_norm:
        return "geografia"
    if "biologia" in m_norm:
        return "biologia"
    if "fisica" in m_norm:
        return "fisica"
    if "quimica" in m_norm:
        return "quimica"
    return m_norm


def extrair_numero_frente(frente_txt):
    if not frente_txt:
        return "u"
    f_norm = normalizar_texto_base(frente_txt)
    if "unica" in f_norm or f_norm == "u" or not f_norm:
        return "u"
    m = re.search(r'\d+', f_norm)
    return m.group(0) if m else "u"


def parsear_nome_simplificado_para_busca(item):
    sub_alvo = normalizar_subcomponente(item['disciplina'])
    frente_alvo = extrair_numero_frente(item['frente']).lower()

    nums = re.findall(r'\d+', item['rotulo'])
    if not nums:
        cap_alvo = [1]
    elif len(nums) == 1:
        cap_alvo = [int(nums[0])]
    else:
        rotulo_lower = item['rotulo'].lower()
        if " a " in rotulo_lower or " à " in rotulo_lower:
            try:
                inicio, fim = int(nums[0]), int(nums[1])
                cap_alvo = list(range(inicio, fim + 1))
            except (ValueError, TypeError):
                cap_alvo = [int(n) for n in nums]
        else:
            cap_alvo = [int(n) for n in nums]

    if item['tipo'] == 'Gabarito':
        rotulo_limpo = re.sub(r'^gabarito\s*', '',
                              item['rotulo'], flags=re.IGNORECASE).strip()
        sub_alvo = normalizar_subcomponente(
            rotulo_limpo if rotulo_limpo else item['disciplina'])
        return sub_alvo, "u", [-1]
    elif item['tipo'] == 'Iniciais':
        return sub_alvo, "u", [0]

    return sub_alvo, frente_alvo, cap_alvo


def extrair_numeros_planilha(val):
    if pd.isna(val):
        return []
    val_str = str(val).strip().lower()
    nums = re.findall(r'\d+', val_str)
    if not nums:
        return []
    if 'a' in val_str or '-' in val_str:
        try:
            inicio, fim = int(nums[0]), int(nums[-1])
            return list(range(inicio, fim + 1))
        except (ValueError, TypeError):
            return [int(n) for n in nums]
    return [int(n) for n in nums]


def rect_to_tuple(rect):
    return (float(rect.left), float(rect.bottom), float(rect.right), float(rect.top))


def pick_crop_source(page):
    if getattr(page, "trimbox", None):
        tb = page.trimbox
        if float(tb.width) > 0 and float(tb.height) > 0:
            return tb
    if getattr(page, "bleedbox", None):
        bb = page.bleedbox
        if float(bb.width) > 0 and float(bb.height) > 0:
            return bb
    return page.cropbox


def aplicar_crop_pdf(caminho_entrada, caminho_saida):
    reader = PdfReader(caminho_entrada)
    writer = PdfWriter()
    for page in reader.pages:
        src = pick_crop_source(page)
        rect = RectangleObject(rect_to_tuple(src))
        page.mediabox = rect
        page.cropbox = rect
        page.trimbox = rect
        page.bleedbox = rect
        try:
            page.artbox = rect
        except Exception:
            pass
        writer.add_page(page)
    with open(caminho_saida, "wb") as f:
        writer.write(f)


def processar_texto_sumario(linhas):
    itens_aulas, itens_gabaritos = [], []
    disciplina_atual, frente_atual, ultima_disciplina_valida = "", "", ""

    i = 0
    total_linhas = len(linhas)
    while i < total_linhas:
        linha = linhas[i].strip()
        if not linha:
            i += 1
            continue

        if linha in [
            "LINGUAGENS, CÓDIGOS E SUAS TECNOLOGIAS",
            "MATEMÁTICA E SUAS TECNOLOGIAS",
            "CIÊNCIAS HUMANAS E SUAS TECNOLOGIAS",
            "CIÊNCIAS DA NATUREZA E SUAS TECNOLOGIAS"
        ]:
            i += 1
            continue

        if "– Frente" in linha or "- Frente" in linha:
            m = re.search(
                r'^(.*?)\s*[–-]\s*(Frente\s+[\w\dº]+)', linha, re.IGNORECASE)
            if m:
                disciplina_atual = m.group(1).strip()
                frente_atual = m.group(2).strip()
                ultima_disciplina_valida = disciplina_atual
            i += 1
            continue

        if linha.startswith("Aula") or linha.startswith("CAPÍTULO") or linha.startswith("Gabarito"):
            item_texto = linha
            while (i + 1 < total_linhas and
                   not re.search(r'\d+$', item_texto) and
                   not linhas[i+1].startswith("Aula") and
                   not linhas[i+1].startswith("CAPÍTULO") and
                   not linhas[i+1].startswith("Gabarito") and
                   not ("– Frente" in linhas[i+1] or "- Frente" in linhas[i+1]) and
                   not "TECNOLOGIAS" in linhas[i+1]):
                i += 1
                item_texto += " " + linhas[i].strip()

            if item_texto.startswith("Gabarito"):
                m_gab = re.match(
                    r'^(Gabarito\s+[^\d]+)\s*[\.\s\u200b\u2009]+(\d+)$', item_texto, re.IGNORECASE)
                if not m_gab:
                    m_gab = re.match(
                        r'^(Gabarito)\s*[\.\s\u200b\u2009]+(\d+)$', item_texto, re.IGNORECASE)
                    rotulo_gab = "Gabarito"
                    pag_gab = int(m_gab.group(2)) if m_gab else 1
                else:
                    rotulo_gab = m_gab.group(1).strip()
                    pag_gab = int(m_gab.group(2))

                if m_gab:
                    disc_gab = disciplina_atual or ultima_disciplina_valida
                    itens_gabaritos.append({
                        'disciplina': disc_gab,
                        'frente': frente_atual,
                        'tipo': 'Gabarito',
                        'rotulo': rotulo_gab,
                        'titulo': rotulo_gab,
                        'pag_inicio': pag_gab
                    })
            else:
                m_aula = re.match(
                    r'^((?:Aula[s]?|CAPÍTULO)\s+\d+(?:\s*(?:e|a|à|-)\s*\d+)?)\s+(.*?)\s*[\.\s\u200b\u2009]+(\d+)$', item_texto, re.IGNORECASE)
                if not m_aula:
                    m_aula = re.match(
                        r'^((?:Aula[s]?|CAPÍTULO)\s+\d+(?:\s*(?:e|a|à|-)\s*\d+)?)\s*[\.\s\u200b\u2009]+(\d+)$', item_texto, re.IGNORECASE)
                    if m_aula:
                        rotulo_raw = m_aula.group(1).strip()
                        titulo = ""
                        pag_in = int(m_aula.group(2))
                else:
                    rotulo_raw = m_aula.group(1).strip()
                    titulo = re.sub(r'[\.\s]+$', '', m_aula.group(2).strip())
                    pag_in = int(m_aula.group(3))

                if m_aula:
                    if disciplina_atual:
                        ultima_disciplina_valida = disciplina_atual

                    eh_primeira_frente = False
                    if itens_aulas:
                        ultima_aula = itens_aulas[-1]
                        if ultima_aula['frente'] != frente_atual or ultima_aula['disciplina'] != disciplina_atual:
                            eh_primeira_frente = True

                    itens_aulas.append({
                        'disciplina': disciplina_atual,
                        'frente': frente_atual,
                        'tipo': 'Aula',
                        'rotulo': rotulo_raw,
                        'titulo': titulo,
                        'pag_inicio': pag_in,
                        'recuar_uma_pagina': eh_primeira_frente
                    })
        i += 1

    primeira_aula_idx = next((idx for idx, item in enumerate(
        itens_aulas) if item['tipo'] == 'Aula'), None)
    if primeira_aula_idx is not None:
        primeira_pag_aula = itens_aulas[primeira_aula_idx]['pag_inicio']
        pag_fim_iniciais = primeira_pag_aula - 2 if primeira_pag_aula > 2 else 1

        if pag_fim_iniciais >= 1:
            itens_aulas.insert(0, {
                'disciplina': itens_aulas[primeira_aula_idx]['disciplina'],
                'frente': itens_aulas[primeira_aula_idx]['frente'],
                'tipo': 'Iniciais',
                'rotulo': 'Iniciais',
                'titulo': 'Iniciais',
                'pag_inicio': 1,
                'pag_fim': pag_fim_iniciais
            })
            itens_aulas[primeira_aula_idx +
                        1]['pag_inicio'] = pag_fim_iniciais + 1
            itens_aulas[primeira_aula_idx + 1]['recuar_uma_pagina'] = False

    for item in itens_aulas:
        if item.get('recuar_uma_pagina', False):
            item['pag_inicio'] -= 1

    total_aulas = len(itens_aulas)
    for idx, item in enumerate(itens_aulas):
        if item['tipo'] == 'Iniciais':
            continue
        pag_limite = next(
            (itens_aulas[p]['pag_inicio'] - 1 for p in range(idx + 1, total_aulas)), None)
        if pag_limite is None and itens_gabaritos:
            pag_limite = itens_gabaritos[0]['pag_inicio'] - 1
        item['pag_fim'] = pag_limite if (
            pag_limite is not None and pag_limite >= item['pag_inicio']) else item['pag_inicio'] + 9

    total_gabaritos = len(itens_gabaritos)
    for idx, item in enumerate(itens_gabaritos):
        pag_limite = next(
            (itens_gabaritos[p]['pag_inicio'] - 1 for p in range(idx + 1, total_gabaritos)), None)
        item['pag_fim'] = pag_limite if (
            pag_limite is not None and pag_limite >= item['pag_inicio']) else item['pag_inicio'] + 3

    return itens_aulas + itens_gabaritos


def gerar_nome_simplificado(item):
    sigla_mat = obter_sigla_disciplina(item['disciplina'])
    if item['tipo'] == 'Iniciais':
        return f"00_Iniciais_{item.get('pag_inicio', 1)}_{item.get('pag_fim', 12)}"
    if item['tipo'] == 'Gabarito':
        return f"{sigla_mat}-Gabarito_{item.get('pag_inicio', 1)}_{item.get('pag_fim', 1)}"

    sigla_fre = extrair_numero_frente(item['frente']).upper()
    rotulo = item['rotulo']
    nums = re.findall(r'\d+', rotulo)

    if not nums:
        str_aulas = "A1"
    elif len(nums) == 1:
        str_aulas = f"A{nums[0]}"
    elif len(nums) == 2:
        rotulo_lower = rotulo.lower()
        if " e " in rotulo_lower:
            str_aulas = f"A{nums[0]}e{nums[1]}"
        elif " a " in rotulo_lower or " à " in rotulo_lower:
            str_aulas = f"A{nums[0]}a{nums[1]}"
        else:
            str_aulas = f"A{nums[0]}e{nums[1]}"
    else:
        str_aulas = f"A{''.join(nums)}"

    titulo = item.get('titulo', '').strip()
    p_ini, p_fim = item.get('pag_inicio', 1), item.get('pag_fim', 1)

    if titulo:
        return f"{sigla_mat}-F{sigla_fre}-{str_aulas}-{titulo}_{p_ini}_{p_fim}"
    else:
        return f"{sigla_mat}-F{sigla_fre}-{str_aulas}_{p_ini}_{p_fim}"


def ler_planilha_inteligente(caminho_excel):
    for h in [0, 1, 2, 3]:
        try:
            df = pd.read_excel(caminho_excel, header=h)
            cols_str = " ".join([str(c) for c in df.columns]).lower()
            if "subcomponente" in cols_str or "cód" in cols_str or "link" in cols_str or "arquivo" in cols_str:
                for col in df.columns:
                    if "cód" in str(col).lower() and "cat" in str(col).lower():
                        df['COD_CAT_LIMPO'] = df[col].astype(
                            str).str.strip().str.lstrip('0')
                        break
                return df
        except Exception:
            continue
    try:
        return pd.read_excel(caminho_excel, header=0)
    except Exception:
        return None


def buscar_nome_na_planilha(item, df_planilha, codigo_pdf_limpo):
    if df_planilha is None:
        return None
    try:
        sub_alvo, frente_alvo, caps_alvo = parsear_nome_simplificado_para_busca(
            item)
        df_busca = df_planilha
        if codigo_pdf_limpo and 'COD_CAT_LIMPO' in df_planilha.columns:
            df_livro = df_planilha[df_planilha['COD_CAT_LIMPO'] == str(
                int(codigo_pdf_limpo))]
            if not df_livro.empty:
                df_busca = df_livro

        for _, row in df_busca.iterrows():
            row_dict = {remover_acentos(str(k)).lower(
            ).strip(): v for k, v in row.items()}
            sub_plan_raw = next((str(row_dict[k]) for k in [
                                'subcomponente', 'disciplina', 'materia', 'componente'] if k in row_dict), "")
            sub_plan = remover_acentos(sub_plan_raw).lower().strip()

            cap_plan_val = next((row_dict[k] for k in ['capitulo', 'cap.', 'aula', 'cap',
                                'numero', 'titulo capitulo', 'titulo do capitulo'] if k in row_dict), None)
            caps_plan = extrair_numeros_planilha(cap_plan_val)

            link_val = ""
            for k, v in row_dict.items():
                if "link" in k or "arquivo" in k:
                    link_val = str(v).strip()
                    break

            if not link_val or link_val.lower() == 'nan':
                continue
            if link_val.lower().endswith('.pdf'):
                link_val = link_val[:-4]

            match_sub = (not sub_alvo or sub_alvo in sub_plan or sub_plan in sub_alvo or any(
                w in sub_plan for w in sub_alvo.split()))

            if item['tipo'] == 'Gabarito':
                continue
            elif item['tipo'] == 'Iniciais':
                if 0 in caps_plan and link_val:
                    return link_val
            else:
                if 0 in caps_plan:
                    continue
                fre_plan_raw = next(
                    (str(row_dict[k]) for k in ['frente', 'frente '] if k in row_dict), "")
                fre_plan_clean = fre_plan_raw.lower().strip()
                fre_plan = "u" if (not fre_plan_clean or fre_plan_clean ==
                                   'nan' or "unica" in fre_plan_clean) else extrair_numero_frente(fre_plan_clean)

                caps_batem = set(caps_alvo).issubset(set(caps_plan)) or set(
                    caps_plan).issubset(set(caps_alvo)) or (caps_alvo == caps_plan)

                if match_sub and (frente_alvo == 'u' or fre_plan == 'u' or fre_plan == frente_alvo) and caps_batem:
                    return link_val
    except Exception:
        pass
    return None

# ==========================================
# INTERFACE STREAMLIT
# ==========================================


uploaded_txt = st.file_uploader("Arquivo de Sumário (.txt)", type=["txt"])
uploaded_pdf = st.file_uploader("PDF do Livro completo (.pdf)", type=["pdf"])
uploaded_excel = st.file_uploader(
    "Planilha de Nomenclatura (.xlsx, .xls) - Opcional", type=["xlsx", "xls"])

if uploaded_txt and uploaded_pdf:
    # Ler TXT
    linhas = uploaded_txt.getvalue().decode("utf-8").splitlines()
    itens_globais = processar_texto_sumario(linhas)

    # Processar Excel se houver
    df_planilha = None
    if uploaded_excel:
        df_planilha = ler_planilha_inteligente(uploaded_excel)

    nome_base = os.path.splitext(uploaded_pdf.name)[0]
    match_num = re.match(r'^(\d+)', nome_base)
    codigo_pdf_limpo = str(int(match_num.group(1))) if match_num else nome_base

    st.subheader("Mapeamento e Nomes Finais")

    dados_finais = []
    for idx, item in enumerate(itens_globais):
        nome_simp = gerar_nome_simplificado(item)
        nome_sugerido = (buscar_nome_na_planilha(
            item, df_planilha, codigo_pdf_limpo) if df_planilha is not None else None) or nome_simp

        col1, col2 = st.columns([1, 2])
        with col1:
            st.text(nome_simp)
        with col2:
            novo_nome = st.text_input(
                f"Nome {idx}", value=nome_sugerido, key=f"input_{idx}", label_visibility="collapsed")
            dados_finais.append((item, novo_nome))

    if st.button("Processar e Separar Aulas", type="primary"):
        with st.spinner("Aplicando crop, dividir páginas e otimizar PDFs..."):
            # Salvar PDF temporariamente para leitura por caminhos físicos
            tmp_pdf_path = "temp_input.pdf"
            with open(tmp_pdf_path, "wb") as f:
                f.write(uploaded_pdf.getbuffer())

            tmp_cropped_path = "temp_cropped.pdf"
            aplicar_crop_pdf(tmp_pdf_path, tmp_cropped_path)

            reader = PdfReader(tmp_cropped_path)
            num_paginas_total = len(reader.pages)

            # Criar um ZIP em memória para download
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for item, nome_personalizado in dados_finais:
                    if not nome_personalizado.lower().endswith('.pdf'):
                        nome_personalizado += '.pdf'
                    nome_limpo = limpar_nome_arquivo(nome_personalizado)

                    idx_inicio = max(
                        0, min(item['pag_inicio'] - 1, num_paginas_total - 1))
                    idx_fim = max(idx_inicio, min(
                        item['pag_fim'] - 1, num_paginas_total - 1))

                    writer_aula = PdfWriter()
                    for pag_idx in range(idx_inicio, idx_fim + 1):
                        writer_aula.add_page(reader.pages[pag_idx])

                    writer_aula.add_outline_item(nome_limpo, page_number=0)

                    pdf_bytes_io = io.BytesIO()
                    writer_aula.write(pdf_bytes_io)

                    zip_file.writestr(nome_limpo, pdf_bytes_io.getvalue())

            # Limpeza de ficheiros temporários
            for p in [tmp_pdf_path, tmp_cropped_path]:
                if os.path.exists(p):
                    os.remove(p)

            st.success("Processamento concluído com sucesso!")
            st.download_button(
                label="Baixar arquivo ZIP com as aulas",
                data=zip_buffer.getvalue(),
                file_name=f"{nome_base}_aulas_separadas.zip",
                mime="application/zip"
            )
