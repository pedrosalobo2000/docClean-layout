import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

# Configuração básica da página
st.set_page_config(
    page_title="DocClean & Layout",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocClean & Layout")
st.write(
    "Bem-vindo! Este aplicativo vai ajudar a formatar automaticamente "
    "seus arquivos .docx (fontes, espaçamento, margens, etc.)."
)

st.divider()

# ============================================================
# PERFIS DE FORMATAÇÃO PRÉ-DEFINIDOS
# ============================================================
PERFIS_FORMATACAO = {
    "ABNT Padrão": {
        "fonte": "Times New Roman", "tamanho": 12, "espacamento": 1.5,
        "margem_superior": 3.0, "margem_inferior": 2.0,
        "margem_esquerda": 3.0, "margem_direita": 2.0,
        "recuo_primeira_linha": False,
    },
    "TCC / Monografia": {
        "fonte": "Times New Roman", "tamanho": 12, "espacamento": 1.5,
        "margem_superior": 3.0, "margem_inferior": 2.0,
        "margem_esquerda": 3.0, "margem_direita": 2.0,
        "recuo_primeira_linha": True,
    },
    "Personalizado": None,
}

st.subheader("⚙️ Opções de formatação")

perfil_escolhido = st.selectbox("Perfil de formatação", list(PERFIS_FORMATACAO.keys()))

if perfil_escolhido != "Personalizado":
    dados_perfil = PERFIS_FORMATACAO[perfil_escolhido]
    st.info(
        f"Perfil **{perfil_escolhido}**: {dados_perfil['fonte']} {dados_perfil['tamanho']}, "
        f"espaçamento {dados_perfil['espacamento']}, margens "
        f"{dados_perfil['margem_superior']}-{dados_perfil['margem_direita']}-"
        f"{dados_perfil['margem_inferior']}-{dados_perfil['margem_esquerda']}cm"
        + (", com recuo na primeira linha" if dados_perfil["recuo_primeira_linha"] else "")
    )
    fonte_escolhida = dados_perfil["fonte"]
    tamanho_escolhido = dados_perfil["tamanho"]
    espacamento_escolhido = dados_perfil["espacamento"]
    margem_sup_escolhida = dados_perfil["margem_superior"]
    margem_inf_escolhida = dados_perfil["margem_inferior"]
    margem_esq_escolhida = dados_perfil["margem_esquerda"]
    margem_dir_escolhida = dados_perfil["margem_direita"]
    recuo_escolhido = dados_perfil["recuo_primeira_linha"]
else:
    col1, col2 = st.columns(2)
    with col1:
        fonte_escolhida = st.selectbox("Fonte", ["Arial", "Times New Roman", "Calibri"])
    with col2:
        tamanho_escolhido = st.selectbox("Tamanho da fonte", [10, 11, 12, 14], index=2)

    espacamento_escolhido = st.selectbox("Espaçamento entre linhas", [1.0, 1.5, 2.0], index=1)

    col3, col4 = st.columns(2)
    with col3:
        margem_sup_escolhida = st.selectbox("Margem superior (cm)", [2.0, 2.5, 3.0], index=1)
        margem_esq_escolhida = st.selectbox("Margem esquerda (cm)", [2.0, 2.5, 3.0], index=1)
    with col4:
        margem_inf_escolhida = st.selectbox("Margem inferior (cm)", [2.0, 2.5, 3.0], index=1)
        margem_dir_escolhida = st.selectbox("Margem direita (cm)", [2.0, 2.5, 3.0], index=1)

    recuo_escolhido = st.checkbox("Aplicar recuo na primeira linha dos parágrafos")

st.divider()

# --- Upload do arquivo ---
arquivo = st.file_uploader("Envie seu arquivo .docx", type=["docx"])


# ============================================================
# REGRAS ABNT (valores de referência para o verificador)
# ============================================================
ABNT_FONTES_VALIDAS = ["Arial", "Times New Roman"]
ABNT_TAMANHO_FONTE = 12
ABNT_ESPACAMENTO = 1.5
ABNT_MARGEM_SUPERIOR_CM = 3.0
ABNT_MARGEM_ESQUERDA_CM = 3.0
ABNT_MARGEM_INFERIOR_CM = 2.0
ABNT_MARGEM_DIREITA_CM = 2.0


def verificar_abnt(documento):
    """
    Analisa o documento e retorna uma lista de problemas encontrados,
    comparando com as regras ABNT definidas acima.
    Cada problema é um dicionário com: tipo, local e mensagem.
    """
    problemas = []

    # --- 1. Verificar margens (a regra vale para o documento todo) ---
    for secao in documento.sections:
        margem_top_cm = secao.top_margin.cm
        margem_left_cm = secao.left_margin.cm
        margem_bottom_cm = secao.bottom_margin.cm
        margem_right_cm = secao.right_margin.cm

        if round(margem_top_cm, 1) != ABNT_MARGEM_SUPERIOR_CM:
            problemas.append({
                "tipo": "Margem",
                "local": "Página",
                "mensagem": f"Margem superior é {margem_top_cm:.1f}cm (ABNT exige {ABNT_MARGEM_SUPERIOR_CM}cm)"
            })
        if round(margem_left_cm, 1) != ABNT_MARGEM_ESQUERDA_CM:
            problemas.append({
                "tipo": "Margem",
                "local": "Página",
                "mensagem": f"Margem esquerda é {margem_left_cm:.1f}cm (ABNT exige {ABNT_MARGEM_ESQUERDA_CM}cm)"
            })
        if round(margem_bottom_cm, 1) != ABNT_MARGEM_INFERIOR_CM:
            problemas.append({
                "tipo": "Margem",
                "local": "Página",
                "mensagem": f"Margem inferior é {margem_bottom_cm:.1f}cm (ABNT exige {ABNT_MARGEM_INFERIOR_CM}cm)"
            })
        if round(margem_right_cm, 1) != ABNT_MARGEM_DIREITA_CM:
            problemas.append({
                "tipo": "Margem",
                "local": "Página",
                "mensagem": f"Margem direita é {margem_right_cm:.1f}cm (ABNT exige {ABNT_MARGEM_DIREITA_CM}cm)"
            })

    # --- 2. Verificar cada parágrafo: fonte, tamanho, espaçamento e alinhamento ---
    for numero, paragrafo in enumerate(documento.paragraphs, start=1):
        if paragrafo.text.strip() == "":
            continue

        espacamento = paragrafo.paragraph_format.line_spacing
        if espacamento is not None and round(espacamento, 1) != ABNT_ESPACAMENTO:
            problemas.append({
                "tipo": "Espaçamento",
                "local": f"Parágrafo {numero}",
                "mensagem": f"Espaçamento é {espacamento} (ABNT exige {ABNT_ESPACAMENTO})"
            })

        try:
            alinhamento = paragrafo.alignment
        except ValueError:
            alinhamento = None

        if alinhamento is not None and alinhamento != WD_ALIGN_PARAGRAPH.JUSTIFY:
            problemas.append({
                "tipo": "Alinhamento",
                "local": f"Parágrafo {numero}",
                "mensagem": "Parágrafo não está justificado"
            })

        for run in paragrafo.runs:
            if run.text.strip() == "":
                continue

            nome_fonte = run.font.name
            if nome_fonte is not None and nome_fonte not in ABNT_FONTES_VALIDAS:
                problemas.append({
                    "tipo": "Fonte",
                    "local": f"Parágrafo {numero}",
                    "mensagem": f"Fonte '{nome_fonte}' não é padrão ABNT (use Arial ou Times New Roman)"
                })

            tamanho_fonte = run.font.size
            if tamanho_fonte is not None and tamanho_fonte.pt != ABNT_TAMANHO_FONTE:
                problemas.append({
                    "tipo": "Tamanho da fonte",
                    "local": f"Parágrafo {numero}",
                    "mensagem": f"Tamanho é {tamanho_fonte.pt}pt (ABNT exige {ABNT_TAMANHO_FONTE}pt)"
                })

    return problemas



# Padrão (regex) para encontrar citações no formato:
# (SOBRENOME, ano) ou (SOBRENOME, ano, p. XX)
PADRAO_CITACAO = re.compile(
    r'\(([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s;\.]*),\s*(\d{1,4})\s*(?:,\s*(p\.?\s*\d+))?\)'
)


def verificar_citacoes(documento):
    """
    Procura citações no formato (AUTOR, ano) ou (AUTOR, ano, p. XX)
    em cada parágrafo do documento, e verifica se seguem o padrão ABNT.
    Retorna uma tupla: (lista_de_citacoes_encontradas, lista_de_problemas)
    """
    citacoes_encontradas = []
    problemas = []

    for numero, paragrafo in enumerate(documento.paragraphs, start=1):
        texto = paragrafo.text
        if texto.strip() == "":
            continue

        for match in PADRAO_CITACAO.finditer(texto):
            autor_bruto = match.group(1).strip()
            ano_bruto = match.group(2).strip()
            pagina_bruta = match.group(3)

            citacao_completa = match.group(0)
            citacoes_encontradas.append({
                "local": f"Parágrafo {numero}",
                "citacao": citacao_completa
            })

            autor_para_checar = autor_bruto.replace("et al.", "").replace(";", "")
            if autor_para_checar.strip() != autor_para_checar.strip().upper():
                problemas.append({
                    "tipo": "Citação - Autor",
                    "local": f"Parágrafo {numero}",
                    "mensagem": f"Em {citacao_completa}, o autor deveria estar em CAIXA ALTA (ex: SOBRENOME)"
                })

            if len(ano_bruto) != 4:
                problemas.append({
                    "tipo": "Citação - Ano",
                    "local": f"Parágrafo {numero}",
                    "mensagem": f"Em {citacao_completa}, o ano '{ano_bruto}' não tem 4 dígitos"
                })

            if pagina_bruta is not None:
                if not re.match(r'^p\.\s\d+$', pagina_bruta):
                    problemas.append({
                        "tipo": "Citação - Página",
                        "local": f"Parágrafo {numero}",
                        "mensagem": f"Em {citacao_completa}, use o formato 'p. XX' (com ponto e espaço) para a página"
                    })

    return citacoes_encontradas, problemas


def formatar_documento(documento, fonte, tamanho, espacamento,
                        margem_superior, margem_inferior, margem_esquerda, margem_direita,
                        aplicar_recuo=False):
    """
    Recebe um objeto Document (python-docx) e aplica:
    - fonte e tamanho em todos os textos (runs)
    - espaçamento entre linhas em todos os parágrafos
    - alinhamento justificado em todos os parágrafos
    - margens em todas as seções do documento
    - recuo na primeira linha (opcional)
    """
    for paragrafo in documento.paragraphs:
        for run in paragrafo.runs:
            run.font.name = fonte
            run.font.size = Pt(tamanho)

    for paragrafo in documento.paragraphs:
        paragrafo.paragraph_format.line_spacing = espacamento
        if paragrafo.text.strip() != "":
            paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if aplicar_recuo:
                paragrafo.paragraph_format.first_line_indent = Cm(1.25)

    for secao in documento.sections:
        secao.top_margin = Cm(margem_superior)
        secao.bottom_margin = Cm(margem_inferior)
        secao.left_margin = Cm(margem_esquerda)
        secao.right_margin = Cm(margem_direita)

    return documento


if arquivo is not None:
    st.success(f"Arquivo recebido: {arquivo.name}")

    documento = Document(io.BytesIO(arquivo.getvalue()))

    total_paragrafos = len(documento.paragraphs)

    preview_texto = ""
    contador = 0
    for paragrafo in documento.paragraphs:
        if paragrafo.text.strip() != "":
            preview_texto += paragrafo.text + "\n\n"
            contador += 1
        if contador >= 3:
            break

    st.divider()
    st.subheader("📊 Informações do documento original")
    st.write(f"**Total de parágrafos:** {total_paragrafos}")
    st.subheader("👀 Prévia do texto original")
    st.text(preview_texto if preview_texto else "O documento não contém texto visível.")

    st.divider()

    st.subheader("🔍 Verificador de conformidade ABNT")
    if st.button("Verificar conformidade ABNT"):
        problemas = verificar_abnt(documento)

        if not problemas:
            st.success("✅ Nenhum problema encontrado! O documento segue as regras ABNT verificadas.")
        else:
            st.warning(f"⚠️ Foram encontrados {len(problemas)} problema(s):")
            for problema in problemas:
                st.write(f"**[{problema['tipo']}]** {problema['local']}: {problema['mensagem']}")

    st.divider()

    st.subheader("📚 Verificador de citações")
    if st.button("Verificar citações"):
        citacoes, problemas_citacoes = verificar_citacoes(documento)

        if not citacoes:
            st.info("Nenhuma citação no formato (AUTOR, ano) foi encontrada no documento.")
        else:
            st.write(f"**{len(citacoes)} citação(ões) encontrada(s):**")
            for citacao in citacoes:
                st.write(f"- {citacao['local']}: `{citacao['citacao']}`")

            st.divider()

            if not problemas_citacoes:
                st.success("✅ Todas as citações encontradas seguem o padrão ABNT verificado.")
            else:
                st.warning(f"⚠️ Foram encontrados {len(problemas_citacoes)} problema(s) nas citações:")
                for problema in problemas_citacoes:
                    st.write(f"**[{problema['tipo']}]** {problema['local']}: {problema['mensagem']}")

    st.divider()

    st.subheader("✨ Formatação automática")
    if st.button("Formatar documento"):
        documento_formatado = formatar_documento(
            documento,
            fonte_escolhida,
            tamanho_escolhido,
            espacamento_escolhido,
            margem_sup_escolhida, margem_inf_escolhida, margem_esq_escolhida, margem_dir_escolhida,
            aplicar_recuo=recuo_escolhido
        )

        buffer_saida = io.BytesIO()
        documento_formatado.save(buffer_saida)
        buffer_saida.seek(0)

        st.success("Documento formatado com sucesso!")

        st.download_button(
            label="⬇️ Baixar documento formatado",
            data=buffer_saida,
            file_name=f"formatado_{arquivo.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

else:
    st.warning("Nenhum arquivo enviado ainda.")
