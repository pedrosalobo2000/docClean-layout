# Lê o conteúdo atual do app.py
with open("app.py", "r", encoding="utf-8") as f:
    conteudo = f.read()

# --- 1. Adicionar "import re" depois de "import io" ---
if "import re" not in conteudo:
    conteudo = conteudo.replace("import io\n", "import io\nimport re\n")

# --- 2. Adicionar o padrão de citação e a função verificar_citacoes ---
# Vamos inserir esse bloco logo antes da linha "def formatar_documento"
bloco_citacoes = '''
# Padrão (regex) para encontrar citações no formato:
# (SOBRENOME, ano) ou (SOBRENOME, ano, p. XX)
PADRAO_CITACAO = re.compile(
    r'\\(([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\\s;\\.]*),\\s*(\\d{1,4})\\s*(?:,\\s*(p\\.?\\s*\\d+))?\\)'
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
                if not re.match(r'^p\\.\\s\\d+$', pagina_bruta):
                    problemas.append({
                        "tipo": "Citação - Página",
                        "local": f"Parágrafo {numero}",
                        "mensagem": f"Em {citacao_completa}, use o formato 'p. XX' (com ponto e espaço) para a página"
                    })

    return citacoes_encontradas, problemas


'''

if "def verificar_citacoes" not in conteudo:
    conteudo = conteudo.replace("def formatar_documento(", bloco_citacoes + "def formatar_documento(")

# --- 3. Adicionar o botão "Verificar citações" na tela ---
bloco_botao = '''    st.divider()

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

    st.subheader("✨ Formatação automática")'''

if "Verificador de citações" not in conteudo:
    conteudo = conteudo.replace(
        '    st.divider()\n\n    st.subheader("✨ Formatação automática")',
        bloco_botao
    )

# Salva o arquivo corrigido
with open("app.py", "w", encoding="utf-8") as f:
    f.write(conteudo)

print("Correção aplicada com sucesso!")
