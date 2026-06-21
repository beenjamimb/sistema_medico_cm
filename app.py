import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os
import json
from datetime import datetime

# Configuração da API do Gemini
genai.configure(api_key="AQ.Ab8RN6LMeBM2Tp6JzQlZHjSwRz2zmwMywc2ay07x9SF0lL5TlA")

st.set_page_config(page_title="Emissor Médico Inteligente", page_icon="🩺", layout="centered")

# Pasta para armazenar o histórico e modelos de descrições geradas
pasta_banco = "banco_descricoes_cirurgicas"
if not os.path.exists(pasta_banco):
    os.makedirs(pasta_banco)

# Inicialização segura de estados da sessão para evitar loopings
if 'pagina_atual' not in st.session_state:
    st.session_state['pagina_atual'] = 'menu'
if 'lista_opcoes_ia' not in st.session_state:
    st.session_state['lista_opcoes_ia'] = []
if 'cid_selecionado' not in st.session_state:
    st.session_state['cid_selecionado'] = ""
if 'sigtap_selecionado' not in st.session_state:
    st.session_state['sigtap_selecionado'] = ""

# Função auxiliar para converter número em extenso
def dias_por_extenso(n):
    unidades = ["", "UM", "DOIS", "TRÊS", "QUATRO", "CINCO", "SEIS", "SETE", "OITO", "NOVE"]
    de_10_a_19 = ["DEZ", "ONZE", "DOZE", "TREZE", "QUATORZE", "QUINZE", "DEZASSEIS", "DEZESSETE", "DEZOITO", "DEZENOVE"]
    dezenas = ["", "", "VINTE", "TRINTA", "QUARENTA", "CINQUENTA", "SESSENTA", "SETENTA", "OITENTA", "NOVENTA"]
    if n == 0: return "ZERO"
    if n == 100: return "CEM"
    if n < 10: return unidades[n]
    if 10 <= n < 20: return de_10_a_19[n - 10]
    if 20 <= n < 100:
        u = n % 10
        d = n // 10
        return f"{dezenas[d]} E {unidades[u]}" if u > 0 else dezenas[d]
    return str(n)

# FUNÇÃO DE CONSULTA QUE TRAZ MÚLTIPLAS OPÇÕES DA ESPECIALIDADE
def consultar_multiplos_dados_ia(termo_usuario):
    if not termo_usuario or len(termo_usuario) < 3:
        return []
    
    prompt = (
        f"Com base no termo médico, suspeita ou fragmento digitado: '{termo_usuario}', "
        "identifique as variações e possibilidades cirúrgicas mais comuns do SUS (Ortopedia/Cirurgia Geral).\n"
        "Retorne OBRIGATORIAMENTE um array JSON contendo até 5 objetos puros (sem markdown, sem aspas triplas de código). "
        "Cada objeto deve mapear uma patologia específica correspondente ao termo. Use exatamente esta estrutura:\n"
        "[\n"
        "  {\n"
        '    "label_exibicao": "NOME DO DIAGNÓSTICO / CIRURGIA COMPLETA EM CAIXA ALTA",\n'
        '    "cid_codigo": "Código CID-10 correspondente",\n'
        '    "cid_descricao": "Descrição oficial na CID-10 em caixa alta",\n'
        '    "sigtap_codigo": "Código SIGTAP/SUS formatado XX.XX.XX.XXX-X",\n'
        '    "sigtap_descricao": "Nome oficial do procedimento cirúrgico no SIGTAP em caixa alta"\n'
        "  }\n"
        "]"
    )
    
    try:
        model = genai.GenerativeModel('gemini-2.5-flash')
        resposta = model.generate_content(prompt).text.strip()
        
        # Limpeza robusta de blocos markdown
        resposta_limpa = resposta.replace("```json", "").replace("```", "").strip()
        return json.loads(resposta_limpa)
    except:
        return []

senha = st.text_input("Digite sua chave de acesso:", type="password")

if senha == "hrcm":

    # ---------------------------------------------------------
    # PÁGINA: MENU PRINCIPAL
    # ---------------------------------------------------------
    if st.session_state['pagina_atual'] == 'menu':
        st.markdown("<h1 style='text-align: center;'>🩺 Sistema de Gestão Cirúrgica</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>Módulos ortopédicos e de cirurgia geral automatizados</p>", unsafe_allow_html=True)
        st.write("\n")
        
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("📝\n\nCriar Documentos", use_container_width=True):
                st.session_state['pagina_atual'] = 'criar'
                st.rerun()
        with col_btn2:
            if st.button("📚\n\nBiblioteca", use_container_width=True):
                st.session_state['pagina_atual'] = 'biblioteca'
                st.rerun()
        with col_btn3:
            if st.button("📊\n\nCenso", use_container_width=True):
                st.session_state['pagina_atual'] = 'censo'
                st.rerun()

    # ---------------------------------------------------------
    # PÁGINA: CRIAR DOCUMENTOS
    # ---------------------------------------------------------
    elif st.session_state['pagina_atual'] == 'criar':
        if st.button("⬅️ Voltar ao Menu Principal"):
            st.session_state['pagina_atual'] = 'menu'
            # Limpa buscas anteriores ao sair
            st.session_state['lista_opcoes_ia'] = []
            st.session_state['cid_selecionado'] = ""
            st.session_state['sigtap_selecionado'] = ""
            st.rerun()
            
        st.header("📝 Emissor Inteligente Especializado")
        
        st.subheader("1. Identificação e Busca Clínica")
        nome = st.text_input("Nome Completo do Paciente:").strip().upper()
        
        col_busca1, col_busca2 = st.columns([3, 1])
        with col_busca1:
            diagnostico_input = st.text_input("Digite a patologia ou estrutura (Ex: hernia, fratura radio, apendicite):", placeholder="Digite e clique em buscar...").strip()
        with col_busca2:
            st.write("\n\n")
            if st.button("🔍 Buscar Códigos", use_container_width=True):
                if diagnostico_input:
                    with st.spinner("Mapeando opções oficiais..."):
                        resultados = consultar_multiplos_dados_ia(diagnostico_input)
                        if resultados:
                            st.session_state['lista_opcoes_ia'] = resultados
                        else:
                            st.error("Nenhuma opção estruturada encontrada para este termo.")
                else:
                    st.warning("Digite algo para buscar.")

        # Exibe o seletor apenas se houver dados na lista da sessão (evita que sumam ao digitar outros campos)
        if st.session_state['lista_opcoes_ia']:
            lista_labels = [opt['label_exibicao'] for opt in st.session_state['lista_opcoes_ia']]
            selecionado = st.selectbox("🎯 Várias opções encontradas! Selecione o caso exato deste paciente:", lista_labels)
            
            # Atualiza as variáveis com base na escolha do médico dentro do selectbox
            for opt in st.session_state['lista_opcoes_ia']:
                if opt['label_exibicao'] == selecionado:
                    st.session_state['cid_selecionado'] = f"{opt['cid_codigo']} - {opt['cid_descricao']}"
                    st.session_state['sigtap_selecionado'] = f"{opt['sigtap_codigo']} - {opt['sigtap_descricao']}"

        # Campos de validação final preenchidos dinamicamente pela seleção
        st.write("---")
        col_fb1, col_fb2 = st.columns(2)
        with col_fb1:
            termo_busca_cid = st.text_input("Confirmação / Ajuste de CID-10:", value=st.session_state['cid_selecionado']).strip().upper()
        with col_fb2:
            procedimento_sigtap = st.text_input("Confirmação / Ajuste de Procedimento SUS:", value=st.session_state['sigtap_selecionado']).strip().upper()

        st.subheader("2. Equipe Médica e Parâmetros")
        col_eq1, col_eq2 = st.columns(2)
        with col_eq1:
            nome_cirurgiao = st.text_input("Nome do Cirurgião:", value="DR. DANIEL ROCHA E SILVA MODESTO").strip().upper()
        with col_eq2:
            nome_anestesista = st.text_input("Nome do Anestesista:", value="DRA. THAYS MEIRELES DOS SANTOS").strip().upper()

        col1, col2 = st.columns(2)
        with col1:
            afastamento = st.number_input("Dias de Afastamento (Atestado):", min_value=0, value=45, step=1)
        with col2:
            retorno = st.number_input("Dias até o Retorno (Encaminhamento):", min_value=0, value=14, step=1)

        st.subheader("3. Detalhes Clínicos Específicos")
        col_det1, col_det2 = st.columns(2)
        with col_det1:
            regiao_anatomia_raiox = st.text_input("Anatomia para Radiografia (Ex: ABDOME / RÁDIO DISTAL ESQUERDO):").strip().upper()
        with col_det2:
            tipo_imobilizacao = st.text_input("Tipo de Imobilização / Curativo (Ex: CURATIVO LIMPO / TALA GESSADA):").strip().upper()
            
        justificativa_fisio = st.text_area("Justificativa Clínica para Fisioterapia (Se Ortopedia):").strip().upper()

        st.subheader("4. Itens do Lote")
        col_ch1, col_ch2 = st.columns(2)
        with col_ch1:
            emitir_atestado = st.checkbox("Atestado Médico", value=True)
            emitir_encaminhamento = st.checkbox("Ficha de Encaminhamento Ambulatorial (Retorno)", value=True)
            emitir_raiox = st.checkbox("Solicitação de Radiografia", value=False)
        with col_ch2:
            emitir_imobilizacao = st.checkbox("Solicitação de Imobilização / Cuidados", value=False)
            emitir_fisioterapia = st.checkbox("Solicitação de Fisioterapia", value=False)
            emitir_cirurgica = st.checkbox("Incluir Descrição Cirúrgica no Lote", value=True)

        if st.button("📦 Gerar e Compilar Lote (.DOCX)", type="primary"):
            if not nome or not termo_busca_cid or not procedimento_sigtap:
                st.error("Preencha o Nome do Paciente e use a busca para validar o CID e Procedimento.")
            else:
                id_busca = "".join(filter(str.isalnum, procedimento_sigtap.split(" - ")[0])).lower()
                caminho_arquivo = os.path.join(pasta_banco, f"{id_busca}.txt")
                
                texto_descricao_final = ""
                
                if emitir_cirurgica:
                    if os.path.exists(caminho_arquivo):
                        with open(caminho_arquivo, "r", encoding="utf-8") as f:
                            linhas = f.readlines()
                        if linhas and linhas[0].startswith("# NOME:"):
                            texto_descricao_final = "".join(linhas[1:]).strip()
                        else:
                            texto_descricao_final = "".join(linhas).strip()
                    else:
                        # Prompt calibrado para descrições perfeitas estruturadas
                        prompt_estrito = (
                            "Você é um cirurgião sênior do SUS experiente em Ortopedia, Traumatologia e Cirurgia Geral. "
                            f"Gere única e estritamente os passos técnicos cronológicos da descrição cirúrgica para o procedimento: [{procedimento_sigtap}]. "
                            "REGRAS CRÍTICAS: 1. ESCREVA TUDO EM CAIXA ALTA (LETRA MAIÚSCULA). "
                            "2. Escreva exatamente em 7 passos enumerados (1 a 7), curtos, objetivos e cirurgicamente impecáveis. "
                            "3. NÃO inclua nenhum cabeçalho (como nome, cirurgião, etc). "
                            "4. NÃO inclua nenhuma observação final, fechamento ou contagem de compressas. "
                            "5. Retorne apenas as 7 linhas numeradas técnicas da cirurgia."
                        )
                        with st.spinner("IA compilando os passos técnicos cirúrgicos da descrição..."):
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            corpo_ia = model.generate_content(prompt_estrito).text.strip()
                            
                            texto_descricao_final = corpo_ia
                            with open(caminho_arquivo, "w", encoding="utf-8") as f:
                                f.write(f"# NOME: {procedimento_sigtap.upper()}\n")
                                f.write(corpo_ia)

                # --- CONFIGURAÇÃO DE DATAS ---
                meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                hoje = datetime.now()
                data_hoje_barra = hoje.strftime('%d/%m/%Y')
                data_extenso = f"{hoje.strftime('%d')} de {meses[hoje.month]} de {hoje.year}"
                
                paginas_documento = []
                
                if emitir_atestado:
                    extenso_dias = dias_por_extenso(afastamento)
                    texto_atestado = f"Atesto para fins trabalhistas e a pedido do interessado que {nome}, esteve sob os cuidados do hospital, a partir do dia {data_hoje_barra}, para tratamento de CID {termo_busca_cid}, necessitando de {afastamento} ({extenso_dias}) dias de repouso para sua convalescença."
                    paginas_documento.append(("ATESTADO MÉDICO", texto_atestado))
                    
                if emitir_encaminhamento:
                    texto_retorno = (
                        f"Nome do Paciente: {nome}\n\n"
                        f"MARCAR RETORNO NO AMBULATÓRIO DA ESPECIALIDADE APÓS {retorno} DIAS DA CIRURGIA, DE ACORDO COM A DISPONIBILIDADE DA AGENDA DO {nome_cirurgiao}.\n\n"
                        f"ORIENTAÇÕES:\n"
                        f"RETIRAR OS PONTOS EM 14 DIAS NA UBS MAIS PRÓXIMA.\n"
                        f"EVITAR PEGAR PESO E GRANDES ESFORÇOS DURANTE 30 DIAS.\n"
                        f"MANTER REPOUSO CONFORME ORIENTADO.\n"
                        f"NÃO UTILIZAR CREMES, POMADAS OU OUTRAS SUBSTÂNCIAS SOBRE A FERIDA OPERATÓRIA QUE NÃO TENHAM SIDO PRESCRITAS POR MÉDICO."
                    )
                    paginas_documento.append(("FICHA DE ENCAMINHAMENTO AMBULATORIAL", texto_retorno))

                if emitir_raiox:
                    texto_raiox = (
                        f"Nome do Paciente: {nome}\n\n"
                        f"SOLICITO\n\n"
                        f"RADIOGRAFIA DE {regiao_anatomia_raiox} AP E PERFIL\n"
                        f"- REALIZAR APENAS NO RETORNO"
                    )
                    paginas_documento.append(("SOLICITAÇÃO DE EXAMES", texto_raiox))

                if emitir_imobilizacao:
                    texto_imobilizacao = (
                        f"Nome do Paciente: {nome}\n\n"
                        f"SOLICITO\n\n"
                        f"CUIDADOS / IMOBILIZAÇÃO TIPO {tipo_imobilizacao}\n\n"
                        f"ORIENTAÇÕES:\n"
                        f"Manter a região limpa e seca. Caso haja imobilização gessada, movimentar extremidades livres.\n"
                        f"Não apoiar peso nem remover proteções sem expressa orientação médica.\n"
                        f"Procurar atendimento de urgência se notar: dor refratária, sangramento ativo, febre, formigamento ou alteração de temperatura local."
                    )
                    paginas_documento.append(("RECEITUÁRIO COMUM", texto_imobilizacao))

                if emitir_fisioterapia and justificativa_fisio:
                    texto_fisioterapia = (
                        f"Nome do Paciente: {nome}\n\n"
                        f"SOLICITO\n\n"
                        f"20 SESSÕES DE FISIOTERAPIA MOTORA\n\n"
                        f"JUSTIFICATIVA:\n"
                        f"{justificativa_fisio}"
                    )
                    paginas_documento.append(("FICHA DE ENCAMINHAMENTO PARA FISIOTERAPIA", texto_fisioterapia))

                # COMPOSIÇÃO DA DESCRIÇÃO CIRÚRGICA
                if emitir_cirurgica and texto_descricao_final:
                    nome_cirurgia_limpo = procedimento_sigtap.split(' - ')[1] if ' - ' in procedimento_sigtap else procedimento_sigtap
                    codigo_sus_limpo = procedimento_sigtap.split(' - ')[0] if ' - ' in procedimento_sigtap else ''
                    
                    cabecalho_fixo = (
                        f"NOME DO PACIENTE: {nome}\n"
                        "DATA DE NASCIMENTO: \n\n"
                        f"CIRURGIA - {nome_cirurgia_limpo}\n"
                        f"LATERALIDADE - \n"
                        f"CÓDIGO SUS - {codigo_sus_limpo}\n\n"
                        f"CIRURGIÃO: {nome_cirurgiao}\n"
                        f"ANESTESISTA: {nome_anestesista}\n\n"
                        "DESCRIÇÃO CIRÚRGICA:\n"
                    )
                    
                    observacoes_fixas = (
                        "\n\nOBSERVAÇÃO:\n"
                        "- NÃO HOUVE INTERCORRÊNCIAS DURANTE O PROCEDIMENTO CIRÚRGICO\n"
                        "- PACIENTE ENCAMINHADO (A) PARA SALA DE RECUPERAÇÃO PÓS ANESTÉSICA, APÓS A LIBERAÇÃO PELA EQUIPE DE ANESTESIOLOGIA"
                    )
                    
                    documento_cirurgico_completo = cabecalho_fixo + texto_descricao_final + observacoes_fixas
                    paginas_documento.append(("FICHA DE DESCRIÇÃO CIRÚRGICA", documento_cirurgico_completo))

                # --- ENGINE DO WORD (TÍTULO 13 / CORPO 11) ---
                doc_word = Document()
                def obter_imagem(nome_base):
                    for ext in [".png", ".jpg", ".jpeg"]:
                        if os.path.exists(nome_base + ext): return nome_base + ext
                    return None

                img_topo, img_rodape = obter_imagem("topo"), obter_imagem("rodape")

                for section in doc_word.sections:
                    section.top_margin = Inches(1.8)
                    section.bottom_margin = Inches(1.6)
                    section.left_margin = Inches(1.2)
                    section.right_margin = Inches(1.2)
                    if img_topo:
                        section.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        section.header.paragraphs[0].add_run().add_picture(img_topo, width=Inches(6.0))
                    if img_rodape:
                        section.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        section.footer.paragraphs[0].add_run().add_picture(img_rodape, width=Inches(6.0))

                for i, (titulo, corpo) in enumerate(paginas_documento):
                    p_titulo = doc_word.add_paragraph()
                    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_titulo.paragraph_format.space_before = Pt(14)
                    p_titulo.paragraph_format.space_after = Pt(24)
                    
                    run_titulo = p_titulo.add_run(titulo)
                    run_titulo.font.name = 'Arial'
                    run_titulo.font.size = Pt(13)
                    run_titulo.bold = True
                    
                    for linha in corpo.split('\n'):
                        if linha.strip() == "":
                            doc_word.add_paragraph().paragraph_format.space_after = Pt(4)
                            continue
                        p_linha = doc_word.add_paragraph()
                        p_linha.paragraph_format.line_spacing = 1.2
                        p_linha.paragraph_format.space_after = Pt(4)
                        
                        if len(linha) < 110 or ":" in linha or list(filter(linha.strip().startswith, ["-", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "20", "ORIENTAÇÕES:", "CIRURGIA", "LATERALIDADE", "CÓDIGO", "FICHA"])):
                            p_linha.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        else:
                            p_linha.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            
                        run_linha = p_linha.add_run(linha)
                        run_linha.font.name = 'Arial'
                        run_linha.font.size = Pt(11)
                    
                    p_assinatura = doc_word.add_paragraph()
                    p_assinatura.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    p_assinatura.paragraph_format.space_before = Pt(36)
                    p_assinatura.add_run(f"Campo Maior - PI, {data_extenso}.\n\n\n__________________________________\nCarimbo e Assinatura").font.name = 'Arial'
                    
                    if i < len(paginas_documento) - 1:
                        doc_word.add_page_break()
                
                bio = io.BytesIO()
                doc_word.save(bio)
                
                st.success("✨ Lote gerado e compilado com sucesso!")
                st.download_button(
                    label="📥 Baixar Prontuário em Lote .DOCX",
                    data=bio.getvalue(),
                    file_name=f"Kit_Cirurgico_{nome.lower().replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    # ---------------------------------------------------------
    # PÁGINA: BIBLIOTECA DE DESCRIÇÕES
    # ---------------------------------------------------------
    elif st.session_state['pagina_atual'] == 'biblioteca':
        if st.button("⬅️ Voltar ao Menu Principal"):
            st.session_state['pagina_atual'] = 'menu'
            st.rerun()
            
        st.header("📚 Biblioteca de Descrições Salvas")
        st.write("Os modelos gerados ficam arquivados aqui. Você pode editá-los e salvá-los permanentemente.")
        
        arquivos = [f for f in os.listdir(pasta_banco) if f.endswith(".txt")]
        
        if not arquivos:
            st.info("Nenhuma descrição salva no histórico local ainda. Elas surgirão conforme novos documentos forem emitidos.")
        else:
            for arquivox in arquivos:
                caminho_completo = os.path.join(pasta_banco, arquivox)
                codigo_arquivo = arquivox.replace(".txt", "").upper()
                
                with open(caminho_completo, "r", encoding="utf-8") as f:
                    linhas = f.readlines()
                
                titulo_exibicao = f"CÓDIGO: {codigo_arquivo}"
                if linhas and linhas[0].startswith("# NOME:"):
                    nome_identificado = linhas[0].replace("# NOME:", "").strip()
                    titulo_exibicao = f"{nome_identificado} - {codigo_arquivo}"
                
                with st.expander(f"📋 {titulo_exibicao}"):
                    conteudo_atual = "".join(linhas)
                    novo_conteudo = st.text_area(f"Editor de texto completo ({codigo_arquivo}):", value=conteudo_atual, height=300, key=arquivox)
                    
                    if st.button(f"Salvar Alterações em {codigo_arquivo}", key=f"btn_{arquivox}"):
                        with open(caminho_completo, "w", encoding="utf-8") as f:
                            f.write(novo_conteudo)
                        st.success(f"Modelo {codigo_arquivo} updated com sucesso!")
                        st.rerun()

    # ---------------------------------------------------------
    # PÁGINA: CENSO HOSPITALAR
    # ---------------------------------------------------------
    elif st.session_state['pagina_atual'] == 'censo':
        if st.button("⬅️ Voltar ao Menu Principal"):
            st.session_state['pagina_atual'] = 'menu'
            st.rerun()
            
        st.header("📊 Censo Hospitalar")
        st.info("Espaço reservado para o Censo de Internações/Plantão. Funcionalidade em desenvolvimento.")

elif senha != "":
    st.error("Chave de acesso inválida.")
