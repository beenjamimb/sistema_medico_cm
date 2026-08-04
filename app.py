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

st.set_page_config(page_title="Emissor Médico Especializado", page_icon="🩺", layout="centered")

# Inicialização de variáveis de sessão
if 'pagina_atual' not in st.session_state:
    st.session_state['pagina_atual'] = 'menu'
if 'lista_opcoes_ia' not in st.session_state:
    st.session_state['lista_opcoes_ia'] = []
if 'cid_selecionado' not in st.session_state:
    st.session_state['cid_selecionado'] = ""

# Função para consulta de CIDs com apoio da IA
def consultar_cid_ia(termo_usuario):
    if not termo_usuario or len(termo_usuario) < 3:
        return []
    
    prompt = (
        f"Com base no termo médico, suspeita ou patologia: '{termo_usuario}', "
        "identifique as variações diagnósticas correspondentes na CID-10.\n"
        "Retorne OBRIGATORIAMENTE um array JSON contendo até 5 objetos puros. "
        "Cada objeto deve conter:\n"
        "[\n"
        "  {\n"
        '    "label_exibicao": "NOME DA PATOLOGIA EM CAIXA ALTA (CÓDIGO CID-10)",\n'
        '    "cid_codigo": "Código CID-10",\n'
        '    "cid_descricao": "Descrição oficial na CID-10 em caixa alta"\n'
        "  }\n"
        "]"
    )
    
    try:
        model = genai.GenerativeModel('gemini-2.5-flash')
        resposta = model.generate_content(prompt).text.strip()
        resposta_limpa = resposta.replace("```json", "").replace("```", "").strip()
        return json.loads(resposta_limpa)
    except:
        return []

senha = st.text_input("Digite sua chave de acesso:", type="password")

if senha == "trabalheconquiste":

    # ---------------------------------------------------------
    # PÁGINA: MENU PRINCIPAL
    # ---------------------------------------------------------
    if st.session_state['pagina_atual'] == 'menu':
        st.markdown("<h1 style='text-align: center;'>🩺 Sistema de Emissão de Documentos Médicos</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>Selecione o módulo ou funcionalidade desejada</p>", unsafe_allow_html=True)
        st.write("\n")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("🩺\n\nDocumentos Cirurgia Geral", use_container_width=True):
                st.session_state['pagina_atual'] = 'cirurgia_geral'
                st.rerun()
        with col_m2:
            if st.button("🦴\n\nDocumentos Ortopedia", use_container_width=True):
                st.session_state['pagina_atual'] = 'ortopedia'
                st.rerun()

        st.write("---")
        st.subheader("⚙️ Módulos Adicionais e Termos (Em Desenvolvimento)")
        
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.button("🩸 Solicitar reserva de sangue (Beta)", disabled=True, use_container_width=True)
            st.button("🏥 Solicitar reserva de UTI (Beta)", disabled=True, use_container_width=True)
            st.button("📝 Termo de Consentimento (TCLE) (Beta)", disabled=True, use_container_width=True)
        with col_b2:
            st.button("💉 Termo de anestesia (Beta)", disabled=True, use_container_width=True)
            st.button("⚠️ Termo de amputação (Beta)", disabled=True, use_container_width=True)

    # ---------------------------------------------------------
    # PÁGINA: CIRURGIA GERAL
    # ---------------------------------------------------------
    elif st.session_state['pagina_atual'] == 'cirurgia_geral':
        if st.button("⬅️ Voltar ao Menu Principal"):
            st.session_state['pagina_atual'] = 'menu'
            st.session_state['lista_opcoes_ia'] = []
            st.session_state['cid_selecionado'] = ""
            st.rerun()
            
        st.header("🩺 Emissor de Documentos - Cirurgia Geral")
        
        st.subheader("1. Selecione os Documentos a serem gerados")
        emitir_atestado = st.checkbox("Atestado Médico", value=False)
        emitir_retorno = st.checkbox("Ficha de Retorno (Ambulatório)", value=False)
        emitir_receita = st.checkbox("Receituário Comum", value=False)
        emitir_histo = st.checkbox("Solicitação de Histopatológico", value=False)
        emitir_antibio = st.checkbox("Solicitação de Antibiograma", value=False)

        algum_selecionado = emitir_atestado or emitir_retorno or emitir_receita or emitir_histo or emitir_antibio

        if algum_selecionado:
            st.write("---")
            st.subheader("2. Dados dos Documentos Selecionados")

            nome = st.text_input("Nome do paciente:").strip().upper()

            # Exibe apenas se precisar de histórico do paciente
            if emitir_histo or emitir_antibio:
                col_id1, col_id2 = st.columns(2)
                with col_id1:
                    data_nascimento = st.text_input("Data de nascimento (Ex: 12/05/1985):").strip()
                with col_id2:
                    nome_mae = st.text_input("Nome da mãe ou pai:").strip().upper()
            else:
                data_nascimento = ""
                nome_mae = ""

            # Exibe busca de CID apenas se for Atestado
            cid_final = ""
            incluir_cid_atestado = False
            afastamento = 0
            if emitir_atestado:
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    afastamento = st.number_input("Dias de afastamento:", min_value=0, value=15, step=1)
                with col_p2:
                    st.write("\n")
                    incluir_cid_atestado = st.checkbox("Incluir CID-10 e campo de autorização do paciente", value=True)

                if incluir_cid_atestado:
                    col_b1, col_b2 = st.columns([3, 1])
                    with col_b1:
                        diag_busca = st.text_input("Buscar diagnóstico (Ex: Apendicite, Hérnia Inguinal):", placeholder="Digite para pesquisar na IA...").strip()
                    with col_b2:
                        st.write("\n\n")
                        if st.button("🔍 Buscar CID", use_container_width=True):
                            if diag_busca:
                                with st.spinner("Buscando CIDs correspondentes..."):
                                    res = consultar_cid_ia(diag_busca)
                                    if res:
                                        st.session_state['lista_opcoes_ia'] = res
                                    else:
                                        st.error("Nenhum CID localizado.")

                    if st.session_state['lista_opcoes_ia']:
                        labels = [opt['label_exibicao'] for opt in st.session_state['lista_opcoes_ia']]
                        sel = st.selectbox("🎯 Escolha a patologia exata para preencher o CID-10:", labels)
                        for opt in st.session_state['lista_opcoes_ia']:
                            if opt['label_exibicao'] == sel:
                                st.session_state['cid_selecionado'] = f"{opt['cid_codigo']} - {opt['cid_descricao']}"

                    cid_final = st.text_input("CID-10 Final:", value=st.session_state['cid_selecionado']).strip().upper()

            # Exibe prazo de retorno se for Ficha de Retorno
            retorno = 0
            if emitir_retorno:
                retorno = st.number_input("Dias até o retorno:", min_value=0, value=14, step=1)

            # Exibe campos específicos se for Histopatológico ou Antibiograma
            material_exame = ""
            quadro_clinico = ""
            usa_antibiotico = "NÃO"
            antibiotico_nome_dias = ""
            if emitir_histo or emitir_antibio:
                material_exame = st.text_input("Material coletado:").strip().upper()
                quadro_clinico = st.text_area("Quadro clínico:").strip().upper()

                if emitir_antibio:
                    col_ab1, col_ab2 = st.columns(2)
                    with col_ab1:
                        usa_antibiotico = st.selectbox("Faz uso de antibiótico?", ["NÃO", "SIM"])
                    with col_ab2:
                        if usa_antibiotico == "SIM":
                            antibiotico_nome_dias = st.text_input("Antibiótico(s) e dias de tratamento:").strip().upper()

            st.write("\n")
            if st.button("📦 Gerar Documentos (.DOCX)", type="primary"):
                if not nome:
                    st.error("Por favor, preencha o Nome do paciente.")
                else:
                    hoje_str = datetime.now().strftime('%d/%m/%Y')
                    
                    data_nasc_txt = data_nascimento if data_nascimento else "_________________________"
                    nome_mae_txt = nome_mae if nome_mae else "__________________________________________________"
                    cid_txt = cid_final if cid_final else "__________________________________________________"
                    
                    # Estrutura do documento: (Título, Conteúdo do Corpo)
                    documentos_gerar = []

                    if emitir_atestado:
                        corpo_atestado = (
                            f"Atesto, para os devidos fins, que o(a) Sr(a) {nome} recebeu atendimento neste serviço no dia ____/____/______, "
                            f"e necessita afastamento de suas atividades por {afastamento} dias, a partir desta data.\n\n"
                        )
                        if incluir_cid_atestado:
                            corpo_atestado += (
                                f"CID-10: {cid_txt}\n\n"
                                f"_____________________________\nAssinatura do médico\n\n"
                                f"* Autorizo a divulgação do diagnóstico (CID)\n\n"
                                f"_____________________________________\nPaciente ou responsável legal"
                            )
                        else:
                            corpo_atestado += f"_____________________________\nAssinatura do médico"
                        
                        documentos_gerar.append(("ATESTADO MÉDICO", corpo_atestado))

                    if emitir_retorno:
                        corpo_retorno = (
                            f"Retorno ao Ambulatório de Cirurgia Geral em {retorno} dias para reavaliação pós-operatória.\n\n"
                            f"Orientações:\n"
                            f"• Manter curativo limpo e seco.\n"
                            f"• Utilizar as medicações prescritas.\n"
                            f"• Evitar esforços físicos até nova avaliação.\n"
                            f"• Manter alimentação e hidratação adequadas.\n"
                            f"• Procurar atendimento de urgência em caso de febre, dor, vômitos, vermelhidão intensa, secreção pela ferida, sangramento ou abertura dos pontos.\n\n\n"
                            f"_____________________________\nAssinatura do médico"
                        )
                        documentos_gerar.append(("ENCAMINHAMENTO PARA RETORNO – AMBULATÓRIO DE CIRURGIA GERAL", corpo_retorno))

                    if emitir_receita:
                        corpo_receita = "\n\n\n\n\n\n\n\n\n\n_____________________________\nAssinatura do médico"
                        documentos_gerar.append(("RECEITUÁRIO COMUM", corpo_receita))

                    if emitir_histo:
                        corpo_histo = (
                            f"NOME DO PACIENTE: {nome}\n"
                            f"DATA DE NASCIMENTO: {data_nasc_txt}\n"
                            f"NOME DA MÃE OU PAI: {nome_mae_txt}\n\n"
                            f"MATERIAL: {material_exame}\n"
                            f"ANÁLISE: Histopatológico\n"
                            f"QUADRO CLÍNICO: {quadro_clinico}\n\n\n\n"
                            f"_____________________________\nAssinatura do médico"
                        )
                        documentos_gerar.append(("SOLICITAÇÃO DE HISTOPATOLÓGICO", corpo_histo))

                    if emitir_antibio:
                        corpo_antibio = (
                            f"NOME DO PACIENTE: {nome}\n"
                            f"DATA DE NASCIMENTO: {data_nasc_txt}\n"
                            f"NOME DA MÃE OU PAI: {nome_mae_txt}\n\n"
                            f"MATERIAL: {material_exame}\n"
                            f"ANÁLISE: Antibiograma\n"
                            f"QUADRO CLÍNICO: {quadro_clinico}\n"
                            f"FAZ USO DE ANTIBIÓTICO? {usa_antibiotico}\n"
                        )
                        if usa_antibiotico == "SIM":
                            corpo_antibio += f"ANTIBIÓTICO(S) E DIAS: {antibiotico_nome_dias}\n\n\n"
                        else:
                            corpo_antibio += "\n\n\n"
                        corpo_antibio += "_____________________________\nAssinatura do médico"
                        documentos_gerar.append(("SOLICITAÇÃO DE ANTIBIOGRAMA", corpo_antibio))

                    # GERADOR WORD (.DOCX)
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

                    for i, (titulo, corpo) in enumerate(documentos_gerar):
                        # TÍTULO CENTRALIZADO
                        p_titulo = doc_word.add_paragraph()
                        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_titulo.paragraph_format.space_after = Pt(18)
                        run_tit = p_titulo.add_run(titulo)
                        run_tit.font.name = 'Arial'
                        run_tit.font.size = Pt(13)
                        run_tit.bold = True

                        # CORPO JUSTIFICADO
                        for linha in corpo.split('\n'):
                            if linha.strip() == "":
                                doc_word.add_paragraph().paragraph_format.space_after = Pt(4)
                                continue
                            p_linha = doc_word.add_paragraph()
                            p_linha.paragraph_format.line_spacing = 1.2
                            p_linha.paragraph_format.space_after = Pt(4)
                            p_linha.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                
                            run_linha = p_linha.add_run(linha)
                            run_linha.font.name = 'Arial'
                            run_linha.font.size = Pt(11)

                        # DATA E CIDADE NO RODAPÉ À DIREITA
                        p_data = doc_word.add_paragraph()
                        p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        p_data.paragraph_format.space_before = Pt(20)
                        run_data = p_data.add_run(f"{hoje_str}, Campo Maior - PI")
                        run_data.font.name = 'Arial'
                        run_data.font.size = Pt(10)

                        if i < len(documentos_gerar) - 1:
                            doc_word.add_page_break()

                    bio = io.BytesIO()
                    doc_word.save(bio)
                    
                    st.success("✨ Documentos gerados com sucesso!")
                    st.download_button(
                        label="📥 Baixar Documentos .DOCX",
                        data=bio.getvalue(),
                        file_name=f"Documentos_Cirurgia_Geral_{nome.lower().replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

    # ---------------------------------------------------------
    # PÁGINA: ORTOPEDIA
    # ---------------------------------------------------------
    elif st.session_state['pagina_atual'] == 'ortopedia':
        if st.button("⬅️ Voltar ao Menu Principal"):
            st.session_state['pagina_atual'] = 'menu'
            st.session_state['lista_opcoes_ia'] = []
            st.session_state['cid_selecionado'] = ""
            st.rerun()
            
        st.header("🦴 Emissor de Documentos - Ortopedia e Traumatologia")
        
        st.subheader("1. Selecione os Documentos a serem gerados")
        emitir_atestado = st.checkbox("Atestado Médico", value=False)
        emitir_retorno = st.checkbox("Ficha de Retorno (Ambulatório)", value=False)
        emitir_receita = st.checkbox("Receituário Comum", value=False)
        emitir_imobilizacao = st.checkbox("Solicitação de Imobilização", value=False)
        emitir_fisio = st.checkbox("Solicitação de Fisioterapia", value=False)
        emitir_antibio = st.checkbox("Solicitação de Antibiograma", value=False)

        algum_selecionado = emitir_atestado or emitir_retorno or emitir_receita or emitir_imobilizacao or emitir_fisio or emitir_antibio

        if algum_selecionado:
            st.write("---")
            st.subheader("2. Dados dos Documentos Selecionados")

            nome = st.text_input("Nome do paciente:").strip().upper()

            # Exibe se for Antibiograma
            if emitir_antibio:
                col_id1, col_id2 = st.columns(2)
                with col_id1:
                    data_nascimento = st.text_input("Data de nascimento (Ex: 12/05/1985):").strip()
                with col_id2:
                    nome_mae = st.text_input("Nome da mãe ou pai:").strip().upper()
            else:
                data_nascimento = ""
                nome_mae = ""

            # Exibe busca de CID se for Atestado
            cid_final = ""
            incluir_cid_atestado = False
            afastamento = 0
            if emitir_atestado:
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    afastamento = st.number_input("Dias de afastamento:", min_value=0, value=45, step=1)
                with col_p2:
                    st.write("\n")
                    incluir_cid_atestado = st.checkbox("Incluir CID-10 e campo de autorização do paciente", value=True)

                if incluir_cid_atestado:
                    col_b1, col_b2 = st.columns([3, 1])
                    with col_b1:
                        diag_busca = st.text_input("Buscar diagnóstico (Ex: Fratura de Fêmur, Torção de Tornozelo):", placeholder="Digite para pesquisar na IA...").strip()
                    with col_b2:
                        st.write("\n\n")
                        if st.button("🔍 Buscar CID", use_container_width=True):
                            if diag_busca:
                                with st.spinner("Buscando CIDs ortopédicos..."):
                                    res = consultar_cid_ia(diag_busca)
                                    if res:
                                        st.session_state['lista_opcoes_ia'] = res
                                    else:
                                        st.error("Nenhum CID localizado.")

                    if st.session_state['lista_opcoes_ia']:
                        labels = [opt['label_exibicao'] for opt in st.session_state['lista_opcoes_ia']]
                        sel = st.selectbox("🎯 Escolha a patologia exata para preencher o CID-10:", labels)
                        for opt in st.session_state['lista_opcoes_ia']:
                            if opt['label_exibicao'] == sel:
                                st.session_state['cid_selecionado'] = f"{opt['cid_codigo']} - {opt['cid_descricao']}"

                    cid_final = st.text_input("CID-10 Final:", value=st.session_state['cid_selecionado']).strip().upper()

            # Exibe se for Ficha de Retorno
            retorno = 0
            if emitir_retorno:
                retorno = st.number_input("Dias até o retorno:", min_value=0, value=14, step=1)

            # Exibe se for Imobilização
            tipo_imobilizacao = ""
            hd_imobilizacao = ""
            if emitir_imobilizacao:
                col_im1, col_im2 = st.columns(2)
                with col_im1:
                    tipo_imobilizacao = st.text_input("Tipo de imobilização (Ex: TALA GESSO MALARES E PALMAR):").strip().upper()
                with col_im2:
                    hd_imobilizacao = st.text_input("Hipótese diagnóstica (HD):", value=cid_final).strip().upper()

            # Exibe se for Fisioterapia
            sessoes_fisio = 0
            quadro_clinico_fisio = ""
            if emitir_fisio:
                sessoes_fisio = st.number_input("Quantidade de sessões de fisioterapia motora:", min_value=1, value=20, step=1)
                quadro_clinico_fisio = st.text_area("Quadro clínico para fisioterapia:").strip().upper()

            # Exibe se for Antibiograma
            material_exame = ""
            quadro_clinico_antibiograma = ""
            usa_antibiotico = "NÃO"
            antibiotico_nome_dias = ""
            if emitir_antibio:
                material_exame = st.text_input("Material coletado (Antibiograma):").strip().upper()
                quadro_clinico_antibiograma = st.text_area("Quadro clínico (Antibiograma):").strip().upper()
                col_ab1, col_ab2 = st.columns(2)
                with col_ab1:
                    usa_antibiotico = st.selectbox("Faz uso de antibiótico?", ["NÃO", "SIM"])
                with col_ab2:
                    if usa_antibiotico == "SIM":
                        antibiotico_nome_dias = st.text_input("Antibiótico(s) e dias de tratamento:").strip().upper()

            st.write("\n")
            if st.button("📦 Gerar Documentos Ortopédicos (.DOCX)", type="primary"):
                if not nome:
                    st.error("Por favor, preencha o Nome do paciente.")
                else:
                    hoje_str = datetime.now().strftime('%d/%m/%Y')
                    
                    data_nasc_txt = data_nascimento if data_nascimento else "_________________________"
                    nome_mae_txt = nome_mae if nome_mae else "__________________________________________________"
                    cid_txt = cid_final if cid_final else "__________________________________________________"
                    
                    documentos_gerar = []

                    if emitir_atestado:
                        corpo_atestado = (
                            f"Atesto, para os devidos fins, que o(a) Sr(a) {nome} recebeu atendimento neste serviço no dia ____/____/______, "
                            f"e necessita afastamento de suas atividades por {afastamento} dias, a partir desta data.\n\n"
                        )
                        if incluir_cid_atestado:
                            corpo_atestado += (
                                f"CID-10: {cid_txt}\n\n"
                                f"_____________________________\nAssinatura do médico\n\n"
                                f"* Autorizo a divulgação do diagnóstico (CID)\n\n"
                                f"_____________________________________\nPaciente ou responsável legal"
                            )
                        else:
                            corpo_atestado += f"_____________________________\nAssinatura do médico"
                        
                        documentos_gerar.append(("ATESTADO MÉDICO", corpo_atestado))

                    if emitir_retorno:
                        corpo_retorno = (
                            f"Retorno ao Ambulatório de Ortopedia e Traumatologia em {retorno} dias para reavaliação clínica e evolução do tratamento.\n\n"
                            f"Orientações:\n"
                            f"• Realizar raio-x de retorno, se houver.\n"
                            f"• Utilizar as medicações prescritas.\n"
                            f"• Procurar atendimento de urgência em caso de febre, dor, vermelhidão, secreção e/ou calor local, abertura dos pontos ou se exposição de placas/pinos/parafusos, se houver.\n\n\n"
                            f"_____________________________\nAssinatura do médico"
                        )
                        documentos_gerar.append(("ENCAMINHAMENTO PARA RETORNO – AMBULATÓRIO DE ORTOPEDIA E TRAUMATOLOGIA", corpo_retorno))

                    if emitir_receita:
                        corpo_receita = "\n\n\n\n\n\n\n\n\n\n_____________________________\nAssinatura do médico"
                        documentos_gerar.append(("RECEITUÁRIO COMUM", corpo_receita))

                    if emitir_imobilizacao:
                        corpo_imobilizacao = (
                            f"Nome do paciente: {nome}\n\n"
                            f"SOLICITO:\n\n"
                            f"IMOBILIZAÇÃO TIPO: {tipo_imobilizacao}\n\n"
                            f"HD: {hd_imobilizacao}\n\n\n\n"
                            f"_____________________________\nAssinatura do médico"
                        )
                        documentos_gerar.append(("SOLICITAÇÃO DE IMOBILIZAÇÃO", corpo_imobilizacao))

                    if emitir_fisio:
                        corpo_fisio = (
                            f"Nome do paciente: {nome}\n\n"
                            f"SOLICITO:\n\n"
                            f"{sessoes_fisio} sessões de fisioterapia motora.\n\n"
                            f"Quadro clínico: {quadro_clinico_fisio}\n\n\n\n"
                            f"_____________________________\nAssinatura do médico"
                        )
                        documentos_gerar.append(("FICHA DE ENCAMINHAMENTO PARA FISIOTERAPIA", corpo_fisio))

                    if emitir_antibio:
                        corpo_antibio = (
                            f"NOME DO PACIENTE: {nome}\n"
                            f"DATA DE NASCIMENTO: {data_nasc_txt}\n"
                            f"NOME DA MÃE OU PAI: {nome_mae_txt}\n\n"
                            f"MATERIAL: {material_exame}\n"
                            f"ANÁLISE: Antibiograma\n"
                            f"QUADRO CLÍNICO: {quadro_clinico_antibiograma}\n"
                            f"FAZ USO DE ANTIBIÓTICO? {usa_antibiotico}\n"
                        )
                        if usa_antibiotico == "SIM":
                            corpo_antibio += f"ANTIBIÓTICO(S) E DIAS: {antibiotico_nome_dias}\n\n\n"
                        else:
                            corpo_antibio += "\n\n\n"
                        corpo_antibio += "_____________________________\nAssinatura do médico"
                        documentos_gerar.append(("SOLICITAÇÃO DE ANTIBIOGRAMA", corpo_antibio))

                    # GERADOR WORD (.DOCX)
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

                    for i, (titulo, corpo) in enumerate(documentos_gerar):
                        # TÍTULO CENTRALIZADO
                        p_titulo = doc_word.add_paragraph()
                        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_titulo.paragraph_format.space_after = Pt(18)
                        run_tit = p_titulo.add_run(titulo)
                        run_tit.font.name = 'Arial'
                        run_tit.font.size = Pt(13)
                        run_tit.bold = True

                        # CORPO JUSTIFICADO
                        for linha in corpo.split('\n'):
                            if linha.strip() == "":
                                doc_word.add_paragraph().paragraph_format.space_after = Pt(4)
                                continue
                            p_linha = doc_word.add_paragraph()
                            p_linha.paragraph_format.line_spacing = 1.2
                            p_linha.paragraph_format.space_after = Pt(4)
                            p_linha.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                
                            run_linha = p_linha.add_run(linha)
                            run_linha.font.name = 'Arial'
                            run_linha.font.size = Pt(11)

                        # DATA E CIDADE NO RODAPÉ À DIREITA
                        p_data = doc_word.add_paragraph()
                        p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        p_data.paragraph_format.space_before = Pt(20)
                        run_data = p_data.add_run(f"{hoje_str}, Campo Maior - PI")
                        run_data.font.name = 'Arial'
                        run_data.font.size = Pt(10)

                        if i < len(documentos_gerar) - 1:
                            doc_word.add_page_break()

                    bio = io.BytesIO()
                    doc_word.save(bio)
                    
                    st.success("✨ Documentos de Ortopedia gerados com sucesso!")
                    st.download_button(
                        label="📥 Baixar Documentos .DOCX",
                        data=bio.getvalue(),
                        file_name=f"Documentos_Ortopedia_{nome.lower().replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

elif senha != "":
    st.error("Chave de acesso inválida.")
