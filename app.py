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
        
        st.subheader("1. Identificação do Paciente")
        nome = st.text_input("Nome do Paciente [nome do paciente]:").strip().upper()
        
        col_id1, col_id2 = st.columns(2)
        with col_id1:
            data_nascimento = st.text_input("Data de Nascimento [DATA DE NASCIMENTO] (Ex: 12/05/1985):").strip()
        with col_id2:
            nome_mae = st.text_input("Nome da Mãe [NOME DA MÃE]:").strip().upper()
            
        data_admissao = st.text_input("Data de Admissão / Atendimento [data de admissão]:", value=datetime.now().strftime('%d/%m/%Y')).strip()

        st.subheader("2. Diagnóstico / CID-10 [código e descrição]")
        col_b1, col_b2 = st.columns([3, 1])
        with col_b1:
            diag_busca = st.text_input("Buscar Diagnóstico (Ex: Apendicite, Hérnia Inguinal):", placeholder="Digite para pesquisar na IA...").strip()
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

        cid_final = st.text_input("CID-10 Final [código e descrição]:", value=st.session_state['cid_selecionado']).strip().upper()

        st.subheader("3. Parâmetros de Prazos")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            afastamento = st.number_input("Dias de Afastamento [dias de afastamento]:", min_value=0, value=15, step=1)
            incluir_cid_atestado = st.checkbox("Incluir linha de CID-10 e assinatura do paciente no Atestado", value=True)
        with col_p2:
            retorno = st.number_input("Dias até o Retorno [dias até o retorno]:", min_value=0, value=14, step=1)

        st.subheader("4. Campos Específicos (Histopatológico / Antibiograma)")
        material_exame = st.text_input("Material Coletado:").strip().upper()
        quadro_clinico = st.text_area("Quadro Clínico:").strip().upper()
        
        col_ab1, col_ab2 = st.columns(2)
        with col_ab1:
            usa_antibiotico = st.selectbox("Faz uso de Antibiótico?", ["NÃO", "SIM"])
        with col_ab2:
            antibiotico_nome_dias = st.text_input("Se sim, Antibiótico(s) e Dias de Tratamento:").strip().upper()

        st.subheader("5. Selecionar Documentos para Gerar")
        emitir_atestado = st.checkbox("Atestado Médico (Cirurgia Geral)", value=True)
        emitir_retorno = st.checkbox("Ficha de Retorno (Ambulatório de Cirurgia Geral)", value=True)
        emitir_receita = st.checkbox("Receituário Comum", value=False)
        emitir_histo = st.checkbox("Solicitação de Histopatológico", value=False)
        emitir_antibio = st.checkbox("Solicitação de Antibiograma", value=False)

        if st.button("📦 Gerar Lote de Documentos (.DOCX)", type="primary"):
            if not nome:
                st.error("Por favor, preencha o Nome do Paciente.")
            else:
                hoje_str = datetime.now().strftime('%d/%m/%Y')
                
                # Tratamento para campos vazios de preenchimento manual
                data_nasc_txt = data_nascimento if data_nascimento else "_________________________"
                nome_mae_txt = nome_mae if nome_mae else "__________________________________________________"
                cid_txt = cid_final if cid_final else "__________________________________________________"
                
                paginas = []

                if emitir_atestado:
                    if incluir_cid_atestado:
                        txt_atestado = (
                            f"ATESTADO MÉDICO\n\n"
                            f"Atesto, para os devidos fins, que o(a) Sr(a) {nome} recebeu atendimento neste serviço no dia {data_admissao}, e necessita afastamento de suas atividades por {afastamento} dias, a partir de {data_admissao}.\n"
                            f"CID-10: {cid_txt}\n"
                            f"{hoje_str}, Campo Maior - PI\n\n"
                            f"_____________________________\nAssinatura do médico\n\n"
                            f"* Autorizo a divulgação do diagnóstico (CID)\n\n\n"
                            f"_____________________________________\nPaciente ou responsável legal"
                        )
                    else:
                        txt_atestado = (
                            f"ATESTADO MÉDICO\n\n"
                            f"Atesto, para os devidos fins, que o(a) Sr(a) {nome} recebeu atendimento neste serviço no dia {data_admissao}, e necessita afastamento de suas atividades por {afastamento} dias, a partir de {data_admissao}.\n"
                            f"{hoje_str}, Campo Maior - PI\n\n"
                            f"_____________________________\nAssinatura do médico"
                        )
                    paginas.append(("ATESTADO MÉDICO", txt_atestado))

                if emitir_retorno:
                    txt_retorno = (
                        f"ENCAMINHAMENTO PARA RETORNO – AMBULATÓRIO DE CIRURGIA GERAL\n\n"
                        f"Retorno ao Ambulatório de Cirurgia Geral em {retorno} dias para reavaliação pós-operatória.\n\n"
                        f"Orientações:\n\n"
                        f"Manter curativo limpo e seco.\n"
                        f"Utilizar as medicações prescritas.\n"
                        f"Evitar esforços físicos até nova avaliação.\n"
                        f"Manter alimentação e hidratação adequadas.\n"
                        f"Procurar atendimento de urgência em caso de febre, dor, vômitos, vermelhidão intensa, secreção pela ferida, sangramento ou abertura dos pontos.\n\n"
                        f"{hoje_str}, Campo Maior - PI\n\n"
                        f"_____________________________\nAssinatura do médico"
                    )
                    paginas.append(("ENCAMINHAMENTO PARA RETORNO", txt_retorno))

                if emitir_receita:
                    txt_receita = (
                        f"RECEITUÁRIO COMUM\n\n\n\n\n\n\n\n\n"
                        f"{hoje_str}, Campo Maior - PI\n\n"
                        f"_____________________________\nAssinatura do médico"
                    )
                    paginas.append(("RECEITUÁRIO COMUM", txt_receita))

                if emitir_histo:
                    txt_histo = (
                        f"SOLICITAÇÃO DE HISTOPATOLÓGICO\n\n"
                        f"NOME DO PACIENTE: {nome}\n"
                        f"DATA DE NASCIMENTO: {data_nasc_txt}\n"
                        f"NOME DA MÃE: {nome_mae_txt}\n\n"
                        f"MATERIAL: {material_exame}\n"
                        f"ANÁLISE: Histopatológico\n"
                        f"QUADRO CLÍNICO: {quadro_clinico}\n\n\n\n"
                        f"{hoje_str}, Campo Maior - PI\n\n"
                        f"_____________________________\nAssinatura do médico"
                    )
                    paginas.append(("SOLICITAÇÃO DE HISTOPATOLÓGICO", txt_histo))

                if emitir_antibio:
                    txt_antibio = (
                        f"SOLICITAÇÃO DE ANTIBIOGRAMA\n\n"
                        f"NOME DO PACIENTE: {nome}\n"
                        f"DATA DE NASCIMENTO: {data_nasc_txt}\n"
                        f"NOME DA MÃE: {nome_mae_txt}\n\n"
                        f"MATERIAL: {material_exame}\n"
                        f"ANÁLISE: Antibiograma\n"
                        f"QUADRO CLÍNICO: {quadro_clinico}\n"
                        f"FAZ USO DE ANTIBIÓTICO? {usa_antibiotico}\n"
                    )
                    if usa_antibiotico == "SIM":
                        txt_antibio += f"Se sim:\nANTIBIÓTICO(S): {antibiotico_nome_dias}\n\n\n"
                    else:
                        txt_antibio += "\n\n\n"
                    txt_antibio += f"{hoje_str}, Campo Maior - PI\n\n_____________________________\nAssinatura do médico"
                    paginas.append(("SOLICITAÇÃO DE ANTIBIOGRAMA", txt_antibio))

                # GERAÇÃO DO WORD (.DOCX)
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

                for i, (titulo, corpo) in enumerate(paginas):
                    for linha in corpo.split('\n'):
                        if linha.strip() == "":
                            doc_word.add_paragraph().paragraph_format.space_after = Pt(4)
                            continue
                        p_linha = doc_word.add_paragraph()
                        p_linha.paragraph_format.line_spacing = 1.2
                        p_linha.paragraph_format.space_after = Pt(4)
                        p_linha.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            
                        run_linha = p_linha.add_run(linha)
                        run_linha.font.name = 'Arial'
                        run_linha.font.size = Pt(11)
                        if "ATESTADO MÉDICO" in linha or "ENCAMINHAMENTO" in linha or "RECEITUÁRIO" in linha or "SOLICITAÇÃO" in linha:
                            run_linha.bold = True
                    
                    if i < len(paginas) - 1:
                        doc_word.add_page_break()
                
                bio = io.BytesIO()
                doc_word.save(bio)
                
                st.success("✨ Documentos de Cirurgia Geral gerados com sucesso!")
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
        
        st.subheader("1. Identificação do Paciente")
        nome = st.text_input("Nome do Paciente [nome do paciente]:").strip().upper()
        
        col_id1, col_id2 = st.columns(2)
        with col_id1:
            data_nascimento = st.text_input("Data de Nascimento [DATA DE NASCIMENTO] (Ex: 12/05/1985):").strip()
        with col_id2:
            nome_mae = st.text_input("Nome da Mãe [NOME DA MÃE]:").strip().upper()
            
        data_admissao = st.text_input("Data de Admissão / Atendimento [data de admissão]:", value=datetime.now().strftime('%d/%m/%Y')).strip()

        st.subheader("2. Diagnóstico / CID-10 [código e descrição]")
        col_b1, col_b2 = st.columns([3, 1])
        with col_b1:
            diag_busca = st.text_input("Buscar Diagnóstico (Ex: Fratura de Fêmur, Torção):", placeholder="Digite para pesquisar na IA...").strip()
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

        cid_final = st.text_input("CID-10 Final [código e descrição]:", value=st.session_state['cid_selecionado']).strip().upper()

        st.subheader("3. Parâmetros de Prazos")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            afastamento = st.number_input("Dias de Afastamento [dias de afastamento]:", min_value=0, value=45, step=1)
            incluir_cid_atestado = st.checkbox("Incluir linha de CID-10 e assinatura do paciente no Atestado", value=True)
        with col_p2:
            retorno = st.number_input("Dias até o Retorno [dias até o retorno]:", min_value=0, value=14, step=1)

        st.subheader("4. Campos Específicos Ortopédicos")
        tipo_imobilizacao = st.text_input("Tipo de Imobilização:").strip().upper()
        hd_imobilizacao = st.text_input("Hipótese Diagnóstica (HD) para Imobilização:", value=cid_final).strip().upper()
        
        sessoes_fisio = st.number_input("Quantidade de Sessões de Fisioterapia Motora:", min_value=1, value=20, step=1)
        quadro_clinico_fisio = st.text_area("Quadro Clínico para Fisioterapia:").strip().upper()

        # Antibiograma
        material_exame = st.text_input("Material Coletado (Antibiograma):").strip().upper()
        quadro_clinico_antibiograma = st.text_area("Quadro Clínico (Antibiograma):").strip().upper()
        col_ab1, col_ab2 = st.columns(2)
        with col_ab1:
            usa_antibiotico = st.selectbox("Faz uso de Antibiótico?", ["NÃO", "SIM"])
        with col_ab2:
            antibiotico_nome_dias = st.text_input("Se sim, Antibiótico(s) e Dias de Tratamento:").strip().upper()

        st.subheader("5. Selecionar Documentos para Gerar")
        emitir_atestado = st.checkbox("Atestado Médico (Ortopedia)", value=True)
        emitir_retorno = st.checkbox("Ficha de Retorno (Ambulatório de Ortopedia)", value=True)
        emitir_receita = st.checkbox("Receituário Comum", value=False)
        emitir_imobilizacao = st.checkbox("Solicitação de Imobilização", value=False)
        emitir_fisio = st.checkbox("Solicitação de Fisioterapia", value=False)
        emitir_antibio = st.checkbox("Solicitação de Antibiograma", value=False)

        if st.button("📦 Gerar Lote de Documentos Ortopédicos (.DOCX)", type="primary"):
            if not nome:
                st.error("Por favor, preencha o Nome do Paciente.")
            else:
                hoje_str = datetime.now().strftime('%d/%m/%Y')
                
                # Tratamento para campos vazios de preenchimento manual
                data_nasc_txt = data_nascimento if data_nascimento else "_________________________"
                nome_mae_txt = nome_mae if nome_mae else "__________________________________________________"
                cid_txt = cid_final if cid_final else "__________________________________________________"
                
                paginas = []

                if emitir_atestado:
                    if incluir_cid_atestado:
                        txt_atestado = (
                            f"ATESTADO MÉDICO\n\n"
                            f"Atesto, para os devidos fins, que o(a) Sr(a) {nome} recebeu atendimento neste serviço no dia {data_admissao}, e necessita afastamento de suas atividades por {afastamento} dias, a partir de {data_admissao}.\n"
                            f"CID-10: {cid_txt}\n"
                            f"{hoje_str}, Campo Maior - PI\n\n"
                            f"_____________________________\nAssinatura do médico\n\n"
                            f"* Autorizo a divulgação do diagnóstico (CID)\n\n\n"
                            f"_____________________________________\nPaciente ou responsável legal"
                        )
                    else:
                        txt_atestado = (
                            f"ATESTADO MÉDICO\n\n"
                            f"Atesto, para os devidos fins, que o(a) Sr(a) {nome} recebeu atendimento neste serviço no dia {data_admissao}, e necessita afastamento de suas atividades por {afastamento} dias, a partir de {data_admissao}.\n"
                            f"{hoje_str}, Campo Maior - PI\n\n"
                            f"_____________________________\nAssinatura do médico"
                        )
                    paginas.append(("ATESTADO MÉDICO", txt_atestado))

                if emitir_retorno:
                    txt_retorno = (
                        f"ENCAMINHAMENTO PARA RETORNO – AMBULATÓRIO DE ORTOPEDIA E TRAUMATOLOGIA\n\n"
                        f"Retorno ao Ambulatório de Ortopedia e Traumatologia em {retorno} dias para reavaliação clínica e evolução do tratamento.\n\n"
                        f"Orientações:\n\n"
                        f"Realizar raio-x de retorno, se houver.\n"
                        f"Utilizar as medicações prescritas.\n"
                        f"Procurar atendimento de urgência em caso de febre, dor, vermelhidão, secreção e/ou calor local, abertura dos pontos ou se exposição de placas/pinos/parafusos, se houver.\n\n\n\n"
                        f"{hoje_str}, Campo Maior - PI\n\n"
                        f"_____________________________\nAssinatura do médico"
                    )
                    paginas.append(("ENCAMINHAMENTO PARA RETORNO", txt_retorno))

                if emitir_receita:
                    txt_receita = (
                        f"RECEITUÁRIO COMUM\n\n\n\n\n\n\n\n\n"
                        f"{hoje_str}, Campo Maior - PI\n\n"
                        f"_____________________________\nAssinatura do médico"
                    )
                    paginas.append(("RECEITUÁRIO COMUM", txt_receita))

                if emitir_imobilizacao:
                    txt_imobilizacao = (
                        f"SOLICITAÇÃO DE IMOBILIZAÇÃO\n\n"
                        f"Nome do Paciente: {nome}\n\n"
                        f"SOLICITO\n\n"
                        f"IMOBILIZAÇÃO TIPO {tipo_imobilizacao}\n\n"
                        f"HD: {hd_imobilizacao}\n\n\n\n"
                        f"{hoje_str}, Campo Maior - PI\n\n"
                        f"_____________________________\nAssinatura do médico"
                    )
                    paginas.append(("SOLICITAÇÃO DE IMOBILIZAÇÃO", txt_imobilizacao))

                if emitir_fisio:
                    txt_fisio = (
                        f"FICHA DE ENCAMINHAMENTO PARA FISIOTERAPIA\n\n"
                        f"Nome do Paciente: {nome}\n\n"
                        f"SOLICITO\n\n"
                        f"{sessoes_fisio} sessões de fisioterapia motora.\n\n"
                        f"Quadro clínico: {quadro_clinico_fisio}\n\n\n\n"
                        f"{hoje_str}, Campo Maior - PI\n\n"
                        f"_____________________________\nAssinatura do médico"
                    )
                    paginas.append(("FICHA DE ENCAMINHAMENTO PARA FISIOTERAPIA", txt_fisio))

                if emitir_antibio:
                    txt_antibio = (
                        f"SOLICITAÇÃO DE ANTIBIOGRAMA\n\n"
                        f"NOME DO PACIENTE: {nome}\n"
                        f"DATA DE NASCIMENTO: {data_nasc_txt}\n"
                        f"NOME DA MÃE: {nome_mae_txt}\n\n"
                        f"MATERIAL: {material_exame}\n"
                        f"ANÁLISE: Antibiograma\n"
                        f"QUADRO CLÍNICO: {quadro_clinico_antibiograma}\n"
                        f"FAZ USO DE ANTIBIÓTICO? {usa_antibiotico}\n"
                    )
                    if usa_antibiotico == "SIM":
                        txt_antibio += f"Se sim:\nANTIBIÓTICO(S): {antibiotico_nome_dias}\n\n\n"
                    else:
                        txt_antibio += "\n\n\n"
                    txt_antibio += f"{hoje_str}, Campo Maior - PI\n\n_____________________________\nAssinatura do médico"
                    paginas.append(("SOLICITAÇÃO DE ANTIBIOGRAMA", txt_antibio))

                # GERAÇÃO DO WORD (.DOCX)
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

                for i, (titulo, corpo) in enumerate(paginas):
                    for linha in corpo.split('\n'):
                        if linha.strip() == "":
                            doc_word.add_paragraph().paragraph_format.space_after = Pt(4)
                            continue
                        p_linha = doc_word.add_paragraph()
                        p_linha.paragraph_format.line_spacing = 1.2
                        p_linha.paragraph_format.space_after = Pt(4)
                        p_linha.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            
                        run_linha = p_linha.add_run(linha)
                        run_linha.font.name = 'Arial'
                        run_linha.font.size = Pt(11)
                        if "ATESTADO MÉDICO" in linha or "ENCAMINHAMENTO" in linha or "RECEITUÁRIO" in linha or "SOLICITAÇÃO" in linha or "FICHA DE ENCAMINHAMENTO" in linha:
                            run_linha.bold = True
                    
                    if i < len(paginas) - 1:
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
