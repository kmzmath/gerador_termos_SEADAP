import streamlit as st
from docx import Document
from num2words import num2words
from decimal import Decimal, getcontext
import io
import re
from dateutil.parser import parse, ParserError
from babel.dates import format_date

st.set_page_config(layout="wide", page_title="Gerador de Termos FUNDOPEM")

getcontext().prec = 12

# --- FUNÇÕES AUXILIARES ---
def format_currency(num_str):
    try:
        num = Decimal(num_str.replace('.', '').replace(',', '.'))
        return f'{float(num):,.2f}'.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
    except (ValueError, TypeError):
        return num_str

def format_cnpj(cnpj_str):
    if not cnpj_str: return ""
    cleaned_cnpj = re.sub(r'\D', '', cnpj_str)
    if len(cleaned_cnpj) != 14: return cnpj_str
    return f"{cleaned_cnpj[0:2]}.{cleaned_cnpj[2:5]}.{cleaned_cnpj[5:8]}/{cleaned_cnpj[8:12]}-{cleaned_cnpj[12:14]}"

def format_cpf(cpf_str):
    if not cpf_str: return ""
    cleaned_cpf = re.sub(r'\D', '', cpf_str)
    if len(cleaned_cpf) != 11: return cpf_str
    return f"{cleaned_cpf[0:3]}.{cleaned_cpf[3:6]}.{cleaned_cpf[6:9]}-{cleaned_cpf[9:11]}"

def format_full_date(date_str):
    if not date_str: return ""
    try:
        dt = parse(date_str, dayfirst=True)
        return dt.strftime('%d.%m.%Y')
    except (ParserError, ValueError, TypeError):
        return date_str

def format_month_year(date_str):
    """Interpreta qualquer formato de data e retorna no padrão 'Mês de Ano'."""
    if not date_str: return ""
    try:
        dt = parse(date_str, dayfirst=True)
        return format_date(dt, 'MMMM de yyyy', locale='pt_BR').capitalize()
    except (ParserError, ValueError, TypeError):
        return date_str

def numero_para_extenso_completo(numero_str, tipo='uif'):
    try:
        numero_limpo = str(numero_str).replace('.', '')
        partes = numero_limpo.split(',')
        inteiro_str = partes[0]
        decimal_str = partes[1].ljust(2, '0') if len(partes) > 1 else '00'
        inteiro = int(inteiro_str)
        decimal = int(decimal_str)
        extenso_inteiro = num2words(inteiro, lang='pt_BR')
        if decimal > 0:
            extenso_decimal = num2words(decimal, lang='pt_BR')
            if tipo == 'uif':
                sufixo_inteiro = "inteiro" if inteiro == 1 else "inteiros"
                sufixo_decimal = "centésimo" if decimal == 1 else "centésimos"
                return f"{extenso_inteiro} {sufixo_inteiro}, e {extenso_decimal} {sufixo_decimal}"
            elif tipo == 'percent':
                return f"{extenso_inteiro} vírgula {extenso_decimal}"
        return extenso_inteiro
    except (ValueError, TypeError):
        return ""

def docx_replace(doc, replacements):
    for p in doc.paragraphs:
        if any(placeholder in p.text for placeholder in replacements):
            for placeholder, value in replacements.items():
                if placeholder in p.text:
                    while placeholder in p.text:
                        inline = p.runs
                        full_text = ''.join(r.text for r in inline)
                        if placeholder not in full_text:
                            break
                        
                        placeholder_start = full_text.find(placeholder)
                        placeholder_end = placeholder_start + len(placeholder)

                        start_run_idx, end_run_idx = -1, -1
                        current_pos = 0
                        for i, run in enumerate(inline):
                            run_len = len(run.text)
                            if start_run_idx == -1 and current_pos + run_len > placeholder_start:
                                start_run_idx = i
                            if end_run_idx == -1 and current_pos + run_len >= placeholder_end:
                                end_run_idx = i
                                break
                            current_pos += run_len
                        
                        if start_run_idx == -1 or end_run_idx == -1: continue

                        start_char_pos = placeholder_start - sum(len(r.text) for r in inline[:start_run_idx])
                        end_char_pos = placeholder_end - sum(len(r.text) for r in inline[:end_run_idx])
                        
                        if start_run_idx == end_run_idx:
                            run = inline[start_run_idx]
                            run.text = run.text[:start_char_pos] + str(value) + run.text[end_char_pos:]
                        else:
                            start_run = inline[start_run_idx]
                            start_run.text = start_run.text[:start_char_pos] + str(value)
                            
                            for i in range(start_run_idx + 1, end_run_idx):
                                inline[i].text = ""

                            end_run = inline[end_run_idx]
                            end_run.text = end_run.text[end_char_pos:]

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace(cell, replacements)

    return doc


# --- INTERFACE GRÁFICA (STREAMLIT) ---
st.title("📄 Gerador Automatizado de Termos de Ajuste")
st.markdown("Preencha os campos abaixo para gerar o documento.")

try:
    with open("Template_RECUPERA EXPRESS.docx", "rb"):
        st.success("✅ Template 'Template_RECUPERA EXPRESS.docx' encontrado!")
except FileNotFoundError:
    st.error("❌ ERRO: O arquivo 'Template_RECUPERA EXPRESS.docx' não foi encontrado.")
    st.stop()


# --- SEÇÃO 1: Informações Gerais ---
st.header("1. Informações Gerais do Termo")
col1, col2 = st.columns(2)
with col1:
    termo_num = st.text_input("Número do Termo", key="termo_num")
    empresa_nome = st.text_input("Nome da Empresa", key="empresa_nome")
    empresa_cnpj = st.text_input("CNPJ", key="empresa_cnpj")
    representante_nome = st.text_input("Nome do Representante Legal", key="representante_nome")
    representante_cpf = st.text_input("CPF do Representante", key="representante_cpf")
    proa_num = st.text_input("Nº do Processo (PROA)", "nº #proa#, de 03 de setembro de 2024", key="proa_num")

with col2:
    empresa_endereco = st.text_input("Endereço Completo", key="empresa_endereco")
    data_inicio = st.text_input("Início da Vigência (ex: 01/07/2025)", key="data_inicio")
    st.subheader("Dados do Parecer")
    parecer_num = st.text_input("Nº do Parecer (ex: 024/2025)", key="parecer_num")
    parecer_data = st.text_input("Data do Parecer (ex: 04.04.2025)", key="parecer_data")
    doe_data = st.text_input("Data do DOE (ex: 11.04.2025)", key="doe_data")


# --- SEÇÃO 2: Detalhes do Projeto ---
st.header("2. Detalhes do Projeto")
col3, col4 = st.columns(2)
with col3:
    empresa_porte = st.selectbox("Porte da Empresa", ["Pequeno", "Médio", "Grande"], index=1, key="empresa_porte")
    municipio = st.text_input("Município", key="municipio")
    corede = st.text_input("COREDE", key="corede")
    qtd_empregos = st.text_input("Quantidade de Empregos", key="qtd_empregos")
    empresa_cgcte = st.text_input("Inscrição Estadual (CGC/TE)", key="empresa_cgcte")
with col4:
    pontos_fundopem = st.text_input("Pontos FUNDOPEM (3.1)", key="pontos_fundopem")
    perc_integrar_str = st.text_input("Percentual INTEGRAR (3.2)", key="perc_integrar_str")
    pontos_idese = st.text_input("Pontos IDESE (Tabela 3.2)", key="pontos_idese")
    pontos_setor = st.text_input("Pontos Setor Industrial (Tabela 3.2)", key="pontos_setor")


# --- SEÇÃO 3: Valores, Prazos e Cálculos ---
st.header("3. Valores, Prazos e Cálculos em Tempo Real")
st.info("Para valores, use vírgula como separador decimal. Para datas, use qualquer formato comum.")
col5, col6 = st.columns(2)
with col5:
    st.subheader("Valores de Entrada (UIF/RS)")
    valor_total_str = st.text_input("Valor Total do Projeto (2.1)", key="valor_total_str")
    valor_apres_inicial_str = st.text_input("Valor Apresentado Inicialmente (2.3)", key="valor_apres_inicial_str")
    valor_inicial_aceito_str = st.text_input("Valor Inicialmente Aceito (2.3.1)", key="valor_inicial_aceito_str")
    equips_2_4_str = st.text_input("Equipamentos (2.4)", key="equips_2_4_str")
    
    st.subheader("Prazos")
    data_final_fruicao = st.text_input("Final da Fruição (ex: 31/12/2032)", key="data_final_fruicao")
    mes_regularidade = st.text_input("Mês da Regularidade (ex: Julho de 2025)", key="mes_regularidade")

with col6:
    st.subheader("Resultados Calculados (Atualizado ao vivo)")
    val_total_for_calc = st.session_state.get('valor_total_str', '0').replace('.', '').replace(',', '.')
    perc_integrar_for_calc = st.session_state.get('perc_integrar_str', '0').replace(',', '.')
    val_inicial_aceito_for_calc = st.session_state.get('valor_inicial_aceito_str', '0').replace('.', '').replace(',', '.')
    
    try:
        limite_max_calc = Decimal(val_total_for_calc) * (Decimal(perc_integrar_for_calc) / Decimal(100))
        valor_liberado_calc = Decimal('0.8') * Decimal(val_inicial_aceito_for_calc) * (Decimal(perc_integrar_for_calc) / Decimal(100))
        
        st.metric(label="Limite Máximo para ser Liberado (4.1.2)", value=f"R$ {format_currency(str(limite_max_calc))}")
        st.metric(label="Valor Liberado para Fruição (4.1.2.1)", value=f"R$ {format_currency(str(valor_liberado_calc))}")
    except Exception:
        st.warning("Preencha os valores de entrada e percentuais para ver os resultados calculados.")

st.divider()

# --- BOTÃO DE GERAÇÃO E LÓGICA FINAL ---
if st.button("🚀 Gerar Documento Word", type="primary", use_container_width=True):
    with st.spinner("Processando... Gerando seu documento, por favor aguarde."):
        
        replacements = {}
        key_to_placeholder = {
            'termo_num': '#xx/aaaa#', 'empresa_nome': '#EMPRESA#', 'empresa_endereco': '#ENDERECO#',
            'empresa_cnpj': '#XX.XXX.XXX/0001-XX#', 'representante_nome': '#REPRESENTANTE#',
            'proa_num': 'nº #proa#, de 03 de setembro de 2024', 'municipio': '#cidade#', 'corede': '#corede#',
            'pontos_idese': '#idese#', 'pontos_setor': '#setint#', 'empresa_porte': '#porte#',
            'empresa_cgcte': '#cgcte#', 'mes_regularidade': '#regularidade#',
        }

        for key, placeholder in key_to_placeholder.items():
            value = st.session_state.get(key)
            if value:
                if key == 'empresa_nome': value = str(value).upper()
                elif key == 'empresa_cnpj': value = format_cnpj(str(value))
                elif key == 'proa_num' and value == "nº #proa#, de 03 de setembro de 2024": continue
                replacements[placeholder] = value
        
        if st.session_state.representante_cpf:
            replacements['CPF #XXX.XXX.XXX-XX#'] = f"CPF {format_cpf(st.session_state.representante_cpf)}"
        
        if st.session_state.data_inicio:
            replacements['#inicio#'] = format_full_date(st.session_state.data_inicio)
        if st.session_state.data_final_fruicao:
            final_fruicao_fmt = format_month_year(st.session_state.data_final_fruicao)
            replacements['#final#'] = final_fruicao_fmt
            replacements['#final2#'] = final_fruicao_fmt.replace(' de ', '/')
        
        if st.session_state.parecer_num and st.session_state.parecer_data and st.session_state.doe_data:
            texto_parecer_final = f"Nº {st.session_state.parecer_num}, DE {st.session_state.parecer_data} (DOE de {st.session_state.doe_data})"
            replacements["Parecer nº xxx/aaaa, de dd.mm.aaaa (DOE de dd.mm.aaaa)"] = f"Parecer {texto_parecer_final}"
        
        if st.session_state.qtd_empregos:
            num_empregos = st.session_state.qtd_empregos
            replacements['#emp#'] = num_empregos
            replacements['(duzentos e noventa e sete)'] = f"({numero_para_extenso_completo(num_empregos, 'geral')})"

        if st.session_state.pontos_fundopem:
            pontos = st.session_state.pontos_fundopem
            replacements['#pontos# (sessenta) pontos'] = f"{pontos} ({numero_para_extenso_completo(pontos, 'geral')}) pontos"
            replacements['#pontos#'] = pontos
        
        if st.session_state.perc_integrar_str:
            perc = st.session_state.perc_integrar_str
            replacements['#integrar#%'] = f"{perc}%"
            replacements['#integrar#'] = perc
            integrar_extenso = numero_para_extenso_completo(perc, 'percent') + " por cento"
            replacements['(trinta e quatro vírgula cinquenta e cinco por cento)'] = f"({integrar_extenso})"
        
        if st.session_state.valor_total_str:
             replacements['693.224,32 UIF/RS (seiscentos e noventa e três mil, duzentos e vinte e quatro inteiros, e trinta e dois centésimos de Unidades de Incentivo do FUNDOPEM/RS)'] = f"{format_currency(st.session_state.valor_total_str)} UIF/RS ({numero_para_extenso_completo(format_currency(st.session_state.valor_total_str))} de Unidades de Incentivo do FUNDOPEM/RS)"
        if st.session_state.valor_apres_inicial_str:
            replacements['193.874,41 UIF/RS (cento e noventa e três mil, oitocentos e setenta e quatro inteiros, e quarenta e um centésimos de Unidades de Incentivo do FUNDOPEM/RS)'] = f"{format_currency(st.session_state.valor_apres_inicial_str)} UIF/RS ({numero_para_extenso_completo(format_currency(st.session_state.valor_apres_inicial_str))} de Unidades de Incentivo do FUNDOPEM/RS)"
        if st.session_state.valor_inicial_aceito_str:
            replacements['159.123,22 UIF/RS (cento e cinquenta e nove mil, cento e vinte e três inteiros, e vinte e dois centésimos de Unidades de Incentivo do FUNDOPEM/RS)'] = f"{format_currency(st.session_state.valor_inicial_aceito_str)} UIF/RS ({numero_para_extenso_completo(format_currency(st.session_state.valor_inicial_aceito_str))} de Unidades de Incentivo do FUNDOPEM/RS)"
        if st.session_state.equips_2_4_str:
            placeholder_2_4 = "Do valor estabelecido no item 2.3.1 desta Cláusula, o montante de 193.874,41 UIF/RS (cento e noventa e três mil, oitocentos e setenta e quatro inteiros, e quarenta e um centésimos de Unidades de Incentivo do FUNDOPEM/RS) contempla os investimentos realizados em equipamentos."
            novo_valor_2_4 = f"Do valor estabelecido no item 2.3.1 desta Cláusula, o montante de {format_currency(st.session_state.equips_2_4_str)} UIF/RS ({numero_para_extenso_completo(format_currency(st.session_state.equips_2_4_str))} de Unidades de Incentivo do FUNDOPEM/RS) contempla os investimentos realizados em equipamentos."
            replacements[placeholder_2_4] = novo_valor_2_4
        
        if st.session_state.perc_integrar_str and st.session_state.valor_total_str:
            limite_max_liberado_calc = Decimal(val_total_for_calc) * (Decimal(perc_integrar_for_calc) / Decimal(100))
            placeholder_lim_max = '239.509,00 UIF/RS (duzentos e trinta e nove mil, e quinhentos e nove inteiros de Unidades de Incentivo do FUNDOPEM/RS)'
            novo_valor_lim_max = f"{format_currency(str(limite_max_liberado_calc))} UIF/RS ({numero_para_extenso_completo(format_currency(str(limite_max_liberado_calc)))} de Unidades de Incentivo do FUNDOPEM/RS)"
            replacements[placeholder_lim_max] = novo_valor_lim_max
        if st.session_state.perc_integrar_str and st.session_state.valor_inicial_aceito_str:
            valor_liberado_fruicao_calc = Decimal('0.8') * Decimal(val_inicial_aceito_for_calc) * (Decimal(perc_integrar_for_calc) / Decimal(100))
            placeholder_val_lib = '62.299,92 UIF/RS (sessenta e dois mil, duzentos e noventa e nove inteiros, e noventa e dois centésimos de Unidades de Incentivo do FUNDOPEM/RS)'
            novo_valor_val_lib = f"{format_currency(str(valor_liberado_fruicao_calc))} UIF/RS ({numero_para_extenso_completo(format_currency(str(valor_liberado_fruicao_calc)))} de Unidades de Incentivo do FUNDOPEM/RS)"
            replacements[placeholder_val_lib] = novo_valor_val_lib
        
        doc = Document("Template_RECUPERA EXPRESS.docx")
        doc = docx_replace(doc, replacements)
        
        bio = io.BytesIO()
        doc.save(bio)
        
        nome_arquivo = f"Termo_Ajuste_{st.session_state.get('termo_num', 'novo').replace('/', '-')}_{st.session_state.get('empresa_nome', 'empresa').replace(' ', '_')}.docx"
        
        st.success("🎉 Documento gerado com sucesso!")
        st.download_button(
            label="Clique aqui para baixar o arquivo Word",
            data=bio.getvalue(),
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )