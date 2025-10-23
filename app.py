import streamlit as st
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# === Estrutura de pastas ===
PASTA_BLOCOS = "blocos"
PASTA_SAIDAS = "saidas"
os.makedirs(PASTA_SAIDAS, exist_ok=True)

# === Função para ler bloco ===
def ler_bloco(nome_arquivo):
    caminho = os.path.join(PASTA_BLOCOS, nome_arquivo)
    if not os.path.exists(caminho):
        return f"⚠️ [Arquivo ausente: {nome_arquivo}]"
    with open(caminho, "r", encoding="utf-8") as f:
        return f.read().strip()

# === Processar marcação [LARANJA] ===
def processar_laranja(texto, paragrafo):
    partes = texto.split("[LARANJA]")
    for i, parte in enumerate(partes):
        if i % 2 == 0:
            paragrafo.add_run(parte)
        else:
            laranja, *resto = parte.split("[/LARANJA]")
            run = paragrafo.add_run(laranja)
            try:
                run.font.highlight_color = 6  # cor laranja
            except Exception:
                pass
            if resto:
                paragrafo.add_run(resto[0])

# === Salvar DOCX ===
def salvar_peticao(texto_final, nome_arquivo="peticao_final.docx"):
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11.5)
    try:
        estilo._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    except Exception:
        pass

    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    for bloco in texto_final.split("\n\n"):
        if not bloco.strip():
            continue
        if "[PARAGRAFO]" in bloco:
            doc.add_paragraph()
            continue

        recuo_completo = "[RECUO_COMPLETO]" in bloco
        sem_recuo = "[SEM_RECUO]" in bloco
        centralizado = "[CENTRALIZADO]" in bloco
        bloco = (
            bloco.replace("[RECUO_COMPLETO]", "")
            .replace("[SEM_RECUO]", "")
            .replace("[CENTRALIZADO]", "")
        )

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        if centralizado:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif recuo_completo:
            p.paragraph_format.left_indent = Cm(4)
            p.paragraph_format.first_line_indent = Cm(0)
        elif sem_recuo:
            p.paragraph_format.first_line_indent = Cm(0)
        else:
            p.paragraph_format.first_line_indent = Cm(4)

        partes = bloco.split("[NEGRITO]")
        for i, parte in enumerate(partes):
            if i % 2 == 0:
                processar_laranja(parte, p)
            else:
                negrito, *resto = parte.split("[/NEGRITO]")
                run = p.add_run(negrito)
                run.bold = True
                if resto:
                    processar_laranja(resto[0], p)

    caminho_saida = os.path.join(PASTA_SAIDAS, nome_arquivo)
    doc.save(caminho_saida)
    return caminho_saida

# === Montar texto ===
def montar_texto(dados):
    texto = ""
    texto += ler_bloco("bloco1_comarca.txt").replace("[INSERIR COMARCA]", dados["comarca"]) + "\n\n"
    texto += ler_bloco("bloco2_qualificacao_completa.txt").replace("[INSERIR QUALIFICAÇÃO INFORMADA]", dados["requerente"]) + "\n\n"
    texto += ler_bloco(f"bloco3_plano_{dados['plano']}.txt") + "\n\n"

    if dados["prioridade"] != "NENHUMA":
        bloco = ler_bloco(f"bloco4_prioridade_{dados['prioridade']}.txt")
        bloco = bloco.replace("[DESCREVER DOENÇA]", dados["doenca"])
        texto += bloco + "\n\n"

    texto += ler_bloco(f"bloco5_gratuidade_{dados['gratuidade']}.txt") + "\n\n"
    texto += ler_bloco("bloco6_fatos.txt").replace("[DESCREVER DOENÇA]", dados["doenca"]) + "\n\n"
    texto += ler_bloco(f"bloco7_negativa_{dados['negativa']}.txt") + "\n\n"
    texto += ler_bloco("bloco8_cdc.txt") + "\n\n"

    bloco9 = ler_bloco(f"bloco9_tipo_{dados['tipo_demanda']}.txt")
    bloco9 = bloco9.replace("[DESCREVER DOENÇA]", dados["doenca"])
    texto += bloco9 + "\n\n"

    texto += ler_bloco(f"bloco10_urgencia_{dados['urgencia_tipo']}.txt").replace("[DESCRIÇÃO URGÊNCIA]", dados["urgencia"]) + "\n\n"
    texto += ler_bloco(f"bloco11_pedidos_{dados['pedido']}.txt")

    return texto

# === INTERFACE WEB (STREAMLIT) ===
st.set_page_config(page_title="🧩 Gerador de Petições", layout="centered")
st.title("🧩 Gerador de Petições")
st.caption("Versão web — uso restrito.")
st.markdown("---")

comarca = st.text_input("Comarca:")
requerente = st.text_input("REQUERENTE, qualificação completa:")
plano = st.selectbox("Plano de Saúde:", ["unimed", "bradesco", "notredame", "samaritano", "amil", "sulamerica"])
prioridade = st.selectbox("Prioridade de Tramitação:", ["NENHUMA", "idoso", "deficiente"])
gratuidade = st.selectbox("Gratuidade de Justiça:", ["NENHUMA", "idoso_ou_tutelado", "menor"])
doenca = st.text_input("Doença / Condição:")
negativa = st.selectbox("Tipo de Negativa:", ["tacita", "outra"])
tipo_demanda = st.selectbox(
    "Tipo de Demanda:",
    ["deficiencia_clinico", "deficiencia_domiciliar", "idoso_clinico", "idoso_domiciliar", "outros"]
)
urgencia = st.text_area("Urgência (parágrafo completo):")
urgencia_tipo = st.selectbox("Tipo de Urgência:", ["clinica", "domiciliar"])
pedido = st.selectbox("Tipo de Pedido:", ["clinica", "domiciliar"])

if st.button("🧩 Gerar Petição"):
    if not (comarca and requerente and plano and doenca and tipo_demanda and pedido):
        st.warning("⚠️ Preencha todos os campos obrigatórios antes de gerar a petição.")
    else:
        dados = {
            "comarca": comarca,
            "requerente": requerente,
            "plano": plano,
            "prioridade": prioridade,
            "gratuidade": gratuidade,
            "doenca": doenca,
            "negativa": negativa,
            "tipo_demanda": tipo_demanda,
            "urgencia": urgencia,
            "urgencia_tipo": urgencia_tipo,
            "pedido": pedido
        }
        texto_final = montar_texto(dados)
        nome_arquivo = f"Peticao_{requerente.replace(' ', '_')}.docx"
        caminho = salvar_peticao(texto_final, nome_arquivo)

        with open(caminho, "rb") as f:
            st.download_button(
                label="📄 Baixar Petição",
                data=f,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        st.success("✅ Petição gerada com sucesso! O formato é idêntico ao modelo original.")


