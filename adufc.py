import fitz  # PyMuPDF
import re
import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import streamlit as st
import zipfile
import uuid
import time

# Configuração da página
st.set_page_config(page_title="Separador de PDF por Cliente")

st.title("🔍 Separador de PDF por Cliente")
plano_selecionado = st.selectbox("Selecione o plano", ["Unimed", "Uniodonto"])
email_file = st.file_uploader("Escolha o arquivo Excel com os e-mails", type="xlsx")
uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

def extrair_nome_titular(texto, plano):
    if plano == "Unimed":
        # Mesma lógica anterior para Unimed
        linhas = texto.splitlines()
        for i, linha in enumerate(linhas):
            if "Carteira:" in linha:
                if i > 0:
                    nome = linhas[i - 1].strip()
                    return re.sub(r"[^\w\sÀ-ÿ]", "", nome)
                break

    elif plano == "Uniodonto":
        # Primeiro ajusta quebras de linha no CPF
        texto_clean = re.sub(r"(\d)-\s*\n\s*(\d)", r"\1-\2", texto)
        # Busca o nome completo que vem junto ao CPF
        match = re.search(
            r"([A-Za-zÀ-ÿ\s-]+?)\s*-\s*\d{3}\.\d{3}\.\d{3}-\d{2}",
            texto_clean
        )
        if match:
            return match.group(1).strip()

    return "cliente_desconhecido"


def separar_por_cliente(pdf_path, plano):
    doc = fitz.open(pdf_path)
    arquivos_gerados, nome_cliente_atual, paginas_atual = [], None, []
    for i, pagina in enumerate(doc):
        texto = pagina.get_text()
        if plano == "Uniodonto" and "CLIENTE DO PLANO UNIMASTER-UNI" in texto:
            if paginas_atual:
                arquivos_gerados.append(salvar_pdf(doc, paginas_atual, nome_cliente_atual))
                paginas_atual = []
            nome_cliente_atual = extrair_nome_titular(texto, plano)
        elif plano == "Unimed" and "Prezado(a) Cliente" in texto:
            if paginas_atual:
                arquivos_gerados.append(salvar_pdf(doc, paginas_atual, nome_cliente_atual))
                paginas_atual = []
            nome_cliente_atual = extrair_nome_titular(texto, plano)
        if nome_cliente_atual:
            paginas_atual.append(i)
    if paginas_atual:
        arquivos_gerados.append(salvar_pdf(doc, paginas_atual, nome_cliente_atual))
    doc.close()
    return arquivos_gerados


def salvar_pdf(doc_original, lista_paginas, nome_arquivo_base):
    novo = fitz.open()
    for num in lista_paginas:
        novo.insert_pdf(doc_original, from_page=num, to_page=num)
    os.makedirs("arquivos_clientes", exist_ok=True)
    nome = os.path.join("arquivos_clientes", f"{nome_arquivo_base}.pdf")
    novo.save(nome)
    novo.close()
    return nome


def enviar_email(dest, nome_cliente, pdf):
    sender = os.getenv("EMAIL")
    pwd = os.getenv("EMAIL_PASSWORD")
    msg = MIMEMultipart()
    msg['From'], msg['To'] = sender, dest
    msg['Subject'] = f"ADUFC - UNIMED/UNIODONTO ({nome_cliente})"
    body = (
        "Prezado(a) Professor(a),\n\n"
        "Seguem em anexo os demonstrativos de suas despesas...\n\n"
        "Atenciosamente,\nSetor de Atendimento ao Docente"
    )
    msg.attach(MIMEText(body, 'plain'))
    with open(pdf, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(pdf))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf)}"'
        msg.attach(part)
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, dest, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)


def criar_zip(arquivos):
    nome = "arquivos_clientes.zip"
    with zipfile.ZipFile(nome, 'w') as z:
        for a in arquivos:
            z.write(a, os.path.basename(a))
    return nome

# --- Fluxo principal ---
if email_file and uploaded_file:
    # Salva PDF temporário
    with open("temp_input.pdf", "wb") as f:
        f.write(uploaded_file.read())

    # Debug: mostra texto bruto para verificação
    with st.expander("🔍 Debug - Texto bruto das páginas"):
        doc_debug = fitz.open("temp_input.pdf")
        for i, pagina in enumerate(doc_debug):
            texto = pagina.get_text()
            st.markdown(f"**Página {i}**")
            st.text_area(f"Texto página {i}", texto, height=200)
        doc_debug.close()

    df_emails = pd.read_excel(email_file)
    with st.spinner("Processando..."):
        arquivos = separar_por_cliente("temp_input.pdf", plano_selecionado)
    st.success(f"{len(arquivos)} PDFs gerados!")

    # Download dos PDFs
    with st.expander("🔍 Arquivos gerados"):
        for a in arquivos:
            nome = os.path.basename(a)
            st.write(f"- {nome}")
            st.download_button(
                f"Baixar {nome}",
                data=open(a, "rb").read(),
                file_name=nome,
                mime="application/pdf",
                key=str(uuid.uuid4())
            )

    # ZIP com todos PDFs
    zip_arquivo = criar_zip(arquivos)
    with open(zip_arquivo, "rb") as fzip:
        st.download_button(
            "📥 Baixar todos os PDFs (ZIP)",
            fzip,
            zip_arquivo,
            "application/zip"
        )

    # ... restante do fluxo de envio e relatórios ...

else:
    st.error("Faça upload do PDF e do Excel para prosseguir.")
