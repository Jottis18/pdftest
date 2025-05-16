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
import uuid  # já coloca no topo se ainda não tiver

st.set_page_config(page_title="Separador de PDF por Cliente")

st.title("🔍 Separador de PDF por Cliente")
st.write("Faça upload do PDF")

# Opção para escolher o plano (Unimed ou Uniodonto)
plano_selecionado = st.selectbox("Selecione o plano", ["Unimed", "Uniodonto"])

# Carregar o arquivo Excel com os nomes e e-mails
email_file = st.file_uploader(
    "Escolha o arquivo Excel com os e-mails", type="xlsx")
uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

# Função para extrair o nome do titular


def extrair_nome_titular(texto):
    match = re.search(r'NOME:\s+([A-Z\s]+)', texto)
    if match:
        nome = match.group(1).strip()
        nome = nome.replace("\n", " ").strip()
        nome = re.sub(r"[^\w\s]", "", nome)  # remove caracteres especiais
        # Remove "CPF" e qualquer coisa depois
        nome = re.sub(r"\s*CPF.*", "", nome)
        return nome
    return "cliente_desconhecido"

# Função para separar o PDF por cliente com base no plano escolhido


def separar_por_cliente(pdf_path, plano):
    doc = fitz.open(pdf_path)
    cliente_docs = []
    nome_cliente_atual = None
    paginas_atual = []
    arquivos_gerados = []

    for i, pagina in enumerate(doc):
        texto = pagina.get_text()

        if plano == "Uniodonto" and "CLIENTE DO PLANO UNIMASTER-UNI" in texto:
            if paginas_atual:
                caminho = salvar_pdf(doc, paginas_atual, nome_cliente_atual)
                arquivos_gerados.append(caminho)
                paginas_atual = []

            nome_cliente_atual = extrair_nome_titular(texto)

        elif plano == "Unimed" and "Prezado(a) Cliente" in texto:
            if paginas_atual:
                caminho = salvar_pdf(doc, paginas_atual, nome_cliente_atual)
                arquivos_gerados.append(caminho)
                paginas_atual = []

            nome_cliente_atual = extrair_nome_titular(texto)

        if nome_cliente_atual:
            paginas_atual.append(i)

    if paginas_atual:
        caminho = salvar_pdf(doc, paginas_atual, nome_cliente_atual)
        arquivos_gerados.append(caminho)

    doc.close()
    return arquivos_gerados

# Função para salvar o PDF gerado


def salvar_pdf(doc_original, lista_paginas, nome_arquivo_base):
    novo_doc = fitz.open()
    for num in lista_paginas:
        novo_doc.insert_pdf(doc_original, from_page=num, to_page=num)

    pasta_destino = "arquivos_clientes"
    os.makedirs(pasta_destino, exist_ok=True)

    nome_arquivo = os.path.join(pasta_destino, f"{nome_arquivo_base}.pdf")
    novo_doc.save(nome_arquivo)
    novo_doc.close()
    return nome_arquivo

# Função para enviar o e-mail


def enviar_email(destinatario, nome_cliente, arquivo_pdf):
    sender_email = os.getenv("EMAIL")
    sender_password = os.getenv("EMAIL_PASSWORD")

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = destinatario
    msg['Subject'] = f"ADUFC - UNIMED Fortaleza e UNIODONTO Fortaleza (Demonstrativo para PROGEP e IR)"

    body = f"Prezado(a) Professor(a),\n\nSeguem em anexo os demonstrativos de suas despesas com plano de saúde e plano odontológico deste ano, esses documentos deverão ser encaminhados à Pró-Reitoria de Gestão de Pessoas (PROGEP) e à Receita Federal.\n\n\nEm caso de dúvidas, a ADUFC recomenda que os/as docentes entrem em contato com a unidade de gestão de pessoas de sua universidade para obter mais informações quanto à forma de entrega da documentação e dos documentos aceitos para comprovação dos gastos.\n\nAtenciosamente,\nSetor de Atendimento ao Docente"
    msg.attach(MIMEText(body, 'plain'))

    with open(arquivo_pdf, "rb") as file:
        part = MIMEApplication(file.read(), Name=os.path.basename(arquivo_pdf))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo_pdf)}"'
        msg.attach(part)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, destinatario, msg.as_string())
        print(f"E-mail enviado para {destinatario}!")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")

# Função para criar um arquivo .zip com todos os PDFs gerados


def criar_zip(arquivos):
    zip_nome = "arquivos_clientes.zip"
    with zipfile.ZipFile(zip_nome, 'w') as zipf:
        for arquivo in arquivos:
            zipf.write(arquivo, os.path.basename(arquivo))
    return zip_nome


# Verificar se os arquivos foram carregados antes de prosseguir
if email_file and uploaded_file:
    df_emails = pd.read_excel(email_file)

    with open("temp_input.pdf", "wb") as f:
        f.write(uploaded_file.read())

    with st.spinner("📂 Processando o arquivo..."):
        arquivos = separar_por_cliente("temp_input.pdf", plano_selecionado)

    st.success(f"✅ {len(arquivos)} arquivos gerados!")

    # Adicionar um botão para mostrar ou esconder os nomes e botões de download dos PDFs
    with st.expander("Clique para ver os arquivos gerados 🔍"):
        for arquivo in arquivos:
            nome_base = os.path.basename(arquivo)
            st.write(f"- {nome_base}")
            st.download_button(
                label=f"Baixar {nome_base}",
                data=open(arquivo, "rb").read(),
                file_name=nome_base,
                mime="application/pdf",
                key=f"download_{nome_base}_{uuid.uuid4()}"
            )

    # Adicionar um botão para baixar todos os arquivos em um zip
    zip_arquivo = criar_zip(arquivos)
    with open(zip_arquivo, "rb") as f:
        st.download_button(
            label="Baixar todos os arquivos 💾",
            data=f,
            file_name=zip_arquivo,
            mime="application/zip"
        )

    # Adicionar um botão para enviar os e-mails
    if st.button("Enviar E-mails ✉️"):
        for arquivo in arquivos:
            nome_arquivo_cliente = os.path.basename(
                arquivo).replace(".pdf", "")
            # Ajustado para 'Docente'
            cliente_info = df_emails[df_emails['Docente']
                                     == nome_arquivo_cliente]

            if not cliente_info.empty:
                email_cliente = cliente_info.iloc[0]['Email']
                enviar_email(email_cliente, nome_arquivo_cliente, arquivo)
                st.write(
                    f"E-mail enviado para {nome_arquivo_cliente} ({email_cliente})")

        st.success("Todos os e-mails foram enviados!")

else:
    st.error("Por favor, faça o upload de ambos os arquivos: PDF e Excel!")
