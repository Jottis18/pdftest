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
    linhas = texto.splitlines()

    if plano == "Unimed":
        for i, linha in enumerate(linhas):
            if "Carteira:" in linha:
                if i > 0:
                    nome = linhas[i - 1].strip()
                    nome = re.sub(r"[^\w\sÀ-ÿ]", "", nome)
                    return nome
                break

    elif plano == "Uniodonto":
        # Permite letras (inclusive acentuadas), espaços e hífen antes do CPF
        match = re.search(
            r'^([A-Za-zÀ-ÿ\s-]+)\s+-\s+\d{3}\.\d{3}\.\d{3}-\d{2}',
            texto,
            re.MULTILINE
        )
        if match:
            nome = match.group(1).strip()
            # Mantém letras (incluindo acentos), espaços e hífen
            return re.sub(r"[^A-Za-zÀ-ÿ\s-]", "", nome)

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
    df_emails = pd.read_excel(email_file)
    with open("temp_input.pdf", "wb") as f:
        f.write(uploaded_file.read())

    with st.spinner("Processando..."):
        arquivos = separar_por_cliente("temp_input.pdf", plano_selecionado)
    st.success(f"{len(arquivos)} PDFs gerados!")

    # Exibe e permite download dos PDFs individualmente
    with st.expander("🔍 Arquivos gerados"):
        for a in arquivos:
            nome = os.path.basename(a)
            st.write(f"- {nome}")
            st.download_button(f"Baixar {nome}", data=open(a, "rb").read(),
                               file_name=nome, mime="application/pdf",
                               key=str(uuid.uuid4()))

    # Botão para baixar todos os PDFs em ZIP
    zip_arquivo = criar_zip(arquivos)
    with open(zip_arquivo, "rb") as fzip:
        st.download_button("📥 Baixar todos os PDFs (ZIP)", fzip, zip_arquivo, "application/zip")

    # Envio de e-mails
    if st.button("Enviar E-mails ✉️"):
        erros_envio = []
        sem_corresp = []
        sucessos = []
        cont = 0

        for pdf in arquivos:
            nome_cliente = os.path.basename(pdf).replace(".pdf", "")
            info = df_emails[df_emails['Docente'] == nome_cliente]
            if info.empty:
                sem_corresp.append({'Docente': nome_cliente})
                st.warning(f"⚠️ Sem correspondência: {nome_cliente}")
                continue

            email = info.iloc[0]['Email']
            ok, err = enviar_email(email, nome_cliente, pdf)
            if ok:
                sucessos.append({'Docente': nome_cliente, 'Email': email})
                st.write(f"✅ {nome_cliente} ({email})")
            else:
                erros_envio.append({'Docente': nome_cliente, 'Email': email, 'Erro': err})
                st.error(f"❌ {nome_cliente} ({email}): {err}")

            time.sleep(0.8)
            cont += 1
            if cont >= 50:
                st.warning("⏳ Aguardando 60s...")
                time.sleep(60)
                cont = 0

        # Relatórios individuais
        if sucessos:
            df_suc = pd.DataFrame(sucessos)
            st.success(f"{len(sucessos)} e-mails enviados com sucesso:")
            st.dataframe(df_suc)
            arquivo_suc = "sucessos_envio.xlsx"
            df_suc.to_excel(arquivo_suc, index=False)
            with open(arquivo_suc, "rb") as fs:
                st.download_button("📄 Baixar relatório de sucessos", fs,
                                   arquivo_suc,
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if erros_envio:
            df_err = pd.DataFrame(erros_envio)
            st.error(f"{len(erros_envio)} falhas de envio:")
            st.dataframe(df_err)
            arquivo_err = "erros_envio.xlsx"
            df_err.to_excel(arquivo_err, index=False)
            with open(arquivo_err, "rb") as fe:
                st.download_button("📄 Baixar relatório de erros", fe,
                                   arquivo_err,
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if sem_corresp:
            df_sem = pd.DataFrame(sem_corresp)
            st.warning(f"{len(sem_corresp)} sem correspondência no Excel:")
            st.dataframe(df_sem)
            arquivo_sem = "sem_correspondencia.xlsx"
            df_sem.to_excel(arquivo_sem, index=False)
            with open(arquivo_sem, "rb") as fsc:
                st.download_button("📄 Baixar log de sem correspondência", fsc,
                                   arquivo_sem,
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ZIP com todos os relatórios Excel
        excel_files = []
        if 'arquivo_suc' in locals(): excel_files.append(arquivo_suc)
        if 'arquivo_err' in locals(): excel_files.append(arquivo_err)
        if 'arquivo_sem' in locals(): excel_files.append(arquivo_sem)
        if excel_files:
            zip_excels = "relatorios_excel.zip"
            with zipfile.ZipFile(zip_excels, "w") as zf:
                for ef in excel_files:
                    zf.write(ef, os.path.basename(ef))
            with open(zip_excels, "rb") as fzip_exc:
                st.download_button(
                    "📥 Baixar todos os relatórios Excel",
                    data=fzip_exc,
                    file_name=zip_excels,
                    mime="application/zip"
                )

        # Finalização
        if not erros_envio and not sem_corresp:
            st.balloons()
            st.success("Tudo processado com sucesso! 🎉")
            st.info("⏳ Esperando 10 minutos para manter o app ativo...")
            time.sleep(600)

else:
    st.error("Faça upload do PDF e do Excel para prosseguir.")
