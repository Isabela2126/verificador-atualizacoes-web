import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import os

try:
    import processador 
except ModuleNotFoundError:
    st.error("Erro: O arquivo 'processador.py' não foi encontrado na mesma pasta do app. Por favor, verifique os nomes e a localização dos arquivos.")
    st.stop()

#Configuração da página
st.set_page_config(layout="wide", page_title="Automação de Emails", page_icon="👩‍💻")
st.set_page_config(
    page_title="Verificador PQ 10",
    layout="wide",
    initial_sidebar_state="collapsed"
)

#Cabeçalho com logo e título
col1, col2 = st.columns([1, 4])

with col1:
    try:
        st.image("logo.png", width=150) 
    except FileNotFoundError:
        st.warning("Arquivo 'logo.png' não encontrado.")

with col2:
    st.title("PQ 10 (atualizações de normas/leis/decretos)")
    st.caption("Ferramenta para verificação automática de datas de atualização")

st.divider()

with st.container(border=True):
    st.subheader("Carregue o Documento")
    st.write("Faça o upload do arquivo **Lista Mestra de Requisitos Legais** (.docx) para iniciar.")
    
    uploaded_file = st.file_uploader(
        "Selecione o arquivo .docx da sua máquina", 
        type=["docx"],
        label_visibility="collapsed"
    )

if uploaded_file is not None:
    with st.container(border=True):
        st.subheader("Inicie a Verificação")
        st.info(f"Arquivo carregado: **{uploaded_file.name}**")
        
        if st.button("🔍 Iniciar Verificação Agora", type="primary", use_container_width=True):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                temp_path = tmp_file.name

            with st.spinner('Aguarde, os sites estão sendo acesados para verificar as datas. Isso pode levar alguns minutos'):
                try:
                    df_resultado, nome_arquivo_excel = processador.executar_verificacao(temp_path)
                    
                    st.success("Verificação concluída com sucesso!")

                    with st.container(border=True):
                        st.subheader("Baixe os Resultados")
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_resultado.to_excel(writer, index=False, sheet_name='Verificacao')
                        excel_bytes = output.getvalue()

                        st.write("Prévia dos resultados:")
                        st.dataframe(df_resultado)
                        st.divider()

                        st.download_button(
                            label="📥 Baixar Planilha de Resultados (.xlsx)",
                            data=excel_bytes,
                            file_name=nome_arquivo_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Ocorreu um erro durante o processamento:")
                    st.exception(e)
                finally:
                    os.remove(temp_path)

st.divider()
st.write("Desenvolvido para otimizar esse processo que é demorado de fazer manualmente")