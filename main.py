import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Preencher CST/CFOP por NCM",
                   layout="centered",
                   initial_sidebar_state="collapsed")

st.title("📊 Preencher CST/CFOP com base no NCM")
st.caption(
    "Envie as planilhas e o sistema preencherá automaticamente os campos com base no NCM."
)
st.divider()

st.subheader("📂 Upload das Planilhas")
col1, col2 = st.columns(2)
with col1:
    lista_file = st.file_uploader("🔹 ListaProdutos", type=["xlsx"])
with col2:
    base_file = st.file_uploader("🔸 Base de Referência", type=["xlsx"])

st.divider()

if lista_file and base_file:
    try:
        df_lista = pd.read_excel(lista_file)
        df_base = pd.read_excel(base_file)

        df_lista.columns = df_lista.columns.str.strip().str.upper()
        df_base.columns = df_base.columns.str.strip().str.upper()

        if not {"NCM", "CST SAÍDA", "CFOP VENDA"}.issubset(df_lista.columns):
            st.error(
                "❌ A planilha *ListaProdutos* precisa conter: NCM, CST SAÍDA, CFOP VENDA."
            )
        elif not {"NCM", "CST SAIDA", "CFOP SAIDA"}.issubset(df_base.columns):
            st.error(
                "❌ A planilha *Base* precisa conter: NCM, CST SAIDA, CFOP SAIDA."
            )
        else:
            with st.spinner("🔄 Preenchendo com base no NCM..."):
                cst_dict = df_base.set_index("NCM")["CST SAIDA"].to_dict()
                cfop_dict = df_base.set_index("NCM")["CFOP SAIDA"].to_dict()

                df_lista["CST SAÍDA"] = df_lista["CST SAÍDA"].fillna(
                    df_lista["NCM"].map(cst_dict))
                df_lista["CFOP VENDA"] = df_lista["CFOP VENDA"].fillna(
                    df_lista["NCM"].map(cfop_dict))

            st.success("✅ Dados preenchidos com sucesso!")

            col3, col4 = st.columns(2)
            col3.metric("CST Saída Preenchidos",
                        df_lista["CST SAÍDA"].notna().sum())
            col4.metric("CFOP Venda Preenchidos",
                        df_lista["CFOP VENDA"].notna().sum())

            with st.expander("👁️ Visualizar os primeiros dados"):
                st.dataframe(df_lista.head(10), use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_lista.to_excel(writer,
                                  index=False,
                                  sheet_name="ListaProdutos")

            st.download_button(
                label="⬇️ Baixar planilha preenchida",
                data=output.getvalue(),
                file_name="ListaProdutos_Preenchido.xlsx",
                mime=
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Erro ao processar os arquivos: {e}")
else:
    st.info("👆 Envie os dois arquivos para iniciar o preenchimento.")
