import streamlit as st
import pandas as pd

st.set_page_config(page_title="Preencher CST/CFOP por NCM", layout="centered")

st.title("🔄 Preencher CST/CFOP com base no NCM")
st.markdown("Envie duas planilhas: **ListaProdutos** e **Base**. O app preencherá automaticamente os campos CST Saída e CFOP Venda com base no NCM.")

lista_file = st.file_uploader("📄 Enviar planilha ListaProdutos", type=["xlsx"])
base_file = st.file_uploader("📄 Enviar planilha Base", type=["xlsx"])

if lista_file and base_file:
    df_lista = pd.read_excel(lista_file)
    df_base = pd.read_excel(base_file)

    if 'NCM' not in df_lista.columns or 'CST Saída' not in df_lista.columns or 'CFOP Venda' not in df_lista.columns:
        st.error("A planilha ListaProdutos deve conter as colunas: NCM, CST Saída e CFOP Venda.")
    elif 'NCM' not in df_base.columns or 'CST SAIDA' not in df_base.columns or 'CFOP SAIDA' not in df_base.columns:
        st.error("A planilha Base deve conter as colunas: NCM, CST SAIDA e CFOP SAIDA.")
    else:
        df_merged = df_lista.merge(df_base, how='left', on='NCM', suffixes=('', '_base'))
        df_lista['CST Saída'] = df_lista['CST Saída'].fillna(df_merged['CST SAIDA'])
        df_lista['CFOP Venda'] = df_lista['CFOP Venda'].fillna(df_merged['CFOP SAIDA'])

        st.success("✅ Preenchimento concluído!")
        
        # Create Excel file for download with proper method
        import io
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_lista.to_excel(writer, index=False, sheet_name='ListaProdutos')
        excel_data = output.getvalue()
        
        st.download_button(
            label="⬇️ Baixar planilha preenchida",
            data=excel_data,
            file_name="ListaProdutos_Preenchido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_lista.head(10))