import streamlit as st
import tempfile

from motor import encontrar_operadoras


st.set_page_config(
    page_title="Apresentador 360",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Apresentador 360")
st.markdown(
    "Gerador automático de Excel e PowerPoint."
)

arquivo_matriz = st.file_uploader(
    "Selecione a Matriz preenchida",
    type=["xlsx"]
)

if arquivo_matriz is not None:

    st.success("Matriz carregada com sucesso!")

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    ) as temp_file:

        temp_file.write(
            arquivo_matriz.read()
        )

        caminho_matriz = temp_file.name

    operadoras = encontrar_operadoras(
        caminho_matriz
    )

    st.subheader("Operadoras encontradas")

    selecionadas = st.multiselect(
        "Selecione até 4 operadoras",
        operadoras,
        max_selections=4
    )

    if selecionadas:

        st.subheader("Definir ordem")

        ordem = []

        for i in range(len(selecionadas)):

            escolha = st.selectbox(
                f"{i+1}ª Operadora",
                selecionadas,
                key=f"ordem_{i}"
            )

            ordem.append(escolha)

        st.write("Ordem escolhida:")
        st.write(ordem)

        if st.button(
            "🚀 Gerar Apresentação"
        ):
            st.success(
                "Motor conectado com sucesso."
            )
