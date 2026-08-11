import os
import tempfile

import streamlit as st

from motor import (
    encontrar_operadoras,
    gerar_excel_final,
)


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

    st.success(
        "Matriz carregada com sucesso!"
    )

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    ) as temp_file:

        temp_file.write(
            arquivo_matriz.getvalue()
        )

        caminho_matriz = temp_file.name

    try:

        operadoras = encontrar_operadoras(
            caminho_matriz
        )

    except Exception as erro:

        st.error(
            "Não foi possível ler a matriz."
        )

        st.exception(erro)

        operadoras = []


    if not operadoras:

        st.warning(
            "Nenhuma operadora foi encontrada."
        )

        if os.path.exists(caminho_matriz):
            os.remove(caminho_matriz)

        st.stop()


    st.subheader(
        "Operadoras encontradas"
    )

    selecionadas = st.multiselect(
        "Selecione de 1 até 4 operadoras",
        options=operadoras,
        max_selections=4
    )


    if selecionadas:

        st.subheader(
            "Definir ordem"
        )

        ordem = []

        for indice in range(
            len(selecionadas)
        ):

            opcoes_disponiveis = [
                operadora
                for operadora in selecionadas
                if operadora not in ordem
            ]

            escolha = st.selectbox(
                f"{indice + 1}ª Operadora",
                options=opcoes_disponiveis,
                key=(
                    f"ordem_"
                    f"{arquivo_matriz.name}_"
                    f"{indice}_"
                    f"{len(selecionadas)}"
                )
            )

            ordem.append(escolha)


        st.markdown(
            "### Ordem escolhida"
        )

        for posicao, operadora in enumerate(
            ordem,
            start=1
        ):
            st.write(
                f"{posicao}º - {operadora}"
            )


        if st.button(
            "🚀 Gerar Excel",
            type="primary"
        ):

            try:

                with st.spinner(
                    "Gerando Excel Alimentador..."
                ):

                    caminho_excel = gerar_excel_final(
                        caminho_matriz,
                        ordem
                    )

                    with open(
                        caminho_excel,
                        "rb"
                    ) as arquivo:

                        excel_bytes = arquivo.read()

                st.success(
                    "Excel gerado com sucesso!"
                )

                st.download_button(
                    label=(
                        "⬇️ Baixar Excel Final"
                    ),
                    data=excel_bytes,
                    file_name=(
                        "Apresentador360_"
                        "Excel_Final.xlsx"
                    ),
                    mime=(
                        "application/vnd."
                        "openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="primary"
                )

            except Exception as erro:

                st.error(
                    "Não foi possível gerar o Excel."
                )

                st.exception(erro)
