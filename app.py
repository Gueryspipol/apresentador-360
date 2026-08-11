import os
import tempfile

import streamlit as st

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

    # --------------------------------------------------------
    # SALVAR MATRIZ TEMPORARIAMENTE
    # --------------------------------------------------------

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
            "Não foi possível ler as operadoras da matriz."
        )

        st.exception(erro)

        operadoras = []

    finally:

        if os.path.exists(caminho_matriz):
            os.remove(caminho_matriz)


    # --------------------------------------------------------
    # OPERADORAS ENCONTRADAS
    # --------------------------------------------------------

    if not operadoras:

        st.warning(
            "Nenhuma operadora foi encontrada na matriz."
        )

        st.stop()


    st.subheader("Operadoras encontradas")

    selecionadas = st.multiselect(
        "Selecione de 1 até 4 operadoras",
        options=operadoras,
        max_selections=4
    )


    # --------------------------------------------------------
    # DEFINIR ORDEM
    # --------------------------------------------------------

    if selecionadas:

        st.subheader("Definir ordem")

        ordem = []

        for indice in range(
            len(selecionadas)
        ):

            # Remove as operadoras já escolhidas
            # nas posições anteriores.
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


        st.markdown("### Ordem escolhida")

        for posicao, operadora in enumerate(
            ordem,
            start=1
        ):

            st.write(
                f"{posicao}º - {operadora}"
            )


        # ----------------------------------------------------
        # VALIDAÇÕES
        # ----------------------------------------------------

        ordem_valida = True

        if len(ordem) != len(set(ordem)):

            ordem_valida = False

            st.error(
                "Não é permitido repetir operadoras."
            )


        if len(ordem) < 1:

            ordem_valida = False

            st.warning(
                "Selecione pelo menos uma operadora."
            )


        if len(ordem) > 4:

            ordem_valida = False

            st.error(
                "Selecione no máximo quatro operadoras."
            )


        # ----------------------------------------------------
        # BOTÃO GERAR
        # ----------------------------------------------------

        if st.button(
            "🚀 Gerar Apresentação",
            type="primary",
            disabled=not ordem_valida
        ):

            st.success(
                "Seleção validada com sucesso."
            )

            st.write(
                "Operadoras enviadas ao motor:"
            )

            st.json(ordem)
