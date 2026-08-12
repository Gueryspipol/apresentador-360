import os
import shutil
import tempfile
import unicodedata

from openpyxl import load_workbook


TEMPLATE_EXCEL = (
    "Excel para alimentar o powerpoint - Preenchido.xlsx"
)


def encontrar_operadoras(arquivo_excel):
    wb = load_workbook(
        arquivo_excel,
        data_only=True
    )

    try:
        operadoras = set()

        operadoras_conhecidas = [
            "Amil Selecionada",
            "Amil Metal",
            "Bradesco",
            "SulAmérica",
            "Omint",
            "Hapvida",
            "Seguros Unimed",
        ]

        for nome_aba in [
            "Matriz FX Et. C COPART",
            "Matriz FX Et. S COPART",
        ]:
            if nome_aba not in wb.sheetnames:
                continue

            ws = wb[nome_aba]

            for row in ws.iter_rows():
                for cell in row:
                    valor = str(
                        cell.value or ""
                    ).strip()

                    if not valor:
                        continue

                    for operadora in operadoras_conhecidas:
                        if (
                            valor.casefold()
                            == operadora.casefold()
                        ):
                            operadoras.add(
                                operadora
                            )

        ordem_preferencial = [
            "Amil Selecionada",
            "Amil Metal",
            "Bradesco",
            "Hapvida",
            "Omint",
            "SulAmérica",
            "Seguros Unimed",
        ]

        return [
            operadora
            for operadora in ordem_preferencial
            if operadora in operadoras
        ]

    finally:
        wb.close()


# =====================================================
# ETAPA 7A
# Elegibilidade + Faixa Etária
# =====================================================


def normalizar_texto(valor):
    if valor is None:
        return ""

    texto = str(valor).strip().upper()

    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(
            caractere
        )
    )

    return " ".join(
        texto.split()
    )


def localizar_cabecalhos_base(ws):
    nomes_procurados = {
        "NOME",
        "ELEGIBILIDADE",
        "SEXO",
        "FAIXA ETARIA",
        "PLANO",
        "PLANO2",
    }

    melhor_linha = None
    melhor_quantidade = 0
    melhor_cabecalhos = {}

    limite_linhas = min(
        ws.max_row,
        50
    )

    for linha in range(
        1,
        limite_linhas + 1
    ):
        cabecalhos_linha = {}
        quantidade = 0

        for coluna in range(
            1,
            ws.max_column + 1
        ):
            valor = normalizar_texto(
                ws.cell(
                    linha,
                    coluna
                ).value
            )

            if not valor:
                continue

            cabecalhos_linha.setdefault(
                valor,
                []
            ).append(
                coluna
            )

            if valor in nomes_procurados:
                quantidade += 1

        if quantidade > melhor_quantidade:
            melhor_quantidade = quantidade
            melhor_linha = linha
            melhor_cabecalhos = (
                cabecalhos_linha
            )

    if melhor_linha is None:
        raise ValueError(
            "Não foi possível localizar os "
            "cabeçalhos da aba Base."
        )

    if melhor_quantidade < 3:
        raise ValueError(
            "A linha de cabeçalhos da aba Base "
            "não foi identificada com segurança."
        )

    return (
        melhor_linha,
        melhor_cabecalhos
    )


def primeira_coluna_base(
    cabecalhos,
    nomes
):
    for nome in nomes:
        nome_normalizado = normalizar_texto(
            nome
        )

        colunas = cabecalhos.get(
            nome_normalizado,
            []
        )

        if colunas:
            return colunas[0]

    return None


def ultima_coluna_base(
    cabecalhos,
    nomes
):
    for nome in nomes:
        nome_normalizado = normalizar_texto(
            nome
        )

        colunas = cabecalhos.get(
            nome_normalizado,
            []
        )

        if colunas:
            return colunas[-1]

    return None


def normalizar_faixa_etaria(valor):
    texto = normalizar_texto(
        valor
    )

    texto = (
        texto
        .replace("ANOS", "")
        .replace("ANO", "")
        .strip()
    )

    texto_compacto = (
        texto
        .replace(" ", "")
        .replace("–", "-")
        .replace("—", "-")
    )

    mapa_faixas = {
        "00-18": "00 - 18",
        "0-18": "00 - 18",
        "ATE18": "00 - 18",
        "19-23": "19 - 23",
        "24-28": "24 - 28",
        "29-33": "29 - 33",
        "34-38": "34 - 38",
        "39-43": "39 - 43",
        "44-48": "44 - 48",
        "49-53": "49 - 53",
        "54-58": "54 - 58",
        "59+": "59+",
        "59OU+": "59+",
        "59OUMAIS": "59+",
        "ACIMADE59": "59+",
    }

    return mapa_faixas.get(
        texto_compacto,
        ""
    )


def ler_dados_demograficos(
    caminho_matriz
):
    wb_matriz = load_workbook(
        caminho_matriz,
        data_only=True
    )

    try:
        if "Base" not in wb_matriz.sheetnames:
            raise ValueError(
                "A aba Base não foi encontrada "
                "na matriz."
            )

        ws_base = wb_matriz["Base"]

        if ws_base.sheet_state != "visible":
            raise ValueError(
                "A aba Base está oculta. "
                "O projeto considera somente "
                "abas visíveis."
            )

        (
            linha_cabecalho,
            cabecalhos
        ) = localizar_cabecalhos_base(
            ws_base
        )

        coluna_nome = primeira_coluna_base(
            cabecalhos,
            ["Nome"]
        )

        coluna_elegibilidade = (
            primeira_coluna_base(
                cabecalhos,
                ["Elegibilidade"]
            )
        )

        coluna_sexo = primeira_coluna_base(
            cabecalhos,
            ["Sexo"]
        )

        coluna_faixa = primeira_coluna_base(
            cabecalhos,
            [
                "Faixa etária",
                "Faixa etaria",
            ]
        )

        coluna_plano = primeira_coluna_base(
            cabecalhos,
            [
                "Plano²",
                "Plano2",
            ]
        )

        if coluna_plano is None:
            coluna_plano = ultima_coluna_base(
                cabecalhos,
                ["Plano"]
            )

        colunas_essenciais = {
            "Nome": coluna_nome,
            "Elegibilidade": (
                coluna_elegibilidade
            ),
            "Faixa etária": coluna_faixa,
            "Plano": coluna_plano,
        }

        faltantes = [
            nome
            for nome, coluna
            in colunas_essenciais.items()
            if coluna is None
        ]

        if faltantes:
            raise ValueError(
                "As seguintes colunas não foram "
                "encontradas na aba Base: "
                f"{faltantes}"
            )

        registros_base = []
        linhas_vazias_seguidas = 0

        for linha in range(
            linha_cabecalho + 1,
            ws_base.max_row + 1
        ):
            nome = ws_base.cell(
                linha,
                coluna_nome
            ).value

            elegibilidade = ws_base.cell(
                linha,
                coluna_elegibilidade
            ).value

            plano = ws_base.cell(
                linha,
                coluna_plano
            ).value

            faixa = ws_base.cell(
                linha,
                coluna_faixa
            ).value

            sexo = (
                ws_base.cell(
                    linha,
                    coluna_sexo
                ).value
                if coluna_sexo
                else None
            )

            linha_vazia = (
                not str(nome or "").strip()
                and
                not str(
                    elegibilidade or ""
                ).strip()
                and
                not str(plano or "").strip()
            )

            if linha_vazia:
                linhas_vazias_seguidas += 1

                if (
                    linhas_vazias_seguidas
                    >= 20
                ):
                    break

                continue

            linhas_vazias_seguidas = 0

            elegibilidade_normalizada = (
                normalizar_texto(
                    elegibilidade
                )
            )

            if elegibilidade_normalizada not in {
                "T",
                "D",
                "TITULAR",
                "DEPENDENTE",
                "FUNCIONARIO",
                "FUNCIONARIO(A)",
            }:
                continue

            if not str(nome or "").strip():
                continue

            if not str(plano or "").strip():
                continue

            registros_base.append({
                "nome": str(nome).strip(),
                "elegibilidade": (
                    elegibilidade_normalizada
                ),
                "sexo": normalizar_texto(
                    sexo
                ),
                "faixa": normalizar_faixa_etaria(
                    faixa
                ),
                "plano": str(plano).strip(),
            })

        if not registros_base:
            raise ValueError(
                "Nenhuma vida válida foi encontrada "
                "na aba Base."
            )

        faixas = [
            "00 - 18",
            "19 - 23",
            "24 - 28",
            "29 - 33",
            "34 - 38",
            "39 - 43",
            "44 - 48",
            "49 - 53",
            "54 - 58",
            "59+",
        ]

        dados_faixa = {
            faixa: {
                "funcionarios": 0,
                "dependentes": 0,
            }
            for faixa in faixas
        }

        feminino = 0
        masculino = 0
        funcionarios = 0
        dependentes = 0

        for registro in registros_base:
            elegibilidade = registro[
                "elegibilidade"
            ]

            sexo = registro["sexo"]
            faixa = registro["faixa"]

            if elegibilidade in {
                "T",
                "TITULAR",
                "FUNCIONARIO",
      
