from openpyxl import load_workbook


def encontrar_operadoras(arquivo_excel):

    wb = load_workbook(
        arquivo_excel,
        data_only=True
    )

    operadoras = set()

    operadoras_conhecidas = [
        "Amil",
        "Bradesco",
        "SulAmérica",
        "Sulamerica",
        "Omint",
        "Hapvida",
        "Seguros Unimed",
        "Porto Seguro"
    ]

    for aba in wb.sheetnames:

        if "Matriz FX" not in aba:
            continue

        ws = wb[aba]

        for row in ws.iter_rows():

            for cell in row:

                valor = str(
                    cell.value or ""
                ).strip()

                if not valor:
                    continue

                for operadora in operadoras_conhecidas:

                    if operadora.lower() in valor.lower():

                        operadoras.add(
                            operadora
                        )

    wb.close()

    return sorted(
        list(operadoras)
    )
``
