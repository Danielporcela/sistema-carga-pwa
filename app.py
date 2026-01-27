from flask import Flask, render_template, request
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import io

app = Flask(__name__)

# ================== REGRAS FIXAS ==================
COL_INI = "B"
COL_FIM = "T"
LINHA_INICIAL = 3
INTERVALO = 28
OFFSET_RESULTADO = 21
TOTAL_SETORES = 54
# ==================================================

def normalizar(valor):
    if valor is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(valor).upper())

def carregar_planilha(arquivo_bytes):
    wb = load_workbook(io.BytesIO(arquivo_bytes), data_only=True)
    return wb.active

def buscar_setor(planilha, codigo):
    col_ini = column_index_from_string(COL_INI)
    col_fim = column_index_from_string(COL_FIM)
    codigo = normalizar(codigo)

    for i in range(TOTAL_SETORES):
        linha = LINHA_INICIAL + i * INTERVALO
        for col in range(col_ini, col_fim + 1):
            valor_setor = planilha.cell(linha, col).value
            if normalizar(valor_setor) == codigo:
                carga = planilha.cell(linha + OFFSET_RESULTADO, col).value
                return valor_setor, carga
    return None

def gerar_ranking(planilha, crescente=True):
    ranking = []
    col_ini = column_index_from_string(COL_INI)
    col_fim = column_index_from_string(COL_FIM)

    for i in range(TOTAL_SETORES):
        linha = LINHA_INICIAL + i * INTERVALO
        for col in range(col_ini, col_fim + 1):
            setor = planilha.cell(linha, col).value
            carga = planilha.cell(linha + OFFSET_RESULTADO, col).value

            if setor and isinstance(carga, (int, float)):
                ranking.append({
                    "setor": setor,
                    "carga": carga
                })

    ranking.sort(key=lambda x: x["carga"], reverse=not crescente)
    return ranking

@app.route("/", methods=["GET", "POST"])
def index():
    resultado = None
    ranking = []
    ordem = request.form.get("ordem", "crescente")

    if request.method == "POST":
        arquivo = request.files.get("arquivo")
        codigo_setor = request.form.get("setor", "").strip()

        if arquivo:
            planilha = carregar_planilha(arquivo.read())

            if codigo_setor:
                resultado = buscar_setor(planilha, codigo_setor)

            ranking = gerar_ranking(
                planilha,
                crescente=(ordem == "crescente")
            )

    return render_template(
        "index.html",
        resultado=resultado,
        ranking=ranking,
        ordem=ordem
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

