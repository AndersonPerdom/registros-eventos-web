import os
from io import BytesIO
from flask import Flask, render_template, send_file, request
from dotenv import load_dotenv
import psycopg2
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

load_dotenv()

app = Flask(__name__)


def get_connection():
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT", "5432"),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD")
    )


@app.route("/")
def index():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT nombre 
        FROM eventos_permitidos
        WHERE activo = TRUE
        ORDER BY nombre;
    """)
    eventos = [row[0] for row in cur.fetchall()]
    cur.close()
    conn.close()

    return render_template("index.html", eventos=eventos)


@app.route("/descargar")
def descargar_excel():
    evento = request.args.get("evento")

    if not evento:
        return "Debe seleccionar un evento", 400

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT nombre, identidad
        FROM registros_eventos
        WHERE LOWER(evento) = LOWER(%s)
        ORDER BY nombre;
    """, (evento,))

    registros = cur.fetchall()
    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = evento[:31]

    headers = ["Número", "Nombre", "Identidad", "Firma"]
    ws.append(headers)

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for cell in ws[1]:
        cell.font = bold
        cell.alignment = center
        cell.border = border

    for i, registro in enumerate(registros, start=1):
        nombre, identidad = registro
        ws.append([i, nombre, identidad, ""])

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 30

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"{evento}_registros.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
