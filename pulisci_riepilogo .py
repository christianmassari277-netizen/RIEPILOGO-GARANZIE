import pandas as pd
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import sys
import os


def processa_file(input_path, output_path):
    records = []
    pattern = re.compile(r"^\s*(\d{7})\s+(\d+)\s+(\d+)\s+(\d+)")

    # Legge il file di testo
    with open(input_path, "r", encoding="latin-1") as f:
        for line in f:
            match = pattern.match(line)
            if match:
                numero_garanzia, suffisso, job, totale_job = match.groups()
                records.append([numero_garanzia, suffisso, int(job), int(totale_job)])

    # Crea DataFrame
    df = pd.DataFrame(records, columns=["NUMERO GARANZIA", "SUFFISSO", "JOB", "TOTALE JOB"])

    # Elimina duplicati
    df = df.drop_duplicates(subset=["NUMERO GARANZIA", "SUFFISSO", "JOB"], keep="first")

    # Ordina i dati
    df = df.sort_values(by=["NUMERO GARANZIA", "SUFFISSO", "JOB"])

    # Calcola il totale
    totale_generale = df["TOTALE JOB"].sum()

    # Crea PDF
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph("Riepilogo Garanzie", styles["Title"]))
    elements.append(Spacer(1, 12))

    # Tabella
    data = [list(df.columns)] + df.astype(str).values.tolist()
    data.append(["", "", "TOTALE", str(totale_generale)])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.lightgrey),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
    ]))

    elements.append(table)

    doc.build(elements)
    print(f"PDF creato con successo: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        input_path = sys.argv[1]
        base, _ = os.path.splitext(input_path)
        output_path = base + "_PULITO.pdf"
        processa_file(input_path, output_path)
    else:
        print("Trascina un file .txt sopra l'eseguibile per elaborarlo.")