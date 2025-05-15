from flask import Flask, request, render_template, send_file, redirect, url_for
import requests
import uuid
from datetime import datetime
import openpyxl
from io import BytesIO
import os

app = Flask(__name__)

def sprawdz_nip_w_vat(nip):
    nip = str(nip).replace('-', '').strip()
    if not nip.isdigit() or len(nip) != 10:
        return nip, "Nieprawidłowy NIP", "Błąd"

    base_url = "https://wl-api.mf.gov.pl/api/search/nip/"
    today = datetime.today().strftime('%Y-%m-%d')
    request_id = str(uuid.uuid4())

    url = f"{base_url}{nip}?date={today}"
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'RequestId': request_id
    }

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            result = response.json()
            subject = result['result']['subject']
            if subject:
                nazwa = subject.get('name', 'Brak danych')
                status_vat = subject.get('statusVat', 'Nieznany')
                return nip, nazwa, status_vat
            else:
                return nip, "Nie znaleziono w rejestrze", "Brak"
        else:
            return nip, "Błąd odpowiedzi", f"Kod {response.status_code}"
    except Exception as e:
        return nip, "Błąd zapytania", str(e)

def wczytaj_nipy_z_excel(file_stream):
    wb = openpyxl.load_workbook(file_stream)
    ws = wb.active
    nipy = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        nip = row[0]
        if nip:
            nipy.append(str(nip).strip())
    return nipy

def generuj_excel(wyniki):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wyniki"
    ws.append(["NIP", "Nazwa podmiotu", "Status VAT"])
    for wiersz in wyniki:
        ws.append(wiersz)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        plik = request.files.get("plik")
        if not plik or not plik.filename.endswith(".xlsx"):
            return render_template("index.html", blad="Wgraj poprawny plik .xlsx.")

        try:
            nipy = wczytaj_nipy_z_excel(plik)
            wyniki = [sprawdz_nip_w_vat(nip) for nip in nipy]

            # Tymczasowe zapisanie do sesji lub pliku — tu zapisujemy do pliku tymczasowego
            excel_data = generuj_excel(wyniki)
            excel_data.save("ostatnie_wyniki.xlsx")

            return render_template("wyniki.html", wyniki=wyniki)

        except Exception as e:
            return render_template("index.html", blad=f"Błąd: {e}")

    return render_template("index.html")


@app.route("/pobierz")
def pobierz():
    if not os.path.exists("ostatnie_wyniki.xlsx"):
        return "Brak pliku do pobrania", 404
    return send_file("ostatnie_wyniki.xlsx", as_attachment=True)
