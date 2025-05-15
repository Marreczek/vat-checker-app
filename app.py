import os
import uuid
import re
from io import BytesIO
from flask import Flask, request, render_template, redirect, url_for, send_file, session
from datetime import datetime
import requests
import openpyxl

app = Flask(__name__)
app.secret_key = "tajny_klucz_do_sesji"  # potrzebny do sesji


# --- Funkcja do sprawdzania NIP w rejestrze VAT ---
def sprawdz_nip_w_vat(nip):
    nip = re.sub(r"\D", "", str(nip))  # usuń wszystko poza cyframi
    if not nip.isdigit() or len(nip) != 10:
        return nip, "Nieprawidłowy NIP", "Błąd"

    base_url = "https://wl-api.mf.gov.pl/api/search/nip/"
    today = datetime.today().strftime('%Y-%m-%d')
    url = f"{base_url}{nip}?date={today}"
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'RequestId': str(uuid.uuid4())
    }

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            result = response.json()
            subject = result.get('result', {}).get('subject')
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


# --- Wczytanie NIP-ów z pliku Excel ---
def wczytaj_nipy_z_excel(plik):
    wb = openpyxl.load_workbook(plik)
    ws = wb.active
    nipy = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        nip = row[0]
        if nip:
            nipy.append(str(nip).strip())
    return nipy


# --- Generowanie pliku Excel z wynikami ---
def generuj_excel(wyniki):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["NIP", "Nazwa podmiotu", "Status VAT"])
    for w in wyniki:
        ws.append(w)
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


# --- Strona główna i obsługa formularza ---
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nip_input = request.form.get("nip", "").strip()
        plik = request.files.get("plik")

        if nip_input:
            wynik = sprawdz_nip_w_vat(nip_input)
            return render_template("wyniki.html", wyniki=[wynik], pojedynczy=True)

        elif plik and plik.filename.endswith(".xlsx"):
            nipy = wczytaj_nipy_z_excel(plik)
            wyniki = [sprawdz_nip_w_vat(nip) for nip in nipy]

            # Zapisz wyniki w sesji
            session['wyniki'] = wyniki

            return render_template("wyniki.html", wyniki=wyniki, pojedynczy=False)

        else:
            return render_template("index.html", blad="Wpisz NIP lub załaduj plik .xlsx.")
    return render_template("index.html")


# --- Endpoint do pobrania wyników z sesji ---
@app.route("/pobierz_wyniki")
def pobierz_wyniki():
    wyniki = session.get('wyniki')
    if not wyniki:
        return redirect(url_for('index'))

    excel = generuj_excel(wyniki)
    return send_file(
        excel,
        as_attachment=True,
        download_name="wyniki_nip.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
