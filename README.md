# Convertitore CSV a XLSX

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Release](https://img.shields.io/badge/release-1.0.1-orange)](https://github.com/tuo-username/convertitore-csv-xlsx/releases)
[![YouTube Tutorial](https://img.shields.io/badge/YouTube-Tutorial-red?logo=youtube)](https://www.youtube.com/@AngoloInformatico)

---

## 🔹 Descrizione

Un'applicazione GUI moderna per convertire file **CSV** in formato **Excel (XLSX)** con pochi clic.
Ideale per chi vuole gestire dati in CSV e trasformarli rapidamente in Excel senza sforzo.

---

## 🎬 Anteprima 

<img width="400" height="298" alt="img" src="https://github.com/user-attachments/assets/bd418452-7277-4dbe-a0b9-adea8eaaac41" />

---

## ⚙️ Requisiti

* **Python 3.8+**
* Librerie Python:

  * `customtkinter`
  * `pandas`
  * `tkinter` (incluso in Python standard)

Installazione rapida:

```bash
pip install customtkinter pandas
```

---

## 🚀 Installazione e Avvio

### Clonare il repository

```bash
git clone https://github.com/tuo-username/convertitore-csv-xlsx.git
cd convertitore-csv-xlsx
```

### Avvio dell'applicazione

```bash
python main.py
```

---

## 💻 Uso

1. Clicca su **Seleziona File CSV** per scegliere il file CSV dal tuo computer.
2. Il percorso del file apparirà nella label sottostante.
3. Clicca su **Converti in XLSX** per salvare il file in formato Excel.
4. Se vuoi, apri subito il file XLSX convertito cliccando **Sì** nel popup.

---

## 🖥️ Compatibilità OS

* **Windows**: apertura file e link con `os.startfile()`.
* **macOS**: apertura file e link con `open`.
* **Linux**: modifica le funzioni `_open_link_1`, `_open_link_2` e `open_file` sostituendo `open` con `xdg-open`:

```python
subprocess.run(['xdg-open', filepath], check=True)
```

---

## ✨ Personalizzazione

* Cambia i link cliccabili modificando `self.url_1` e `self.url_2` in `main.py`.
* Modifica il font o il tema scuro/chiaro in `__init__` secondo le tue preferenze.

---

## 📝 Crediti

* Creato da **Alex Lignola** - Release 1.0.1
* © 2025 Convertitore CSV a XLSX - Tutti i diritti riservati
* [Tutorial video su YouTube](https://www.youtube.com/@AngoloInformatico)

---

## 📄 Licenza

Rilasciato sotto licenza **MIT**. Vedi il file [LICENSE](LICENSE) per i dettagli.

---

## 🔗 Riferimenti

* [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) – GUI moderna per Tkinter
* [Pandas](https://pandas.pydata.org/) – Libreria per gestione dati e conversione CSV/XLSX
* [Tkinter](https://docs.python.org/3/library/tkinter.html) – Libreria GUI standard di Python
