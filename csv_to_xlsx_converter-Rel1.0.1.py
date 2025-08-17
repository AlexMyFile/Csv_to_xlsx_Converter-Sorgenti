import customtkinter
import pandas as pd
from tkinter import filedialog, messagebox
import os
import subprocess  # Per aprire il file
from tkinter import font  # Per gestire i font

class CSVtoXLSXApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Convertitore CSV a XLSX")
        self.geometry("400x270")  # Aumenta un po' l'altezza
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("dark-blue")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure((0, 1, 2, 3, 4, 5), weight=1)  # Aggiungi righe per le label

        # Font Sogei
        try:
            sogei_font = font.Font(family="Sogei", size=32)
            self.title_font = customtkinter.CTkFont(family=sogei_font['family'], size=sogei_font['size'])
        except:
            # Se il font Sogei non è disponibile, usa un font di sistema
            self.title_font = customtkinter.CTkFont(family="Helvetica", size=32, weight="bold")
            print("Font 'Sogei' non trovato. Utilizzo 'Helvetica' in sostituzione.")

        # Label titolo personalizzata
        self.title_label = customtkinter.CTkLabel(self, text="Convertitore File", font=self.title_font)  # Titolo più breve
        self.title_label.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")

        self.select_csv_button = customtkinter.CTkButton(self, text="Seleziona File CSV", command=self.select_csv)
        self.select_csv_button.grid(row=1, column=0, padx=20, pady=5, sticky="ew")

        self.csv_filepath_label = customtkinter.CTkLabel(self, text="Nessun file CSV selezionato")
        self.csv_filepath_label.grid(row=2, column=0, padx=20, sticky="ew")

        self.convert_button = customtkinter.CTkButton(self, text="Converti in XLSX", command=self.convert_csv_to_xlsx, state="disabled")
        self.convert_button.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        # Label copyright - Parte 1
        self.copyright_font = customtkinter.CTkFont(family="Segoe UI", size=12, underline=True)
        self.copyright_label_1 = customtkinter.CTkLabel(self, text="Created by Alex Lignola - Release 1.0.1",
                                                      font=self.copyright_font, text_color="#1E90FF", cursor="hand2")
        self.copyright_label_1.grid(row=4, column=0, padx=20, pady=(5, 0), sticky="ew")
        self.copyright_label_1.bind("<Button-1>", self._open_link_1)  # Definisci _open_link_1

        # Label copyright - Parte 2
        self.copyright_label_2 = customtkinter.CTkLabel(self, text="© 2025 Convertitore CSV a XLSX - All rights reserved.",
                                                      font=self.copyright_font, text_color="#1E90FF", cursor="hand2")
        self.copyright_label_2.grid(row=5, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.copyright_label_2.bind("<Button-1>", self._open_link_2)  # Definisci _open_link_2

        self.csv_filepath = None
        self.url_1 = "https://www.youtube.com/@AngoloInformatico"  # Sostituisci con il tuo URL
        self.url_2 = "https://www.youtube.com/@AngoloInformatico"  # Sostituisci con il tuo URL

    def select_csv(self):
        self.csv_filepath = filedialog.askopenfilename(
            title="Seleziona file CSV",
            filetypes=[("CSV files", "*.csv")]
        )
        if self.csv_filepath:
            self.csv_filepath_label.configure(text=self.csv_filepath)
            self.convert_button.configure(state="normal")
        else:
            self.csv_filepath_label.configure(text="Nessun file CSV selezionato")
            self.convert_button.configure(state="disabled")

    def open_file(self, filepath):
        """Apre il file utilizzando l'applicazione predefinita del sistema operativo."""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(filepath)
            elif os.name == 'posix':  # macOS e Linux
                subprocess.run(['open', filepath], check=True)
            else:
                messagebox.showerror("Errore", f"Impossibile aprire il file su questo sistema operativo: {os.name}")
        except FileNotFoundError:
            messagebox.showerror("Errore", f"Il file non è stato trovato: {filepath}")
        except Exception as e:
            messagebox.showerror("Errore", f"Si è verificato un errore nell'apertura del file:\n{e}")

    def convert_csv_to_xlsx(self):
        if self.csv_filepath:
            try:
                df = pd.read_csv(self.csv_filepath)
                base_name = os.path.splitext(os.path.basename(self.csv_filepath))[0]
                default_xlsx_filename = f"{base_name}.xlsx"
                initial_dir = os.path.dirname(self.csv_filepath)

                xlsx_filepath = filedialog.asksaveasfilename(
                    title="Salva come file XLSX",
                    defaultextension=".xlsx",
                    initialfile=default_xlsx_filename,
                    initialdir=initial_dir,
                    filetypes=[("XLSX files", "*.xlsx")]
                )
                if xlsx_filepath:
                    df.to_excel(xlsx_filepath, index=False)
                    messagebox.showinfo("Successo", f"File CSV convertito e salvato come:\n{xlsx_filepath}")
                    ask_open = messagebox.askyesno("Apri File", "Vuoi aprire il file XLSX convertito?")
                    if ask_open:
                        self.open_file(xlsx_filepath)
            except FileNotFoundError:
                messagebox.showerror("Errore", "File CSV non trovato.")
            except Exception as e:
                messagebox.showerror("Errore", f"Si è verificato un errore durante la conversione:\n{e}")
        else:
            messagebox.showerror("Errore", "Nessun file CSV selezionato.")

    def _open_link_1(self, event):
        """Apre il primo URL."""
        try:
            if os.name == 'nt':
                os.startfile(self.url_1)
            elif os.name == 'posix':
                subprocess.run(['open', self.url_1], check=True)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il link: {e}")

    def _open_link_2(self, event):
        """Apre il secondo URL."""
        try:
            if os.name == 'nt':
                os.startfile(self.url_2)
            elif os.name == 'posix':
                subprocess.run(['open', self.url_2], check=True)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il link: {e}")

if __name__ == "__main__":
    app = CSVtoXLSXApp()
    app.mainloop()