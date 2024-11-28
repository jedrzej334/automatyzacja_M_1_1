import tkinter as tk
from tkinter import simpledialog
from openpyxl import load_workbook
import clr

clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

# Add a reference to the APx API        
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 8.1\\API\\AudioPrecision.API2.dll")
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 8.1\\API\\AudioPrecision.API.dll")

from AudioPrecision.API import *
from System.Drawing import Point
from System.Windows.Forms import Application, Button, Form, Label
from System.IO import Directory, Path

APx = APx500_Application(APxOperatingMode.BenchMode, True)


class AudioInterfaceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Interfejs Audio")
        self.root.geometry("500x700")

        ##### Zmienne początkowe ######
        self.reference_voltage_94dB = 1.0  # Napięcie referencyjne dla 94 dB SPL (przykładowo 1 V RMS)
        self.value_dB = 94.0  # Początkowy poziom dB
        self.value_voltage = self.reference_voltage_94dB  # Początkowe napięcie odpowiadające 94 dB
        self.freqList = [63, 125, 250, 500, 1000, 2000, 4000, 8000, 16000]  # Częstotliwości
        self.freq_var = tk.IntVar(value=0)  # Domyślna częstotliwość: 63 Hz
        self.values_list = []  # Lista do przechowywania wyników

        # Tworzenie GUI
        self.label = tk.Label(self.root, text="Interfejs Audio", font=("Arial", 16))
        self.label.place(x=150, y=20)

        self.level_label = tk.Label(self.root, text="Poziom (dB):", font=("Arial", 12))
        self.level_label.place(x=50, y=100)

        self.level_value_label = tk.Label(self.root, text=f"{self.value_dB} dB", font=("Arial", 12))
        self.level_value_label.place(x=150, y=100)

        self.start_button = tk.Button(self.root, text="Rozpocznij pomiar", command=self.start_measurement)
        self.start_button.place(x=150, y=200)

    def setGeneratorParams(self):
        """
        Ustawia napięcie generatora zgodnie z bieżącym poziomem napięcia.
        """
        try:
            print(f"Ustawiam napięcie generatora na: {self.value_voltage:.3f} V RMS")
            APx.BenchMode.Generator.Levels.SetValue(OutputChannelIndex.Ch1, self.value_voltage)
            APx.BenchMode.Generator.On = True
        except Exception as e:
            print(f"Błąd podczas ustawiania generatora: {e}")

    def change_dB(self, new_dB):
        """
        Zmienia poziom dB na zadany poziom i odpowiednio dostosowuje napięcie.
        """
        self.value_dB = new_dB
        self.value_voltage = self.reference_voltage_94dB * (10 ** ((self.value_dB - 94.0) / 20.0))  # Skaluje napięcie względem 94 dB SPL
        self.level_value_label.config(text=f"{self.value_dB:.1f} dB")
        self.setGeneratorParams()

    def ask_and_save(self):
        """
        Pyta użytkownika o wartość i zapisuje ją do pliku Excel.
        """
        # Okno dialogowe z pytaniem o wartość
        user_input = simpledialog.askfloat("Wprowadź wartość", f"Podaj wartość dB zmierzoną przy {self.value_dB} dB:")

        # Jeśli użytkownik podał wartość
        if user_input is not None:
            self.values_list.append((self.value_dB, user_input))  # Zapisz wartość dB i zmierzoną wartość

            # Otwórz i zapisz dane do pliku Excel
            try:
                file_path = 'C:/Users/akust/Desktop/Automatyzacja_Jędrzej/automatyzacja_test.xlsx'
                wb = load_workbook(file_path)  # Wczytaj plik Excel
                ws = wb.active

                # Zapisz wartości do kolumn
                current_row = ws.max_row + 1
                ws[f'A{current_row}'] = self.value_dB  # Zapisz poziom dB
                ws[f'B{current_row}'] = user_input  # Zapisz zmierzoną wartość

                wb.save(file_path)  # Zapisz zmiany
                print(f"Dane zapisane do pliku Excel: {self.value_dB} dB, {user_input} dB")
            except FileNotFoundError:
                print("Nie znaleziono pliku Excel. Upewnij się, że ścieżka jest poprawna.")

    def start_measurement(self):
        """
        Rozpoczyna sekwencję pomiarów od 94 dB do 89 dB.
        """
        self.change_dB(94)  # Ustaw początkowy poziom na 94 dB
        self.ask_and_save()  # Poproś o wartość i zapisz

        self.change_dB(89)  # Zmień poziom na 89 dB
        self.ask_and_save()  # Poproś o wartość i zapisz


# Główna część aplikacji
if __name__ == "__main__":
    root = tk.Tk()
    app = AudioInterfaceApp(root)
    root.mainloop()
    APx.BenchMode.Generator.On = False
