import tkinter as tk
from tkinter import simpledialog
from openpyxl import load_workbook
import sys
import clr
import time

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
        self.value_dB = 0.0  # Początkowy poziom dB
        self.value_voltage = 1.0  # Początkowe napięcie w Voltach RMS (1 V RMS = 0 dBV)
        self.freqList = [63, 125, 250, 500, 1000, 2000, 4000, 8000, 16000]  # Częstotliwości
        self.freq_var = tk.IntVar(value=0)  # Domyślna częstotliwość: 63 Hz
        self.values_list = []  # Lista do przechowywania wyników

        # Tworzenie GUI
        self.label = tk.Label(self.root, text="Wybierz częstotliwość", font=("Arial", 14))
        self.label.place(x=150, y=20)

        # Suwak do wyboru częstotliwości
        self.freq_label = tk.Label(self.root, text="Częstotliwość (Hz):", font=("Arial", 12))
        self.freq_label.place(x=50, y=70)

        self.frequency_slider = tk.Scale(self.root, from_=0, to=len(self.freqList)-1, orient=tk.HORIZONTAL,
                                          label="Częstotliwość (Hz)", variable=self.freq_var,
                                          tickinterval=1)
        self.frequency_slider.place(x=50, y=110, width=400)

        self.freq_value_label = tk.Label(self.root, text=f"Wybrana częstotliwość: {self.freqList[self.freq_var.get()]} Hz", font=("Arial", 12))
        self.freq_value_label.place(x=150, y=150)

        self.frequency_slider.bind("<Motion>", self.update_frequency_label)

        # Okno do wyświetlania poziomu dB
        self.level_label = tk.Label(self.root, text="Poziom (dB):", font=("Arial", 12))
        self.level_label.place(x=50, y=200)

        self.level_value_label = tk.Label(self.root, text=f"{self.value_dB} dB", font=("Arial", 12))
        self.level_value_label.place(x=150, y=240)

        # Przyciski do zmiany poziomu dB
        self.increase_01_button = tk.Button(self.root, text="Zwiększ o 0.1 dB", command=lambda: self.change_dB(0.1))
        self.increase_01_button.place(x=50, y=280)

        self.decrease_01_button = tk.Button(self.root, text="Zmniejsz o 0.1 dB", command=lambda: self.change_dB(-0.1))
        self.decrease_01_button.place(x=200, y=280)

        # Przyciski dodatkowe
        self.increase_3_2_button = tk.Button(self.root, text="+3.2", command=lambda: self.change_dB(3.2))
        self.increase_3_2_button.place(x=50, y=460)

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

    def change_dB(self, step):
        """
        Zmienia poziom dB i odpowiednio dostosowuje napięcie.
        """
        # Oblicz nowe napięcie
        new_voltage = self.value_voltage * (10 ** (step / 20.0))
        self.value_voltage = max(new_voltage, 0.001)  # Upewnij się, że napięcie nie spada poniżej sensownego minimum

        # Aktualizacja poziomu dB
        self.value_dB += step
        if self.value_dB < 0:
            self.value_dB = 0

        # Aktualizacja etykiety GUI
        self.level_value_label.config(text=f"{self.value_dB:.1f} dB")

        # Aktualizacja generatora
        self.setGeneratorParams()

    def update_frequency_label(self, event):
        """
        Aktualizuje częstotliwość generatora i GUI w czasie rzeczywistym.
        """
        self.freq_value_label.config(text=f"Wybrana częstotliwość: {self.freqList[self.freq_var.get()]} Hz")
        currentFreq = self.freqList[self.freq_var.get()]
        APx.BenchMode.Generator.Frequency.Value = currentFreq


# Główna część aplikacji
if __name__ == "__main__":
    root = tk.Tk()
    app = AudioInterfaceApp(root)
    root.mainloop()
    APx.BenchMode.Generator.On = False
