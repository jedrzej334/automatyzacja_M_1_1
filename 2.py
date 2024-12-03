import tkinter as tk
from tkinter import simpledialog, messagebox
import math
import clr

# Załadowanie bibliotek APx
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 8.1\\API\\AudioPrecision.API2.dll")
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 8.1\\API\\AudioPrecision.API.dll")

from AudioPrecision.API import *
from System.Drawing import Point
from System.Windows.Forms import Application, Button, Form, Label
from System.IO import Directory, Path

# Inicjalizacja aplikacji APx
APx = APx500_Application(APxOperatingMode.BenchMode, True)

class AudioInterfaceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ustawienie Napięcia i dB")
        self.root.geometry("400x300")

        # Początkowe ustawienie napięcia (w Voltach)
        self.initial_voltage = 0.1  # Początkowe napięcie (w Voltach)
        self.level_V = self.initial_voltage  # Aktualne napięcie
        self.value_dB = 0.0  # Początkowy poziom dB

        # Nagłówek
        self.label = tk.Label(self.root, text="Ustawienie poziomu dB", font=("Arial", 14))
        self.label.pack(pady=20)

        # Przycisk do ustawienia napięcia
        self.set_button = tk.Button(self.root, text="Ustaw napięcie dla 80 dB", command=self.set_voltage_for_80dB)
        self.set_button.pack(pady=20)

    def set_voltage_for_80dB(self):
        # Zapytanie o aktualny poziom dB z miernika
        current_dB = simpledialog.askfloat("Wartość dB z miernika", "Podaj aktualną wartość dB z miernika:", minvalue=-100.0, maxvalue=100.0)
        
        if current_dB is None:
            return  # Jeśli nie podano wartości dB, zakończ

        # Obliczanie napięcia, aby uzyskać 80 dB
        target_dB = 80.0
        voltage_needed = self.calculate_voltage_for_target_dB(current_dB, target_dB)

        # Ustawienie napięcia w generatorze
        self.set_generator_voltage(voltage_needed)

        # Wyświetlenie komunikatu o ustawieniu napięcia
        messagebox.showinfo("Zmiana napięcia", f"Ustawiono napięcie {voltage_needed:.4f} V, aby uzyskać 80 dB")

    def calculate_voltage_for_target_dB(self, current_dB, target_dB):
        # Przeliczenie napięcia na podstawie wzoru: V2 = V1 * 10^((dB_target - dB_current) / 20)
        voltage_needed = self.level_V * 10 ** ((target_dB - current_dB) / 20)
        return voltage_needed

    def set_generator_voltage(self, voltage):
        # Ustawienie napięcia w generatorze APx
        # Zamiana napięcia na jednostki mV (APx oczekuje wartości w milivoltach)
        voltage_mV = voltage * 1000
        APx.BenchMode.Generator.Levels.SetValue(OutputChannelIndex.Ch1, voltage_mV)
        APx.BenchMode.Generator.On = True  # Upewnij się, że generator jest włączony

# Uruchomienie głównej aplikacji
if __name__ == "__main__":
    root = tk.Tk()
    app = AudioInterfaceApp(root)
    root.mainloop()
    APx.BenchMode.Generator.On = False  # Wyłączenie generatora po zakończeniu pracy
