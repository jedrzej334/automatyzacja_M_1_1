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
        self.level_V = 0.1  # Początkowe napięcie, ale potem będzie odczytywane z generatora
        self.value_dB = 0.0  # Początkowy poziom dB

        # Nagłówek
        self.label = tk.Label(self.root, text="Ustawienie poziomu dB", font=("Arial", 14))
        self.label.pack(pady=20)

        # Przycisk do ustawienia napięcia
        self.set_button = tk.Button(self.root, text="Ustaw napięcie dla 80 dB", command=self.set_voltage_for_80dB)
        self.set_button.pack(pady=20)

    def set_voltage_for_80dB(self):
        # Sprawdzenie połączenia z generatorem - próbujemy wykonać operację, aby sprawdzić, czy urządzenie jest dostępne
        try:
            # Próba odczytu wartości napięcia z generatora
            current_voltage_mV = APx.BenchMode.Generator.Levels.GetValue(OutputChannelIndex.Ch1)
            current_voltage = current_voltage_mV / 1000  # Przemiana na Volty
        except Exception as e:
            messagebox.showerror("Błąd połączenia", f"Brak połączenia z urządzeniem APx: {str(e)}")
            return

        # Obliczanie wartości dB dla tego napięcia
        current_dB = self.calculate_dB_from_voltage(current_voltage)

        # Zapytanie o wartość dB, którą chcesz osiągnąć
        target_dB = 80.0
        # Obliczenie nowego napięcia
        voltage_needed = self.calculate_voltage_for_target_dB(current_dB, target_dB)

        # Ustawienie nowego napięcia w generatorze
        try:
            self.set_generator_voltage(voltage_needed)
            # Wyświetlenie komunikatu o ustawieniu napięcia
            messagebox.showinfo("Zmiana napięcia", f"Ustawiono napięcie {voltage_needed:.4f} V, aby uzyskać 80 dB")
        except Exception as e:
            messagebox.showerror("Błąd ustawiania napięcia", f"Nie udało się ustawić napięcia: {str(e)}")

    def calculate_dB_from_voltage(self, voltage):
        # Obliczenie dB na podstawie wzoru: dB = 20 * log10(V2 / V1)
        # Gdzie V2 to napięcie, a V1 to napięcie odniesienia (np. 1V)
        reference_voltage = 1.0  # Załóżmy, że napięcie odniesienia to 1V
        dB = 20 * math.log10(voltage / reference_voltage)
        return dB

    def calculate_voltage_for_target_dB(self, current_dB, target_dB):
        # Przeliczenie napięcia na podstawie wzoru: V2 = V1 * 10^((dB_target - dB_current) / 20)
        reference_voltage = 1.0  # Napięcie odniesienia, np. 1V
        voltage_needed = reference_voltage * 10 ** ((target_dB - current_dB) / 20)
        return voltage_needed

    def set_generator_voltage(self, voltage):
        # Ustawienie napięcia w generatorze APx
        # Zamiana napięcia na jednostki mV (APx oczekuje wartości w milivoltach)
        voltage_mV = voltage * 1000
        APx.BenchMode.Generator.Levels.SetValue(OutputChannelIndex.Ch1, voltage_mV)
        APx.BenchMode.Generator.On = True  # Upewnij się, że generator jest włączony

# Uruchomienie głównej aplikacji
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = AudioInterfaceApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Błąd aplikacji", f"Nie udało się uruchomić aplikacji: {str(e)}")
    finally:
        if APx.BenchMode.Generator.On:
            APx.BenchMode.Generator.On = False  # Wyłączenie generatora po zakończeniu pracy
