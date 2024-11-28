import tkinter as tk
from tkinter import simpledialog
from openpyxl import load_workbook
import clr

# Dodanie odniesień do bibliotek APx500
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 8.1\\API\\AudioPrecision.API2.dll")
clr.AddReference(r"C:\\Program Files\\Audio Precision\\APx500 8.1\\API\\AudioPrecision.API.dll")

from AudioPrecision.API import APx500_Application, APxOperatingMode, OutputChannelIndex
from System.IO import Directory, Path

# Inicjalizacja aplikacji APx500
APx = APx500_Application(APxOperatingMode.BenchMode, True)

class AudioInterfaceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Interfejs Audio")
        self.root.geometry("500x700")

        ##### Zmienne początkowe #####
        self.level_V = 0.0  # Poziom napięcia
        self.value_dB = 0.0  # Początkowy poziom dB
        self.freqList = [63, 125, 250, 500, 1000, 2000, 4000, 8000, 16000]  # Częstotliwości
        self.freq_var = tk.IntVar(value=0)  # Domyślna częstotliwość: 63 Hz
        self.values_list = []  # Lista do przechowywania wyników

        ##### Tworzenie GUI #####
        # Nagłówek
        self.label = tk.Label(self.root, text="Wybierz częstotliwość", font=("Arial", 14))
        self.label.place(x=150, y=20)

        # Suwak do wyboru częstotliwości
        self.freq_label = tk.Label(self.root, text="Częstotliwość (Hz):", font=("Arial", 12))
        self.freq_label.place(x=50, y=70)

        self.frequency_slider = tk.Scale(self.root, from_=0, to=len(self.freqList) - 1, orient=tk.HORIZONTAL,
                                         label="Częstotliwość (Hz)", variable=self.freq_var, tickinterval=1)
        self.frequency_slider.place(x=50, y=110, width=400)

        self.freq_value_label = tk.Label(self.root, text=f"Wybrana częstotliwość: {self.freqList[self.freq_var.get()]} Hz", font=("Arial", 12))
        self.freq_value_label.place(x=150, y=150)

        self.frequency_slider.bind("<Motion>", self.update_frequency_label)

        # Przycisk do zmiany częstotliwości na 1000 Hz
        self.change_freq_button = tk.Button(self.root, text="Ustaw 1000 Hz", command=lambda: self.set_frequency(1000))
        self.change_freq_button.place(x=150, y=200)

        # Poziom dB
        self.level_label = tk.Label(self.root, text="Poziom (dB):", font=("Arial", 12))
        self.level_label.place(x=50, y=250)

        self.level_value_label = tk.Label(self.root, text=f"{self.value_dB:.1f} dB", font=("Arial", 12))
        self.level_value_label.place(x=150, y=290)

        # Przyciski do zmiany poziomu dB
        self.increase_01_button = tk.Button(self.root, text="Zwiększ o 0.1 dB", command=lambda: self.change_dB(0.1))
        self.increase_01_button.place(x=50, y=330)

        self.decrease_01_button = tk.Button(self.root, text="Zmniejsz o 0.1 dB", command=lambda: self.change_dB(-0.1))
        self.decrease_01_button.place(x=200, y=330)

        # Przycisk do rozpoczęcia pomiarów
        self.start_measurement_button = tk.Button(self.root, text="Rozpocznij pomiar", command=self.start_measurement)
        self.start_measurement_button.place(x=150, y=400)

    ##### Funkcje #####
    def set_generator_params(self, level_dB):
        """
        Ustawia napięcie odpowiadające danemu poziomowi dB.
        """
        self.level_V = 10 ** (level_dB / 20)
        APx.BenchMode.Generator.Levels.SetValue(OutputChannelIndex.Ch1, self.level_V)
        APx.BenchMode.Generator.On = True
        print(f"Generator ustawiony na {self.level_V:.4f} V (dla {level_dB} dB)")

    def set_frequency(self, freq):
        """
        Ustawia częstotliwość generatora.
        """
        if freq in self.freqList:
            self.freq_var.set(self.freqList.index(freq))
            self.freq_value_label.config(text=f"Wybrana częstotliwość: {freq} Hz")
            APx.BenchMode.Generator.Frequency.Value = freq
            print(f"Ustawiono częstotliwość na {freq} Hz")
        else:
            print(f"Częstotliwość {freq} Hz nie jest dostępna.")

    def update_frequency_label(self, event):
        """
        Aktualizuje etykietę wybranej częstotliwości.
        """
        freq = self.freqList[self.freq_var.get()]
        self.freq_value_label.config(text=f"Wybrana częstotliwość: {freq} Hz")

    def change_dB(self, step):
        """
        Zmienia poziom dB o podany krok i aktualizuje generator.
        """
        self.value_dB += step
        self.level_value_label.config(text=f"{self.value_dB:.1f} dB")
        self.set_generator_params(self.value_dB)

    def start_measurement(self):
        """
        Rozpoczyna pomiar od 94 dB do 89 dB i zapisuje wyniki do Excela.
        """
        self.set_frequency(1000)  # Ustaw częstotliwość na 1000 Hz
        self.change_dB(94)  # Rozpocznij od 94 dB
        self.ask_and_save()

        self.change_dB(-5)  # Przejdź do 89 dB
        self.ask_and_save()

    def ask_and_save(self):
        """
        Prosi użytkownika o wartość z miernika i zapisuje do Excela.
        """
        user_input = simpledialog.askfloat("Wprowadź wartość", "Podaj wartość dB z miernika:")
        if user_input is not None:
            self.values_list.append(user_input)
            print(f"Zapisano wartość: {user_input} dB")

            try:
                wb = load_workbook('C:/Users/akust/Desktop/Automatyzacja_Jędrzej/automatyzacja_test.xlsx')
                ws = wb.active
                for i, value in enumerate(self.values_list, start=1):
                    ws[f'P{i}'] = value
                wb.save('C:/Users/akust/Desktop/Automatyzacja_Jędrzej/automatyzacja_test.xlsx')
                print("Dane zapisane do Excela.")
            except FileNotFoundError:
                print("Nie znaleziono pliku Excel.")

# Główna część aplikacji
if __name__ == "__main__":
    root = tk.Tk()
    app = AudioInterfaceApp(root)
    root.mainloop()
    APx.BenchMode.Generator.On = False
