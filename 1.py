    def setGeneratorParams(self, level_dB):
        """
        Ustaw poziom generatora w oparciu o podany poziom w dB.
        """
        try:
            # Konwersja z dB na napięcie RMS w Voltach
            level_V = 10 ** (level_dB / 20.0)  # dBV -> Voltage
            print(f"Ustawiam poziom generatora na: {level_V:.3f} V RMS")

            # Ustawianie poziomu generatora na kanał 1
            APx.BenchMode.Generator.Levels.SetValue(OutputChannelIndex.Ch1, level_V)
            APx.BenchMode.Generator.On = True
        except Exception as e:
            print(f"Błąd podczas ustawiania generatora: {e}")

    def change_dB(self, step):
        """
        Zmienia poziom dB i aktualizuje generator.
        """
        self.value_dB += step
        if self.value_dB < 0:
            self.value_dB = 0
        self.level_value_label.config(text=f"{self.value_dB:.1f} dB")

        # Aktualizacja generatora
        self.setGeneratorParams(self.value_dB)
