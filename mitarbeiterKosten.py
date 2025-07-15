class MitarbeiterKosten:
    def __init__(self, leistung, stunden, satz, gesamtkosten):
        self.leistung = leistung
        self.stunden = stunden
        self.satz = satz
        self.gesamtkosten = gesamtkosten
    
    def __str__(self):
        return f"Leistung: {self.leistung}, Stunden: {self.stunden}, Satz: {self.satz}, Gesamtkosten: {self.gesamtkosten}"