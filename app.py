import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET
import uuid
from xml.dom import minidom

class XMLGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Pozyczki migrator xml")

        self.label1 = tk.Label(root, text="Plik glowny source, zawierajacy wszyskie dane:")
        self.label1.pack(pady=5)
        self.file_path1 = tk.Entry(root, width=50)
        self.file_path1.pack(pady=5)
        self.browse_button1 = tk.Button(root, text="Pliki", command=self.browse_file1)
        self.browse_button1.pack(pady=5)

        self.label2 = tk.Label(root, text="Plik z kodem i guidem:")
        self.label2.pack(pady=5)
        self.file_path2 = tk.Entry(root, width=50)
        self.file_path2.pack(pady=5)
        self.browse_button2 = tk.Button(root, text="Pliki", command=self.browse_file2)
        self.browse_button2.pack(pady=5)

        self.process_button = tk.Button(root, text="Generuj pozyczki do XML", command=self.process_files)
        self.process_button.pack(pady=20)

    def browse_file1(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_path1.delete(0, tk.END)
            self.file_path1.insert(0, file_path)

    def browse_file2(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_path2.delete(0, tk.END)
            self.file_path2.insert(0, file_path)

    def process_files(self):
        file1_path = self.file_path1.get()
        file2_path = self.file_path2.get()

        if not file1_path or not file2_path:
            messagebox.showerror("Error", "Wymagane sa oba pliki")
            return

        try:
            plik1 = pd.read_excel(file1_path)
            plik2 = pd.read_excel(file2_path)

            def format_code(kod):
                return '9' + str(kod).zfill(6)[-6:]

            plik1['Żyrant 1 KOD'] = plik1['Żyrant 1 KOD'].apply(format_code)
            plik1['Żyrant 2 KOD'] = plik1['Żyrant 2 KOD'].apply(format_code)
            plik1['Pracownik'] = plik1['Pracownik'].apply(format_code)

            plik2['Kod'] = plik2['Kod'].apply(format_code)

            merged_zyrant1 = plik1.merge(plik2[['Kod', 'Guid']], left_on='Żyrant 1 KOD', right_on='Kod', how='left')
            plik1['Żyrant_1_GUID'] = merged_zyrant1['Guid']

            merged_zyrant2 = plik1.merge(plik2[['Kod', 'Guid']], left_on='Żyrant 2 KOD', right_on='Kod', how='left')
            plik1['Żyrant_2_GUID'] = merged_zyrant2['Guid']

            merged_pracownik = plik1.merge(plik2[['Kod', 'Guid']], left_on='Pracownik', right_on='Kod', how='left')
            plik1['Pracownik_GUID'] = merged_pracownik['Guid']

            output_path = "wynik.xlsx"
            plik1.to_excel(output_path, index=False)
            messagebox.showinfo("Success", "Zapiasno wynik.xlsx.")

            self.run_second_process()

        except Exception as e:
            messagebox.showerror("Error", f"{e}")

    def run_second_process(self):
        try:
            df = pd.read_excel("wynik.xlsx")

            
            model_xml_content = """<?xml version="1.0" encoding="Unicode"?>
<session xmlns="http://www.soneta.pl/schema/business">
  
</session>"""

            model_temp_xml_content = """<?xml version="1.0" encoding="utf-8"?>
<session xmlns="http://www.soneta.pl/schema/business">
  
</session>"""

            xml_content = model_xml_content.replace('encoding="Unicode"', 'encoding="utf-8"')

            tree = ET.ElementTree(ET.fromstring(xml_content))
            root = tree.getroot()
            if root.tag.startswith('{http://www.soneta.pl/schema/business}'):
                root.tag = root.tag.split('}', 1)[1]
            root.attrib['xmlns:ns0'] = "http://www.soneta.pl/schema/business"

            def generate_id():
                return str(uuid.uuid4())

            def create_fund_pozyczkowy(record):
                fund_id = generate_id()
                fund = ET.Element('FundPozyczkowy', id=f'FundPozyczkowy_{fund_id}')

                pracownik = ET.SubElement(fund, 'Pracownik')
                pracownik.text = str(record["Pracownik_GUID"])

                definicja = ET.SubElement(fund, 'Definicja')
                definicja.text = str("e88c19a4-ec01-47f6-9c6f-85730ca4434e")

                okres = ET.SubElement(fund, 'Okres')
                okres.text = str(record['Okres']) + ".."

                saldo_bo = ET.SubElement(fund, 'SaldoBO')
                saldo_bo.text = f"{record['SaldoBO']} PLN"

                ET.SubElement(fund, 'FundPozyczkowyExtension')

                extension2 = ET.SubElement(fund, 'FundPozyczkowyExtension')
                extension_inner = ET.SubElement(extension2, 'FundPozyczkowyExtension', id=f'FundPozyczkowyExtension_{generate_id()}')

                host2 = ET.SubElement(extension_inner, 'Host')
                host2.text = f'FundPozyczkowy_{fund_id}'

                return fund, fund_id

            def create_zyrant_pozyczki(pozyczka_id, record, czy_zyrant_2):
                zyrant_pozyczki = ET.Element('ŻyrantPożyczki', id=f'ŻyrantPożyczki_{generate_id()}')

                pozyczka_ref = ET.SubElement(zyrant_pozyczki, 'Pozyczka')
                pozyczka_ref.text = f'Pozyczka_{pozyczka_id}'

                pracownik = ET.SubElement(zyrant_pozyczki, 'Pracownik')
                pracownik.text = str(record["Żyrant_1_GUID"])
                if czy_zyrant_2: pracownik.text = str(record["Żyrant_2_GUID"])

                priorytet = ET.SubElement(zyrant_pozyczki, 'Priorytet')
                priorytet.text = '1'

                splaty_od = ET.SubElement(zyrant_pozyczki, 'SplatyOd')
                splaty_od.text = '(pusty)'

                element_raty = ET.SubElement(zyrant_pozyczki, 'ElementRaty')

                kwota = ET.SubElement(zyrant_pozyczki, 'Kwota')
                kwota.text = '0.00 PLN'

                procent = ET.SubElement(zyrant_pozyczki, 'Procent')
                procent.text = '0.00%'

                return zyrant_pozyczki

            def create_pozyczki(record, fund_id):
                pozyczka_id = generate_id()
                pozyczka = ET.Element('Pozyczka', id=f'Pozyczka_{pozyczka_id}')

                fundusz = ET.SubElement(pozyczka, 'Fundusz')
                fundusz.text = f'FundPozyczkowy_{fund_id}'

                data = ET.SubElement(pozyczka, "Data")
                data.text = str(record['Data'])

                stan = ET.SubElement(pozyczka, "Stan")
                stan.text = "NieSpłacona"

                kwota = ET.SubElement(pozyczka, "Kwota")
                kwota.text = f"{record['Kwota']} PLN"

                element = ET.SubElement(pozyczka, "Element")
                element.text = "49c56125-e7f5-4b5a-b858-8701e7d304f6"

                splacona = ET.SubElement(pozyczka, "Splacona")
                splacona.text = "False"

                zyrant1 = ET.SubElement(pozyczka, "Zyrant1")
                zyrant1.text = ""

                zyrant2 = ET.SubElement(pozyczka, "Zyrant2")
                ET.SubElement(zyrant2, 'Osoba')
                ET.SubElement(zyrant2, 'Adres')
                ET.SubElement(zyrant2, 'Telefon')

                splaty_od = ET.SubElement(pozyczka, "SplatyOd")
                splaty_od.text = str(record['SplatyOd'])

                ilosc_rat = ET.SubElement(pozyczka, "IloscRat")
                ilosc_rat.text = str(record['IloscRat'])

                kwota_raty = ET.SubElement(pozyczka, "KwotaRaty")
                kwota_raty.text = f"{record['KwotaRaty']} PLN"

                splata_roznicy = ET.SubElement(pozyczka, "SplataRoznicy")
                splata_roznicy.text = "OstatniąRatą"

                element_raty = ET.SubElement(pozyczka, "ElementRaty")
                element_raty.text = "e90e6b4c-a5f6-4077-8f3a-de35d493a98d"

                typ = ET.SubElement(pozyczka, "Typ")
                typ.text = str(record['Typ'])
                

                sposob = ET.SubElement(pozyczka, "Sposob")
                sposob.text = str(record['Sposob'])

                procent = ET.SubElement(pozyczka, "Procent")
                procent.text = str(record['Procent'])

                odsetki_za_odroczenie = ET.SubElement(pozyczka, "OdsetkiZaOdroczenie")
                odsetki_za_odroczenie.text = "False"

                algorytm_raty = ET.SubElement(pozyczka, "AlgorytmRaty")
                algorytm_raty.text = "7774fbd9-9b5b-436d-8fac-58ba83921a4c"
                bilans_otwarcia = ET.SubElement(pozyczka, "BilansOtwarcia")
                bilans_otwarcia.text = "False"
                indywidualny_rachunek_bankowy = ET.SubElement(pozyczka, "IndywidualnyRachunekBankowy")

                pozyczka_extension = ET.SubElement(pozyczka, 'PozyczkaExtension', id=f'PozyczkaExtension_{generate_id()}')
                host = ET.SubElement(pozyczka_extension, 'Host')
                host.text = f'Pozyczka_{pozyczka_id}'
                numer_pozyczki = ET.SubElement(pozyczka_extension, 'NumerPozyczki')
                uwagi = ET.SubElement(pozyczka_extension, 'Uwagi')

                zyranci = ET.SubElement(pozyczka, 'Żyranci')
                zyranci.append(create_zyrant_pozyczki(pozyczka_id, record, False))
                zyranci.append(create_zyrant_pozyczki(pozyczka_id, record, True))
                return pozyczka

            for _, row in df.iterrows():
                new_fund, fund_id = create_fund_pozyczkowy(row)
                root.append(new_fund)
                new_pozyczka = create_pozyczki(row, fund_id)
                root.append(new_pozyczka)

            def prettify(elem):
                rough_string = ET.tostring(elem, 'utf-8')
                reparsed = minidom.parseString(rough_string)
                return reparsed.toprettyxml(indent="  ")

            pretty_xml_as_string = prettify(root)
            with open("dane.xml", "w", encoding='utf-8') as f:
                f.write(pretty_xml_as_string)

            messagebox.showinfo("Success", "Zapisano dane.xml.")

        except Exception as e:
            messagebox.showerror("Error", f": {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = XMLGeneratorGUI(root)
    root.mainloop()
