import pandas as pd
import os.path
import datetime
import tkinter.messagebox
import customtkinter
import pandastable as pt
from CTkMessagebox import CTkMessagebox
import os.path

class App(customtkinter.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.iconbitmap('_internal\icon.ico')
        self.title("Sandėlio programa")
        self.geometry("410x470")
        self.resizable(width=False, height=False)
        customtkinter.set_default_color_theme("blue")
        customtkinter.set_appearance_mode("dark")

        self.button_1 = customtkinter.CTkButton(self, text="Sandėlio papildymas", command=self.papildyti_sandeli, width=200, height=30)
        self.button_1.pack(side="top", pady=10)

        self.button_2 = customtkinter.CTkButton(self, text="Perkelti medžiagas į gamybą", command=self.perkelti_is_sandelio, width=200, height=30)
        self.button_2.pack(side="top", pady=10)

        self.button_3 = customtkinter.CTkButton(self, text="Medžiagų likučiai sandėlyje", command=self.sandelio_likutis, width=200, height=30)
        self.button_3.pack(side="top", pady=10)

        self.button_4 = customtkinter.CTkButton(self, text="Medžiagų likučiai gamyboje", command=self.gamybos_likutis, width=200, height=30)
        self.button_4.pack(side="top", pady=10)

        self.button_5 = customtkinter.CTkButton(self, text="Gamybos darbų registras", command=self.darbu_registras, width=200, height=30)
        self.button_5.pack(side="top", pady=10)

        self.button_9 = customtkinter.CTkButton(self, text="Peržiūrėti gamybos darbų registrą", command=self.gamybos_registras, width=200, height=30)
        self.button_9.pack(side="top", pady=10)

        self.button_6 = customtkinter.CTkButton(self, text="Išeiti", command=self.destroy, width=200, height=30)
        self.button_6.pack(side="top", pady=10)

        self.button_7 = customtkinter.CTkSwitch(self, text="Pakeisti likučių lango vaizdą", onvalue="on", offvalue="off")
        self.button_7.place(x=15, y=430)

        self.button_8 = customtkinter.CTkSwitch(self, text="Šviesusis režimas", command=self.mode, onvalue="light", offvalue="dark")
        self.button_8.place(x=255, y=430)

        self.toplevel_window = None

    def mode(self):
        if self.button_8.get() == "light":
            customtkinter.set_appearance_mode("light")
        elif self.button_8.get() == "dark":
            customtkinter.set_appearance_mode("dark")

    def sandelio_likutis(self):

        if os.path.isfile('Sandėlio istorija.xlsx'):

            df = pd.read_excel('Sandėlio istorija.xlsx')
            df1 = df.groupby('Partijos numeris')['Polietileno plėvelė 0.85x50m'].sum().round(2).reset_index()
            df2 = df.groupby('Partijos numeris')['Hidroizoliacinė plėvelė 0.65x45m'].sum().round(2).reset_index()
            df2.drop(columns=['Partijos numeris'], inplace=True)
            df3 = df.groupby('Partijos numeris')['Sausas tinko mišinys 2.5kg'].sum().round(2).reset_index()
            df3.drop(columns=['Partijos numeris'], inplace=True)
            df4 = df.groupby('Partijos numeris')['Cementinis mišinys 4kg'].sum().round(2).reset_index()
            df4.drop(columns=['Partijos numeris'], inplace=True)

            if self.button_7.get() == "on":
                df_likuciai = pd.concat([df1, df2, df3, df4], axis=1)
                df_likuciai.set_index('Partijos numeris', inplace=True)
                df_likuciai = df_likuciai.transpose().reset_index()
                df_likuciai.rename(columns={'index': 'Medžiaga'}, inplace=True)
                df_likuciai = df_likuciai.melt(id_vars=["Medžiaga"], var_name="Partijos numeris", value_name="Likutis")
                df_likuciai = df_likuciai[df_likuciai.Likutis != 0]

                dTDa1 = tkinter.Toplevel()
                dTDa1.iconbitmap('_internal\icon.ico')
                dTDa1.title('Medžiagų likučiai sandėlyje')
                dTDaPT = pt.Table(dTDa1, dataframe=df_likuciai, showtoolbar=True, showstatusbar=True)
                dTDaPT.show()

            elif self.button_7.get() == "off":

                df_likuciai = pd.concat([df1, df2, df3, df4], axis=1)

                dTDa1 = tkinter.Toplevel()
                dTDa1.iconbitmap('_internal\icon.ico')
                dTDa1.title('Medžiagų likučiai sandėlyje')
                dTDaPT = pt.Table(dTDa1, dataframe=df_likuciai, showtoolbar=True, showstatusbar=True)
                dTDaPT.show()

        else:
            CTkMessagebox(title="Error", message="Patikrinkite failą 'Sandėlio istorija'!", icon="cancel")

    def gamybos_likutis(self):

        if os.path.isfile('Gamybos medžiagų istorija.xlsx'):

            df = pd.read_excel('Gamybos medžiagų istorija.xlsx')
            df1 = df.groupby('Medžiagos ID')['Polietileno plėvelė, m'].sum().round(2).reset_index()
            df2 = df.groupby('Medžiagos ID')['Hidroizoliacinė plėvelė, m'].sum().round(2).reset_index()
            df2.drop(columns=['Medžiagos ID'], inplace=True)
            df3 = df.groupby('Medžiagos ID')['Sausas tinko mišinys, kg'].sum().round(2).reset_index()
            df3.drop(columns=['Medžiagos ID'], inplace=True)
            df4 = df.groupby('Medžiagos ID')['Cementinis mišinys, kg'].sum().round(2).reset_index()
            df4.drop(columns=['Medžiagos ID'], inplace=True)

            if self.button_7.get() == "on":
                df_likuciai = pd.concat([df1, df2, df3, df4], axis=1)
                df_likuciai.set_index('Medžiagos ID', inplace=True)
                df_likuciai = df_likuciai.transpose().reset_index()
                df_likuciai.rename(columns={'index': 'Medžiaga'}, inplace=True)
                df_likuciai = df_likuciai.melt(id_vars=["Medžiaga"], var_name="Medžiagos ID", value_name="Likutis")
                df_likuciai = df_likuciai[df_likuciai.Likutis != 0]

                dTDa1 = tkinter.Toplevel()
                dTDa1.iconbitmap('_internal\icon.ico')
                dTDa1.title('Medžiagų likučiai gamyboje')
                dTDaPT = pt.Table(dTDa1, dataframe=df_likuciai, showtoolbar=True, showstatusbar=True)
                dTDaPT.show()

            elif self.button_7.get() == "off":
                df_likuciai = pd.concat([df1, df2, df3, df4], axis=1)

                dTDa1 = tkinter.Toplevel()
                dTDa1.iconbitmap('_internal\icon.ico')
                dTDa1.title('Medžiagų likučiai gamyboje')
                dTDaPT = pt.Table(dTDa1, dataframe=df_likuciai, showtoolbar=True, showstatusbar=True)
                dTDaPT.show()
        else:
            CTkMessagebox(title="Error", message="Patikrinkite failą 'Gamybos medžiagų istorija'!", icon="cancel")

    def gamybos_registras(self):

        if os.path.isfile('Gamybos darbų registras.xlsx'):

            df = pd.read_excel('Gamybos darbų registras.xlsx')
            df['Papildoma informacija'] = df['Papildoma informacija'].astype(str)
            df1 = df.groupby('Užsakymo numeris')['Polietileno plėvelė, m'].sum().round(2).reset_index()
            df2 = df.groupby('Užsakymo numeris')['Hidroizoliacinė plėvelė, m'].sum().round(2).reset_index()
            df2.drop(columns=['Užsakymo numeris'], inplace=True)
            df3 = df.groupby('Užsakymo numeris')['Sausas tinko mišinys, kg'].sum().round(2).reset_index()
            df3.drop(columns=['Užsakymo numeris'], inplace=True)
            df4 = df.groupby('Užsakymo numeris')['Cementinis mišinys, kg'].sum().round(2).reset_index()
            df4.drop(columns=['Užsakymo numeris'], inplace=True)
            df5 = df.groupby('Užsakymo numeris')['Papildoma informacija'].apply(lambda x: ', '.join(x)).reset_index()
            df5 = df5['Papildoma informacija'].str.replace('nan,', '')
            df5 = df5.str.replace('nan', '')

            if self.button_7.get() == "on":
                df_likuciai = pd.concat([df1, df2, df3, df4], axis=1)
                df_likuciai.set_index('Užsakymo numeris', inplace=True)
                df_likuciai = df_likuciai.transpose().reset_index()
                df_likuciai.rename(columns={'index': 'Medžiaga'}, inplace=True)
                df_likuciai = df_likuciai.melt(id_vars=["Medžiaga"], var_name="Užsakymo numeris", value_name="Sunaudota")
                df_likuciai = df_likuciai[df_likuciai.Sunaudota != 0]

                dTDa1 = tkinter.Toplevel()
                dTDa1.iconbitmap('_internal\icon.ico')
                dTDa1.title('Gamybos darbų registras')
                dTDaPT = pt.Table(dTDa1, dataframe=df_likuciai, showtoolbar=True, showstatusbar=True)
                dTDaPT.show()

            elif self.button_7.get() == "off":
                df_likuciai = pd.concat([df1, df2, df3, df4, df5], axis=1)
                dTDa1 = tkinter.Toplevel()
                dTDa1.iconbitmap('_internal\icon.ico')
                dTDa1.title('Gamybos darbų registras')
                dTDaPT = pt.Table(dTDa1, dataframe=df_likuciai, showtoolbar=True, showstatusbar=True)
                dTDaPT.show()
        else:
            CTkMessagebox(title="Error", message="Patikrinkite failą 'Gamybos darbų registras'!", icon="cancel")

    def papildyti_sandeli(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = Pridejimo_langas()
        else:
            self.toplevel_window.focus()

    def perkelti_is_sandelio(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = Nurasymo_is_sandelio_langas()

        else:
            self.toplevel_window.focus()

    def darbu_registras(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = Darbu_registro_langas()
        else:
            self.toplevel_window.focus()

class Pridejimo_langas(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.after(100, lambda: self.focus())
        self.after(205, lambda: self.iconbitmap('_internal\icon.ico'))
        self.geometry("400x400")
        self.resizable(width=False, height=False)
        self.title("Sandėlio papildymas")
        self.label = customtkinter.CTkLabel(self, text="Pasirinkite pildomą medžiagą")
        self.label.pack(padx=10, pady=10)

        self.button_1 = customtkinter.CTkComboBox(self, values=['Polietileno plėvelė', 'Hidroizoliacinė plėvelė', 'Sausas tinko mišinys', 'Cementinis mišinys'])
        self.button_1.pack(side="top", padx=20, pady=10)

        self.button_2 = customtkinter.CTkEntry(self, placeholder_text="Įveskite pildomą kiekį")
        self.button_2.pack(side="top", padx=20, pady=10)

        self.button_3 = customtkinter.CTkEntry(self, placeholder_text="Partijos numeris")
        self.button_3.pack(side="top", padx=20, pady=10)

        self.button_4 = customtkinter.CTkEntry(self, placeholder_text="Atsakingas asmuo")
        self.button_4.pack(side="top", padx=20, pady=10)

        self.button_7 = customtkinter.CTkEntry(self, placeholder_text="Papildoma informacija")
        self.button_7.pack(side="top", padx=20, pady=10)

        self.button_5 = customtkinter.CTkButton(self, text="Pildyti", command=self.pildymas)
        self.button_5.pack(side="top", padx=20, pady=10)

    def pildymas(self):
        naudotas = self.button_2.get()
        numeris = self.button_3.get()
        zmogus = self.button_4.get()
        komentaras = self.button_7.get()
        if self.button_1.get() == 'Polietileno plėvelė' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "":
            medziaga = self.button_1.get()
            prideti_medziaga(medziaga, naudotas, numeris, zmogus, komentaras)
        elif self.button_1.get() == 'Hidroizoliacinė plėvelė' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "":
            medziaga = self.button_1.get()
            prideti_medziaga(medziaga, naudotas, numeris, zmogus, komentaras)
        elif self.button_1.get() == 'Sausas tinko mišinys' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "":
            medziaga = self.button_1.get()
            prideti_medziaga(medziaga, naudotas, numeris, zmogus, komentaras)
        elif self.button_1.get() == 'Cementinis mišinys' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "":
            medziaga = self.button_1.get()
            prideti_medziaga(medziaga, naudotas, numeris, zmogus, komentaras)
        else:
            CTkMessagebox(title="Error", message="Patikrinkite įvestus duomenis!", icon="cancel")

class Nurasymo_is_sandelio_langas(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.after(100, lambda: self.focus())
        self.after(205, lambda: self.iconbitmap('_internal\icon.ico'))
        self.geometry("400x450")
        self.resizable(width=False, height=False)

        self.title("Perkelti medžiagas į gamybą")
        self.label = customtkinter.CTkLabel(self, text="Pasirinkite perkeliamą medziagą")
        self.label.pack(padx=10, pady=10)

        self.button_1 = customtkinter.CTkComboBox(self, values=['Polietileno plėvelė', 'Hidroizoliacinė plėvelė', 'Sausas tinko mišinys', 'Cementinis mišinys'])
        self.button_1.pack(side="top", padx=20, pady=10)

        self.button_2 = customtkinter.CTkEntry(self, placeholder_text="Kiekis")
        self.button_2.pack(side="top", padx=20, pady=10)

        self.button_3 = customtkinter.CTkEntry(self, placeholder_text="Partijos numeris")
        self.button_3.pack(side="top", padx=20, pady=10)

        self.button_7 = customtkinter.CTkEntry(self, placeholder_text="Medžiagos ID")
        self.button_7.pack(side="top", padx=20, pady=10)

        self.button_4 = customtkinter.CTkEntry(self, placeholder_text="Atsakingas asmuo")
        self.button_4.pack(side="top", padx=20, pady=10)

        self.button_8 = customtkinter.CTkEntry(self, placeholder_text="Papildoma informacija")
        self.button_8.pack(side="top", padx=20, pady=10)

        self.button_5 = customtkinter.CTkButton(self, text="Perkelti", command=self.perkelimas)
        self.button_5.pack(side="top", padx=20, pady=10)

    def perkelimas(self):
        naudotas = self.button_2.get()
        numeris = self.button_3.get()
        zmogus = self.button_4.get()
        komentaras = self.button_8.get()
        serijos_nr_spaudejo = self.button_7.get()
        if self.button_1.get() == 'Polietileno plėvelė' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "" and self.button_7.get() != "":
            medziaga = self.button_1.get()
            atimti_medziaga_is_sandelio(medziaga, naudotas, numeris, zmogus, serijos_nr_spaudejo, komentaras)
        elif self.button_1.get() == 'Hidroizoliacinė plėvelė' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "" and self.button_7.get() != "":
            medziaga = self.button_1.get()
            atimti_medziaga_is_sandelio(medziaga, naudotas, numeris, zmogus, serijos_nr_spaudejo, komentaras)
        elif self.button_1.get() == 'Sausas tinko mišinys' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "" and self.button_7.get() != "":
            medziaga = self.button_1.get()
            atimti_medziaga_is_sandelio(medziaga, naudotas, numeris, zmogus, serijos_nr_spaudejo, komentaras)
        elif self.button_1.get() == 'Cementinis mišinys' and self.button_2.get() != "" and self.button_3.get() != "" and self.button_4.get() != "" and self.button_7.get() != "":
            medziaga = self.button_1.get()
            atimti_medziaga_is_sandelio(medziaga, naudotas, numeris, zmogus, serijos_nr_spaudejo, komentaras)
        else:
            CTkMessagebox(title="Error", message="Patikrinkite įvestus duomenis!", icon="cancel")

class Darbu_registro_langas(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.after(100, lambda: self.focus())
        self.after(205, lambda: self.iconbitmap('_internal\icon.ico'))
        self.geometry("400x450")
        self.resizable(width=False, height=False)

        self.title("Gamybos darbų registras")
        self.label = customtkinter.CTkLabel(self, text="Pasirinkite sunaudotą medžiagą")
        self.label.pack(padx=10, pady=10)

        self.button_1 = customtkinter.CTkComboBox(self, values=['Polietileno plėvelė', 'Hidroizoliacinė plėvelė', 'Sausas tinko mišinys', 'Cementinis mišinys'])
        self.button_1.pack(side="top", padx=20, pady=10)

        self.button_2 = customtkinter.CTkEntry(self, placeholder_text="Įveskite kiekį")
        self.button_2.pack(side="top", padx=20, pady=10)

        self.button_7 = customtkinter.CTkEntry(self, placeholder_text="Medžiagos ID")
        self.button_7.pack(side="top", padx=20, pady=10)

        self.button_3 = customtkinter.CTkEntry(self, placeholder_text="Užsakymo numeris")
        self.button_3.pack(side="top", padx=20, pady=10)

        self.button_4 = customtkinter.CTkEntry(self, placeholder_text="Atsakingas asmuo")
        self.button_4.pack(side="top", padx=20, pady=10)

        self.button_8 = customtkinter.CTkEntry(self, placeholder_text="Papildoma informacija")
        self.button_8.pack(side="top", padx=20, pady=10)

        self.button_5 = customtkinter.CTkButton(self, text="Registruoti", command=self.nurasymas)
        self.button_5.pack(side="top", padx=20, pady=10)

    def nurasymas(self):
        naudotas = self.button_2.get()
        zmogus = self.button_4.get()
        serijos_nr_spaudejo = self.button_7.get()
        uzsakymo_nr = self.button_3.get()
        komentaras = self.button_8.get()
        if self.button_1.get() == 'Polietileno plėvelė' and self.button_2.get() != "" and self.button_4.get() != "" and self.button_7.get() != "" and self.button_3.get() != "":
            medziaga = self.button_1.get()
            uzsakyme_panaudota_medziaga(medziaga, naudotas, zmogus, serijos_nr_spaudejo, uzsakymo_nr, komentaras)
        elif self.button_1.get() == 'Hidroizoliacinė plėvelė' and self.button_2.get() != "" and self.button_4.get() != "" and self.button_7.get() != "" and self.button_3.get() != "":
            medziaga = self.button_1.get()
            uzsakyme_panaudota_medziaga(medziaga, naudotas, zmogus, serijos_nr_spaudejo, uzsakymo_nr, komentaras)
        elif self.button_1.get() == 'Sausas tinko mišinys' and self.button_2.get() != "" and self.button_4.get() != "" and self.button_7.get() != "" and self.button_3.get() != "":
            medziaga = self.button_1.get()
            uzsakyme_panaudota_medziaga(medziaga, naudotas, zmogus, serijos_nr_spaudejo, uzsakymo_nr, komentaras)
        elif self.button_1.get() == 'Cementinis mišinys' and self.button_2.get() != "" and self.button_4.get() != "" and self.button_7.get() != "" and self.button_3.get() != "":
            medziaga = self.button_1.get()
            uzsakyme_panaudota_medziaga(medziaga, naudotas, zmogus, serijos_nr_spaudejo, uzsakymo_nr, komentaras)
        else:
            CTkMessagebox(title="Error", message="Patikrinkite įvestus duomenis!", icon="cancel")

def prideti_medziaga(medziaga, naudotas, numeris, zmogus, komentaras):
    try:
        medziaga1 = []
        medziaga2 = []
        medziaga3 = []
        medziaga4 = []
        serijos_nr = []
        info = []
        data = []
        asmuo = []
        today = datetime.datetime.today()
        kiekis = naudotas
        kiekis2 = int(kiekis)
        if medziaga == "Polietileno plėvelė":
            medziaga1.append(kiekis2)
        elif medziaga == "Hidroizoliacinė plėvelė":
            medziaga2.append(kiekis2)
        elif medziaga == "Sausas tinko mišinys":
            medziaga3.append(kiekis2)
        elif medziaga == "Cementinis mišinys":
            medziaga4.append(kiekis2)
        serijos_nr.append(numeris)
        data.append(today.strftime("%Y-%m-%d %H:%M:%S"))
        asmuo.append(zmogus)
        info.append(komentaras)

        df1 = pd.DataFrame(medziaga1, columns=['Polietileno plėvelė 0.85x50m'])
        df2 = pd.DataFrame(medziaga2, columns=['Hidroizoliacinė plėvelė 0.65x45m'])
        df3 = pd.DataFrame(medziaga3, columns=['Sausas tinko mišinys 2.5kg'])
        df4 = pd.DataFrame(medziaga4, columns=['Cementinis mišinys 4kg'])
        df5 = pd.DataFrame(asmuo, columns=['Atsakingas asmuo'])
        df6 = pd.DataFrame(data, columns=['Data'])
        df7 = pd.DataFrame(serijos_nr, columns=['Partijos numeris'])
        df8 = pd.DataFrame(info, columns=['Papildoma informacija'])
        df = pd.concat([df1, df2, df3, df4, df7, df5, df6, df8], axis=1)

        if os.path.isfile('Sandėlio istorija.xlsx'):
            with pd.ExcelWriter("Sandėlio istorija.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
                ws = pd.read_excel('Sandėlio istorija.xlsx')
                eilute = len(ws) + 1
                df.to_excel(writer, index=False, header=False, startrow=eilute)
            CTkMessagebox(title="Info", message="Sandėlys sėkmingai papildytas")
        else:
            df.to_excel("Sandėlio istorija.xlsx", index=False, header=True)
            CTkMessagebox(title="Info", message="Sandėlys sėkmingai papildytas")

        if os.path.isfile('_internal\master_files\Sandėlio istorija master.xlsx'):
            with pd.ExcelWriter("_internal\master_files\Sandėlio istorija master.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
                ws = pd.read_excel('_internal\master_files\Sandėlio istorija master.xlsx')
                eilute = len(ws) + 1
                df.to_excel(writer, index=False, header=False, startrow=eilute)
        else:
            df.to_excel("_internal\master_files\Sandėlio istorija master.xlsx", index=False, header=True)

    except ValueError:
        CTkMessagebox(title="Error", message="Sandėlio pildymo kiekis turi būti sveikasis skaičius!", icon="cancel")

def atimti_medziaga_is_sandelio(medziaga, naudotas, numeris, zmogus, serijos_nr_spaudejo, komentaras):
    medziaga1 = []
    medziaga2 = []
    medziaga3 = []
    medziaga4 = []
    serijos_nr = []
    serijos_nr_sp = []
    data = []
    asmuo = []
    info = []
    info2 = []
    today = datetime.datetime.today()
    kiekis = naudotas
    kiekis2 = -abs(int(kiekis))
    if medziaga == "Polietileno plėvelė":
        medziaga1.append(kiekis2)
    elif medziaga == "Hidroizoliacinė plėvelė":
        medziaga2.append(kiekis2)
    elif medziaga == "Sausas tinko mišinys":
        medziaga3.append(kiekis2)
    elif medziaga == "Cementinis mišinys":
        medziaga4.append(kiekis2)
    serijos_nr.append(numeris)
    data.append(today.strftime("%Y-%m-%d %H:%M:%S"))
    asmuo.append(zmogus)
    serijos_nr_sp.append(serijos_nr_spaudejo)
    info.append(komentaras)

    df1 = pd.DataFrame(medziaga1, columns=['Polietileno plėvelė 0.85x50m']).round(2)
    df2 = pd.DataFrame(medziaga2, columns=['Hidroizoliacinė plėvelė 0.65x45m']).round(2)
    df3 = pd.DataFrame(medziaga3, columns=['Sausas tinko mišinys 2.5kg']).round(2)
    df4 = pd.DataFrame(medziaga4, columns=['Cementinis mišinys 4kg']).round(2)
    df5 = pd.DataFrame(asmuo, columns=['Atsakingas asmuo'])
    df6 = pd.DataFrame(data, columns=['Data'])
    df7 = pd.DataFrame(serijos_nr, columns=['Partijos numeris'])
    df8 = pd.DataFrame(info2, columns=['Papildoma informacija'])

    dfs1 = pd.DataFrame(medziaga1, columns=['Polietileno plėvelė, m']).apply(lambda x: x * -30).round(2)
    dfs2 = pd.DataFrame(medziaga2, columns=['Hidroizoliacinė plėvelė, m']).apply(lambda x: x * -45).round(2)
    dfs3 = pd.DataFrame(medziaga3, columns=['Sausas tinko mišinys, kg']).apply(lambda x: x * -2.5).round(2)
    dfs4 = pd.DataFrame(medziaga4, columns=['Cementinis mišinys, kg']).apply(lambda x: x * -4).round(2)
    dfs5 = pd.DataFrame(asmuo, columns=['Atsakingas asmuo'])
    dfs6 = pd.DataFrame(data, columns=['Data'])
    dfs7 = pd.DataFrame(serijos_nr, columns=['Partijos numeris'])
    dfs8 = pd.DataFrame(serijos_nr_sp, columns=['Medžiagos ID'])
    dfs9 = pd.DataFrame(info, columns=['Papildoma informacija'])

    df = pd.concat([df1, df2, df3, df4, df7, df5, df6, df8], axis=1)

    if os.path.isfile('Sandėlio istorija.xlsx'):
        with pd.ExcelWriter("Sandėlio istorija.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            ws = pd.read_excel('Sandėlio istorija.xlsx')
            eilute = len(ws) + 1
            df.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=eilute)
    else:
        df.to_excel("Sandėlio istorija.xlsx", index=False, header=True)

    if os.path.isfile('_internal\master_files\Sandėlio istorija master.xlsx'):
        with pd.ExcelWriter("_internal\master_files\Sandėlio istorija master.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            ws = pd.read_excel('_internal\master_files\Sandėlio istorija master.xlsx')
            eilute = len(ws) + 1
            df.to_excel(writer, index=False, header=False, startrow=eilute)
    else:
        df.to_excel("_internal\master_files\Sandėlio istorija master.xlsx", index=False, header=True)

    dfs = pd.concat([dfs1, dfs2, dfs3, dfs4, dfs7, dfs8, dfs5, dfs6, dfs9], axis=1)

    if os.path.isfile('Gamybos medžiagų istorija.xlsx'):
        with pd.ExcelWriter("Gamybos medžiagų istorija.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            ws = pd.read_excel('Gamybos medžiagų istorija.xlsx')
            eilute = len(ws) + 1
            dfs.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=eilute)
        CTkMessagebox(title="Info", message="Perkeliamos medžiagos sėkmingai užregistruotos")
    else:
        dfs.to_excel("Gamybos medžiagų istorija.xlsx", index=False, header=True)
        CTkMessagebox(title="Info", message="Perkeliamos medžiagos sėkmingai užregistruotos")

    if os.path.isfile('_internal\master_files\Gamybos medžiagų istorija master.xlsx'):
        with pd.ExcelWriter("_internal\master_files\Gamybos medžiagų istorija master.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            ws = pd.read_excel('_internal\master_files\Gamybos medžiagų istorija master.xlsx')
            eilute = len(ws) + 1
            dfs.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=eilute)
    else:
        dfs.to_excel("_internal\master_files\Gamybos medžiagų istorija master.xlsx", index=False, header=True)

def uzsakyme_panaudota_medziaga(medziaga, naudotas, zmogus, serijos_nr_spaudejo, aprasymas, komentaras):
    medziaga1 = []
    medziaga2 = []
    medziaga3 = []
    medziaga4 = []
    serijos_nr_sp = []
    data = []
    asmuo = []
    serijos_nr = []
    darbo_aprasymas = []
    info = []
    today = datetime.datetime.today()
    kiekis = naudotas
    kiekis2 = -abs(float(kiekis))
    if medziaga == "Polietileno plėvelė":
        medziaga1.append(kiekis2)
    elif medziaga == "Hidroizoliacinė plėvelė":
        medziaga2.append(kiekis2)
    elif medziaga == "Sausas tinko mišinys":
        medziaga3.append(kiekis2)
    elif medziaga == "Cementinis mišinys":
        medziaga4.append(kiekis2)
    data.append(today.strftime("%Y-%m-%d %H:%M:%S"))
    asmuo.append(zmogus)
    serijos_nr_sp.append(serijos_nr_spaudejo)
    darbo_aprasymas.append(aprasymas)
    info.append(komentaras)

    df1 = pd.DataFrame(medziaga1, columns=['Polietileno plėvelė, m']).round(2)
    df2 = pd.DataFrame(medziaga2, columns=['Hidroizoliacinė plėvelė, m']).round(2)
    df3 = pd.DataFrame(medziaga3, columns=['Sausas tinko mišinys, kg']).round(2)
    df4 = pd.DataFrame(medziaga4, columns=['Cementinis mišinys, kg']).round(2)
    df5 = pd.DataFrame(asmuo, columns=['Atsakingas asmuo'])
    df6 = pd.DataFrame(data, columns=['Data'])
    df7 = pd.DataFrame(serijos_nr, columns=['Partijos numeris'])
    df8 = pd.DataFrame(serijos_nr_sp, columns=['Medžiagos ID'])
    df9 = pd.DataFrame(info, columns=['Papildoma informacija'])

    dfs1 = pd.DataFrame(medziaga1, columns=['Polietileno plėvelė, m']).round(2)
    dfs2 = pd.DataFrame(medziaga2, columns=['Hidroizoliacinė plėvelė, m']).round(2)
    dfs3 = pd.DataFrame(medziaga3, columns=['Sausas tinko mišinys, kg']).round(2)
    dfs4 = pd.DataFrame(medziaga4, columns=['Cementinis mišinys, kg']).round(2)
    dfs5 = pd.DataFrame(asmuo, columns=['Atsakingas asmuo'])
    dfs6 = pd.DataFrame(data, columns=['Data'])
    dfs8 = pd.DataFrame(serijos_nr_sp, columns=['Medžiagos ID'])
    dfs9 = pd.DataFrame(darbo_aprasymas, columns=['Užsakymo numeris'])
    dfs10 = pd.DataFrame(info, columns=['Papildoma informacija'])

    df = pd.concat([df1, df2, df3, df4, df7, df8, df5, df6, df9], axis=1)

    if os.path.isfile('Gamybos medžiagų istorija.xlsx'):
        with pd.ExcelWriter("Gamybos medžiagų istorija.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            ws = pd.read_excel('Gamybos medžiagų istorija.xlsx')
            eilute = len(ws) + 1
            df.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=eilute)
    else:
        df.to_excel("Gamybos medžiagų istorija.xlsx", index=False, header=True)

    if os.path.isfile('_internal\master_files\Gamybos medžiagų istorija master.xlsx'):
        with pd.ExcelWriter("_internal\master_files\Gamybos medžiagų istorija master.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            ws = pd.read_excel('_internal\master_files\Gamybos medžiagų istorija master.xlsx')
            eilute = len(ws) + 1
            df.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=eilute)
    else:
        df.to_excel("_internal\master_files\Gamybos medžiagų istorija master.xlsx", index=False, header=True)

    dfs = pd.concat([dfs9, dfs1, dfs2, dfs3, dfs4, dfs8, dfs5, dfs6, dfs10], axis=1)

    if os.path.isfile('Gamybos darbų registras.xlsx'):
        with pd.ExcelWriter("Gamybos darbų registras.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            ws = pd.read_excel('Gamybos darbų registras.xlsx')
            eilute = len(ws) + 1
            dfs.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=eilute)
        CTkMessagebox(title="Info", message=f"Užsakymas '{aprasymas}' sekmingai uzregistruotos")
    else:
        dfs.to_excel("Gamybos darbų registras.xlsx", index=False, header=True)
        CTkMessagebox(title="Info", message=f"Užsakymas '{aprasymas}' sekmingai uzregistruotos")

if __name__ == "__main__":
    app = App()
    app.mainloop()