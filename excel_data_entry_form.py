"""GUI to help me fill in a big excel form."""
import pandas as pd
import PySimpleGUI as sg

# ---------------------------- FILENAMES ---------------------------------------

FILENAMES = ['1935_Landauer.jpg', '1935_Lehmann_K1.jpg', '1935_Levi.jpg', '1935_Salomon_A_K2.jpg', '1935_Salomon_B.jpg',
             '1936_Dannenberg.jpg', '1936_Labhard.jpg', '1936_Lilienthal.jpg', '1937_Berend.jpg', '1937_Cassy.jpg',
             '1937_Holkar.jpg', '1937_Koch-Sichel.jpg', '1937_Kronheimer.jpg', '1937_Lustig.jpg', '1937_May.jpg',
             '1937_Nacamuli.jpg', '1937_Petry.jpg', '1937_Ruppert.jpg', '1937_Siemens.jpg', '1937_Thiel.jpg',
             '1937_Vikar.jpg', '1938_Baalen.jpg', '1938_Furrer.jpg', '1938_Jellinek.jpg', '1938_Lehmann_K2.jpg',
             '1938_Lipovsky.jpg', '1938_Schwersenzer.jpg', '1938_Vögler.jpg', '1939_Bonheim.jpg', '1939_Frank.jpg',
             '1939_Glaessner.jpg', '1939_Goetjes.jpg', '1939_Gut.jpg', '1939_Gutmann.jpg', '1939_Hassan-Izzet.jpg',
             '1939_Jacob.jpg', '1939_Jacobsohn.jpg', '1939_Kamber.jpg', '1939_Kastor.jpg', '1939_Krolik.jpg',
             '1939_Kroner.jpg', '1939_Levy.jpg', '1939_Lysley.jpg', '1939_Philips.jpg', '1939_Rehfisch.jpg',
             '1939_Römer.jpg', '1939_Schinke.jpg', '1939_Schlessmann.jpg', '1939_Terboven.jpg', '1939_Ullmann_K3.jpg',
             '1939_VonSchwarzenberg.jpg', '1940_Bibus-Jäger_K1.jpg', '1940_Bibus-Jäger_K2.jpg', '1940_Boehnlen.jpg',
             '1940_Bonardelly.jpg', '1940_Bult.jpg', '1940_Burckhardt.jpg', '1940_Burger-Kehl.jpg', '1940_Donnet.jpg',
             '1940_Ferchel.jpg', '1940_Herfeld.jpg', '1940_Kaufmann.jpg', '1940_Willy.jpg', '1941_Baronin.jpg',
             '1941_Berber.jpg', '1941_Bernheim.jpg', '1941_Fiedler.jpg', '1941_Frölicher.jpg', '1941_Madjid.jpg',
             '1941_Pannwitz.jpg', '1941_Sauerbruch.jpg', '1941_Schwarz.jpg', '1942_Ammon.jpg', '1942_Balli.jpg',
             '1942_Jaeger.jpg', '1942_Robinson.jpg', '1942_Vieli.jpg', '1942_VonDaragan_K2.jpg', '1943_Bauer.jpg',
             '1943_Ehrlich.jpg', '1943_Feist.jpg', '1943_Hirschland.jpg', '1943_Jeker.jpg', '1943_Lützelschwab.jpg',
             '1943_Meyer.jpg', '1943_Neumann.jpg', '1943_Schkaff.jpg', '1944_Baumann-Fricker.jpg', '1944_Bollag.jpg',
             '1944_Grieshaber.jpg', '1944_Lorang-Ringier.jpg', '1944_Müller.jpg', '1944_Watson.jpg',
             '1945_Jacottet.jpg', '1945_Schuchhalter.jpg', '1945_Witting.jpg', '1946_Figge_K1.jpg', '1946_Figge_K2.jpg',
             '1946_Kobelt.jpg', '1946_Kuranda.jpg', '1946_Lepage.jpg', '1946_Minger_K1.jpg', '1946_Muheim.jpg',
             '1946_Phethean.jpg', '1946_Schneider-Zweifel.jpg', '1946_Selig.jpg', '1947_Arletti.jpg',
             '1947_Camarasesco.jpg', '1947_Domec.jpg', '1947_Giordanou.jpg', '1947_Heer.jpg', '1947_Pfister.jpg',
             '1947_Rothmann.jpg', '1947_Sabbagh.jpg', '1947_Schneider.jpg', '1947_Sharpe.jpg', '1947_Sigler.jpg',
             '1947_Soloveytchik.jpg', '1948_Ades.jpg', '1948_Almqvist.jpg', '1948_Amsler.jpg', '1948_Asti.jpg',
             '1948_Ausnit.jpg', '1948_Blackwell-Wegmann.jpg', '1948_Braunschweig.jpg', '1948_Bronsten.jpg',
             '1948_Byk_K1.jpg', '1948_Byk_K2.jpg', '1948_Chorley.jpg', '1948_DeRothschild.jpg', '1948_Dubois.jpg',
             '1948_Eisenbach.jpg', '1948_Entwistle.jpg', '1948_Feinstein.jpg', '1948_Floersheimer.jpg',
             '1948_Frohnknecht.jpg', '1948_Glahn.jpg', '1948_Holzer.jpg', '1948_Jellinek.jpg', '1948_Juda.jpg',
             '1948_KaeserVonTobel.jpg', '1948_Katz.jpg', '1948_Khaitan.jpg', '1948_Lievre.jpg', '1948_Melamid.jpg',
             '1948_Minger_K2.jpg', '1948_Pateras.jpg', '1948_Popper.jpg', '1948_Poser.jpg', '1948_Ratkin.jpg',
             '1948_Reisfeld.jpg', '1948_Rothschild.jpg', '1948_Rueff.jpg', '1948_Selver.jpg', '1948_Tuttnauer.jpg',
             '1948_Voellmy.jpg', '1948_VonBernard.jpg', '1948_Wald.jpg', '1948_William.jpg', '1949_Balla.jpg',
             '1949_Bond.jpg', '1949_Brann.jpg', '1949_Carre.jpg', '1949_Getty.jpg', '1949_Giannini.jpg',
             '1949_Grossmann.jpg', '1949_Hirsch.jpg', '1949_Jenny.jpg', '1949_Prager.jpg', '1949_Seizer.jpg',
             '1949_Sinsheimer.jpg', '1949_Steimle.jpg', '1949_Steinberg.jpg', '1949_Sternberg.jpg', '1949_Weil.jpg',
             '1949_Weldon.jpg', '1950_Kohn.jpg', '1950_Loebel.jpg', '1950_Rice.jpg', '1950_Steinmetz.jpg',
             '1950_Trifonoff.jpg', '1950_Wormser.jpg', '1951_Grassia.jpg', '1951_Herz.jpg', '1951_Heymann.jpg',
             '1951_Hilpert.jpg', '1951_Huguenin.jpg', '1951_Joseson.jpg', '1951_Jsrael.jpg', '1951_Kaufmann.jpg',
             '1951_Lagnado.jpg', '1951_Lowenthal.jpg', '1951_Luder.jpg', '1951_Mayer.jpg', '1951_Rosenau.jpg',
             '1951_Samuel.jpg', '1951_Selby.jpg', '1951_Yarht.jpg', '1952_Eisner.jpg', '1952_Eudlitz.jpg',
             '1952_Guirche.jpg', '1952_Knittel.jpg', '1952_Maltzan.jpg', '1952_Masturzo.jpg', '1952_Mises.jpg',
             '1952_Posner.jpg', '1952_Rosenstiel.jpg', '1953_Adler.jpg', '1953_Hengelhaupt.jpg', '1953_Rassini.jpg',
             '1953_Weizman.jpg', '1954_Kiwit.jpg', '1955_Feldmann.jpg', '1956_Engel.jpg', '1957_Duerrenmatt.jpg',
             '1958_Ebner.jpg', '1960_Aschenasy.jpg', '1961_Heuss.jpg', '1963_Guttman.jpg', '1964_Burger.jpg',
             'DATUM_Bracher.jpg', 'DATUM_Muntwyler.jpg', 'DATUM_VonWattenwyl.jpg']

# ------------------------------- GUI ------------------------------------------

sg.theme("DarkTeal9")

EXCEL_FILE = 'gaestekarteien_from_gui.xlsx'
df = pd.read_excel(EXCEL_FILE)


layout = [
    [sg.Text('Füllen Sie bitte die folgenden Felder aus:')],
    [sg.Text('Adressfeld:')],
    [sg.Text('Kartenname', size=(15, 1)), sg.Combo(FILENAMES, key='Kartenname')],
    [sg.Text('Name', size=(15, 1)), sg.InputText(key='Name')],
    [sg.Text('Todesdatum', size=(15, 1)), sg.InputText(key='Todesdatum')],
    [sg.Text('Adresse', size=(15, 1)), sg.InputText(key='Adresse')],
    [sg.Text('Notizen druckschriftl.', size=(15, 1)), sg.InputText(key='NotizenDruck1')],
    [sg.Text('Notizen handschriftl.', size=(15, 1)), sg.InputText(key='NotizenHand1')],
    [sg.Text('Aufenthalts-Feld:')],
    [sg.Text('Aufenthalt', size=(15, 1)), sg.InputText(key='Aufenthalt')],
    [sg.Text('Notizen druckschriftl.', size=(15, 1)), sg.InputText(key='NotizenDruck2')],
    [sg.Text('Notizen handschriftl.', size=(15, 1)), sg.InputText(key='NotizenHand2')],
    [sg.Submit(), sg.Exit()]
]

window = sg.Window('Gästekarten - Data entry form', layout)


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')

window.close()
