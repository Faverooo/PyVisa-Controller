import queue
from threading import Thread
import threading
from GUI_graph import *
from exitwindow import *
import exitwindow
import sys
import time
import os
import subprocess
import pyvisa as pv
import math
import datetime
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
#CREAZIONE DIR SALVATAGGIO DATI
try:
    os.makedirs("Salvataggi Test")
except:
    pass
percorso = os.getcwd()
os.chdir(f"{percorso}/Salvataggi Test")
percorso = os.getcwd()

#PER FASI tab_widget
fasi = {
    "Tempo": [],
    "Potenza": []
}
#PER MEMORIZZAZIONE VALORI
tabella = {
    "Tempo": [],
    "Corrente": [],
    "Tensione": [],
    "Potenza": [],
    "Resistenza": []
}
attesa= queue.Queue()
avvio = bool(0)
stop = bool(0)
collegamento = 0
vm = 0
voltage = 0
Res = 5 #POI VIENE MISURATA PRIMA DELL'INIZIO CICLI
my_thread = 0
cicli = 0

#RICERCA STRUMENTO E VALIDAZIONE ID
def collegamento_strumento():
    global vm, collegamento
    rm = pv.ResourceManager()
    collegamento=0
    identificazione=""
    dispositivi=0
    list_collegati=rm.list_opened_resources()
    if len(list_collegati)!=0:
        vm.close()
        time.sleep(0.1)
    if collegamento==0:
        try:
            dispositivi=rm.list_resources()
        except:pass
        for i in range(len(dispositivi)):
            vm=rm.open_resource(dispositivi[i])
            try:
                identificazione=vm.query("*IDN?")
                print(identificazione)
            except:pass
            if identificazione[:6]=="HHY230":
                collegamento=1
                print("fineeee")
                break
            else:
                vm.close()
                time.sleep(0.1)
    print("fine")
    if collegamento==0:
        raise Exception("nessun alimentatore trovato")

def errore_strumento():
    errore = QtWidgets.QErrorMessage()
    errore.showMessage("Collega lo strumento")
    errore.exec()

def errore_operazione():
    errore = QtWidgets.QErrorMessage()
    errore.showMessage("Errore Durante l'operazione")
    errore.exec()


class Edit(Ui_TESTALIM):
    def __init__(self, Widget):
        self.setupUi(Widget)
        self.tableWidget.verticalHeader().setVisible(False)
        self.horizontalLayout_20=QtWidgets.QHBoxLayout(self.frame)#AGGIUNGERE LAYOUT AL POSTO DI FRAME VUOTO
        self.horizontalLayout_20.setObjectName("horizontalLayout_20")
        self.figure=plt.figure()
        self.canvas=FigureCanvas(self.figure)
        self.horizontalLayout_20.addWidget(self.canvas)#CANVAS DENTRO AL LAYOUT
        self.check()


    #CHECK TASTI E LINK FUNZIONI
    def check(self):
        self.start_on.clicked.connect(self.lettura_cicli)
        self.aggiungi.clicked.connect(self.tabella)
        self.elimina.clicked.connect(self.cancriga)
        self.ovp_on.clicked.connect(self.ovp_acceso)
        self.ovp_off.clicked.connect(self.ovp_spento)
        self.ocp_on.clicked.connect(self.ocp_acceso)
        self.ocp_off.clicked.connect(self.ocp_spento)
        self.start_off.clicked.connect(self.spegni)
        self.salvataggi.clicked.connect(self.apricartella)
        self.collega_strumento.clicked.connect(self.collegamento)
        

    def collegamento(self):
        if avvio==0:
            try:
                collegamento_strumento()
            except:
                errore_strumento()

    #APERTURA EXITWINDOW 
    def spegni(self):
        self.Dialog = QtWidgets.QDialog()
        self.ui = Ui_Second_screen()
        self.ui.setupUi(self.Dialog)
        self.Dialog.setFixedSize(400, 200)
        self.Dialog.exec()

    #CONTROLLO OCP
    def ocp_acceso(self):
        if collegamento == 1:
            try:
                value = self.ocp.text()
                value = int(value)
                vm.write(f'CURR:PROT {value}')
                vm.write('CURR:PROT:STAT ON')
                time.sleep(0.05)
                self.led_ocp.setStyleSheet("background-color: rgb(0, 200, 0);\n"
                                           "border-radius:10px;")
            except:
                errore_operazione()
        else:
            errore_strumento()

    def ocp_spento(self):
        if collegamento == 1:
            try:
                vm.write('CURR:PROT:STAT OFF')
                time.sleep(0.05)
                self.led_ocp.setStyleSheet("background-color: rgb(180, 0, 0);\n"
                                           "border-radius:10px;")
                self.ocp.clear()
            except:
                errore_operazione()
        else:
            errore_strumento()

    #CONTROLLO OVP
    def ovp_acceso(self):
        if collegamento == 1:
            try:
                value = self.ovp_line.text()
                value = int(value)
                vm.write(f'VOLT:PROT {value}')
                vm.write('VOLT:PROT:STAT ON')
                time.sleep(0.05)
                self.led_ovp.setStyleSheet("background-color: rgb(0, 200, 0);\n"
                                           "border-radius:10px;")
            except:
                errore_operazione()
        else:
            errore_strumento()

    def ovp_spento(self):
        if collegamento == 1:
            try:
                vm.write('VOLT:PROT:STAT OFF')
                time.sleep(0.05)
                self.led_ovp.setStyleSheet("background-color: rgb(180, 0, 0);\n"
                                           "border-radius:10px;")
                self.ovp_line.clear()
            except:
                errore_operazione()
        else:
            errore_strumento()

    #TABULAZIONE DATI
    def tabella(self):
        if avvio == 0:
            try:
                tempo = int(self.tempo.text())
                potenza = int(self.potenza.text())
                numrighe = self.tableWidget.rowCount()
                self.tableWidget.insertRow(numrighe)
                self.tableWidget.setItem(numrighe, 0, QtWidgets.QTableWidgetItem(str(tempo)))
                self.tableWidget.setItem(numrighe, 1, QtWidgets.QTableWidgetItem(str(potenza)))
                global fasi
                fasi["Tempo"].append(tempo)
                fasi["Potenza"].append(potenza)
                self.tempo.clear()
                self.potenza.clear()
            except:
                errore = QtWidgets.QErrorMessage()
                errore.showMessage("Inserisci un numero, non un testo")
                errore.exec()
        else:
            errore = QtWidgets.QErrorMessage()
            errore.showMessage("Programma avviato, spegnere il programma per modificare i paramentri")
            errore.exec()

    #ELIMINAZIONE DATI IN TABELLA E NEL DICT
    def cancriga(self):
        if avvio == 0:
            numrighe = self.tableWidget.rowCount()
            self.tableWidget.removeRow(numrighe - 1)
            try:
                fasi["Tempo"].pop(-1)
                fasi["Potenza"].pop(-1)
            except:
                pass
        else:
            errore = QtWidgets.QErrorMessage()
            errore.showMessage("Programma avviato, spegnere il programma per modificare i paramentri")
            errore.exec()

    #APERTURA CARTELLA MAC E WINDOWS
    def apricartella(self):
        try:
            subprocess.call(["open", percorso])  # MACOS
        except:
            subprocess.Popen(f'explorer "{percorso}"')  # WINDOWS

    #LETTURA VALORI E SCRITTURA
    def lettura_cicli(self):
        if collegamento == 1:
            global cicli
            try:
                cicli = int(self.n_cicli.text())
            except:
                
                return None
            self.threadprog()
        else:
            errore_strumento()

    #CREAZIONE THREAD PARALLELO AL MAIN
    def threadprog(self):
        #DIVISIONE THREAD => IL PROGRAMMA GIRA IN UN SECONDO PROCESSO
        if avvio == 0:
            global my_thread
            my_thread = Thread(target=self.partenza_cicli)
            my_thread.start()
        else:
            errore = QtWidgets.QErrorMessage()
            errore.showMessage("Programma già avviato")
            errore.exec()

    #ESECUZIONE CICLI
    def partenza_cicli(self):
        global cicli, avvio, Res
        avvio = 1
        TOT_tempo = 0
        exitwindow.stop=0
        self.figure.clear()
        self.canvas.draw()
        for i in range(len(fasi["Tempo"])):
            TOT_tempo += fasi["Tempo"][i]
        TOT_tempo = (TOT_tempo * cicli) / 60
        x = datetime.datetime.now()
        filenm = str(x.strftime("%Y%m%d-%H%M%S_log"))#NOME FILE DATA E ORA
        file_log = open(f"{filenm}.txt", "w")
        file_log.write("Tempo;Tensione;Corrente;Potenza;Resistenza" + "\n")
        file_log.close()
        self.led_start.setStyleSheet("background-color: rgb(0, 200, 0);\n"
                                     "border-radius:10px;")
        #MISURAZIONE INIZIALE RESISTENZA
        try:
            vm.write('source:voltage:level 5')
            time.sleep(0.5)
            vm.write('source:current:level 2')
            time.sleep(0.5)
            vm.write('OUTP ON')
            time.sleep(0.5)
            vm.write('MEAS:CURR?')
            cor = float(vm.read())
            Res = 5 / cor
        except:
            errore_strumento()
            return
        self.ciclo_target.display(str(cicli))
        errori = 0
        first_line = 0
        default_time = time.time()
        #INIZIO TEST
        for i in range(cicli):
            self.ciclo_npw.display(str(i + 1))
            for el in range(len(fasi["Potenza"])):
                ciclo_time = time.time()
                target = 0
                while True:
                    trascorso = time.time() - ciclo_time
                    if trascorso > target:
                        global my_thread 
                        try:                        #SETTING TENSIONE
                            self.lettura(default_time, fasi["Potenza"][el], filenm, first_line)
                            voltage = float(math.sqrt((fasi["Potenza"][el]) * Res))
                            vm.write(f'source:voltage:level {voltage}')
                            time.sleep(0.1)
                            if voltage != 0:        #SETTING CORRENTE
                                current = float((voltage / Res) + 2)
                                vm.write(f'source:current:level {current}')
                            time.sleep(0.1)
                        except:
                            errori += 1
                            elapsed = time.time() - default_time
                            print(f"ERRORE A TEMPO: {elapsed}")
                        tempo_passato = (time.time() - default_time) / 60
                        tempo_mancante = TOT_tempo - tempo_passato
                        self.lcd_trascorso.display(f"{tempo_passato:.0f}")
                        self.lcd_mancante.display(f"{tempo_mancante:.0f}")
                        target += 1
                        first_line = 1
                        if errori > 20:
                            self.poweroff(filenm)#FUNZIONE SPEGNIMENTO E SALVATAGGIO
                            print("FINE DOVUTA AD ERRORI")
                            return None
                        if target >= (fasi["Tempo"][el]):
                            break
                    if exitwindow.stop == 1:
                        self.poweroff(filenm)
                        return
                    debug_1 = threading.main_thread().is_alive()#CONTROLLO CHIUSURA THREAD PRINCIPALE
                    if (debug_1 == False):
                        self.poweroff(filenm)
                        return
                time.sleep(0.35)
        self.poweroff(filenm)

    #LETTURA DATI E SCRITTURA
    def lettura(self, default_time, pot, filenm, first_line):
        global tabella, Res
        vm.write('MEAS:VOLT?')
        voltCH1 = float(vm.read())
        time.sleep(0.1)
        vm.write('MEAS:CURR?')
        AmpCH1 = float(vm.read())
        time.sleep(0.1)
        Potenza = voltCH1 * AmpCH1
        Res = 0
        try:
            if pot != 0:
                Res = voltCH1 / AmpCH1
        except:
            pass
        self.lcd_ohm.display(f"{Res:.2f}")
        self.lcd_A.display(f"{AmpCH1:.2f}")
        self.lcd_W.display(f"{Potenza:.2f}")
        self.lcd_v.display(f"{voltCH1:.2f}")
        if pot != 0:
            Res = voltCH1 / AmpCH1
        if first_line == 1:
            timeelapsed = time.time() - default_time
            tabella["Resistenza"].append(Res)
            tabella["Potenza"].append(Potenza)
            tabella["Corrente"].append(AmpCH1)
            tabella["Tensione"].append(voltCH1)
            tabella["Tempo"].append(timeelapsed)
            if len(tabella["Resistenza"])>900:
                tabella["Resistenza"].pop(0)
                tabella["Potenza"].pop(0)
                tabella["Corrente"].pop(0)
                tabella["Tensione"].pop(0)
                tabella["Tempo"].pop(0)
            print(
                f"Time={timeelapsed:.2f} Volt={voltCH1:.2f}V  Ampere= {AmpCH1:.2f}A Potenza={Potenza:.2f}W Resistenza={Res:.2f}omh" + "\n")
            file_log = open(f"{filenm}.txt", "a")#A=APPEND
            file_log.write(f"{timeelapsed:.2f};{voltCH1:.2f};{AmpCH1:.2f};{Potenza:.2f};{Res:.2f}" + "\n")
            file_log.close()
            self.mostra_grafico()

    def mostra_grafico(self):
        
        self.figure.clear()
        ax0 = plt.subplot()
        ax0.plot(tabella["Tempo"],
                tabella["Resistenza"],
                color="blue",
                label='RESISTENZA', 
                linewidth=0.5)
        ax1 = ax0.twinx()
        ax1.plot(tabella["Tempo"], tabella["Potenza"],color="red",label='POTENZA',linewidth=0.5)
        #ax0.set_title('RESISTENZA-POTENZA')
        ax0.tick_params(axis='y', labelsize=8)
        ax1.tick_params(axis='y', labelsize=8)
        ax0.tick_params(axis='x', labelsize=8)
        ax0.locator_params(axis='x', nbins=6)
        ax0.spines['left'].set_color('blue')
        ax0.spines['left'].set_linewidth(3)
        ax0.tick_params(axis='y',colors='blue')
        ax1.spines['right'].set_color('red')
        ax1.spines['right'].set_linewidth(3)
        ax1.tick_params(axis='y',colors='red')
        ax0.legend(framealpha=0.5, loc='upper left',fontsize=8)
        ax1.legend(framealpha=0.5, loc='upper right',fontsize=8)
        self.canvas.draw()
        
    #SPEGNIMENTO, AZZERAMENTO VALORI E AZZERAMENTO VARIABILI, SCRITTURA CSV
    def poweroff(self, filenm):
        #AZZERAMENTO VALORI
        global avvio, connessione,tabella
        avvio = 0
        connessione = 0
        exitwindow.stop = 0
        self.led_start.setStyleSheet("background-color: rgb(180, 0, 0);\n"
                                     "border-radius:10px;")
        self.lcd_mancante.display("0")
        self.lcd_trascorso.display("0")
        self.lcd_ohm.display("0")
        self.lcd_A.display("0")
        self.lcd_W.display("0")
        self.lcd_v.display("0")
        self.ciclo_npw.display("0")
        self.ciclo_target.display("0")
        #PROVA A SPEGNERE 
        try:
            vm.write('OUTP OFF')
            time.sleep(0.1)
            vm.write('source:voltage:level 0')
        except:pass
        #SALVATAGGIO LOG
        try:
            df = pd.read_fwf(f"{filenm}.txt")
            df.to_csv(f"{filenm}.csv", index=False)
            #convert=pd.DataFrame(tabella)
            #convert.to_csv(f"{filenm}.csv",index=False)
        except:
            print("ERRORE CREAZIONE CSV")
        tabella = {
        "Tempo": [],
        "Corrente": [],
        "Tensione": [],
        "Potenza": [],
        "Resistenza": []
        }

app = QtWidgets.QApplication(sys.argv)
Widget = QtWidgets.QWidget()
ui = Edit(Widget)
Widget.show()
Widget.setFixedSize(1110, 600)
sys.exit(app.exec())