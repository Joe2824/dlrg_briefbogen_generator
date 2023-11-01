import os
import sys
import warnings
import datetime
import importlib
import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate, RichText
from PyQt6.QtCore import QSize
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QCheckBox, QPlainTextEdit, QFileDialog, QVBoxLayout, QHBoxLayout, QWidget, QScrollArea, QButtonGroup, QRadioButton, QMessageBox, QToolBar

basedir = os.path.dirname(__file__)

try:
    from ctypes import windll  # Only exists on Windows.
    APPID = 'joe2824.dlrgbriefbogengenerator'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(APPID)
except ImportError:
    pass

if '_PYIBoot_SPLASH' in os.environ and importlib.util.find_spec("pyi_splash"):
    import pyi_splash
    pyi_splash.update_text('Loading...')
    pyi_splash.close()

if  importlib.util.find_spec("win32com"):
    from win32com.client import *
    def get_version_number(file_path):
        information_parser = Dispatch("Scripting.FileSystemObject")
        print(information_parser)
        version = information_parser.GetFileVersion(file_path)
        return version
    VERSION = f'v{get_version_number(sys.argv[0])}'
else:
    VERSION = 'DEV VERSION'


class DLRGBriefbogenGenerator(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(f'DLRG Briefbogen Generator {VERSION}')
        self.setMinimumWidth(500)

        self.folder_path = None
        self.output_folder = None
        self.gen_data_file = None
        self.files = []
        self.allgemein_df = None
        self.vorstand_df = None
        self.jugend_df = None
        self.year = datetime.datetime.now().year

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        toolbar = QToolBar("Toolbar")
        toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(toolbar)

        button_generate_example = QAction("Daten Vorlage exportieren", self)
        button_generate_example.setStatusTip(
            "Exportiert eine Vorlage welche mit eigenen Daten befüllt werden kann")
        button_generate_example.triggered.connect(self.export_gen_data_example)
        toolbar.addAction(button_generate_example)

        button_about = QAction("Über", self)
        button_about.setStatusTip("Zeige Programm infos")
        button_about.triggered.connect(self.show_info)
        toolbar.addAction(button_about)

        folder_layout = QHBoxLayout()
        layout.addLayout(folder_layout)

        folder_label = QLabel('Template Ordner:')
        folder_layout.addWidget(folder_label)

        self.folder_entry = QLineEdit()
        folder_layout.addWidget(self.folder_entry)

        folder_button = QPushButton('Auswählen', clicked=self.select_folder)
        folder_layout.addWidget(folder_button)

        self.file_area = QScrollArea()
        self.file_area.setWidgetResizable(True)
        self.file_area.setMinimumSize(300, 60)
        self.file_area.setMaximumSize(1920, 220)
        self.file_area.hide()

        layout.addWidget(self.file_area)

        file_widget = QWidget()
        self.file_area.setWidget(file_widget)

        self.file_layout = QVBoxLayout()
        file_widget.setLayout(self.file_layout)

        self.gen_data_label = QLabel('Daten Quelle:')
        self.gen_data_file_entry = QLineEdit()
        self.gen_data_file_button = QPushButton('Auswählen')
        self.gen_data_file_button.clicked.connect(self.select_gen_data_file)

        gen_data_layout = QHBoxLayout()
        gen_data_layout.addWidget(self.gen_data_label)
        gen_data_layout.addWidget(self.gen_data_file_entry)
        gen_data_layout.addWidget(self.gen_data_file_button)

        output_widget = QWidget(central_widget)
        output_widget.setLayout(gen_data_layout)
        layout.addWidget(output_widget)

        self.group_button_label = QLabel(
            'Welche Daten sollen verwendet werden?')
        self.vorstand_radio = QRadioButton('Vorstand')
        self.vorstand_radio.setChecked(True)
        self.jugend_radio = QRadioButton('Jugend')
        self.group_button = QButtonGroup()
        self.group_button.addButton(self.vorstand_radio)
        self.group_button.addButton(self.jugend_radio)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.group_button_label)
        button_layout.addWidget(self.vorstand_radio)
        button_layout.addWidget(self.jugend_radio)

        self.button_widget = QWidget(central_widget)
        self.button_widget.setLayout(button_layout)
        layout.addWidget(self.button_widget)
        self.button_widget.hide()

        self.output_folder_label = QLabel('Ausgabe Pfad:')
        self.output_folder_entry = QLineEdit()
        self.output_folder_button = QPushButton('Auswählen')
        self.output_folder_button.clicked.connect(self.select_output_folder)

        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_folder_label)
        output_layout.addWidget(self.output_folder_entry)
        output_layout.addWidget(self.output_folder_button)

        output_widget = QWidget(central_widget)
        output_widget.setLayout(output_layout)
        layout.addWidget(output_widget)

        selected_files_button = QPushButton(
            'Generiere Briefbogen', clicked=self.generate_briefbogen)
        layout.addWidget(selected_files_button)

        self.selected_files_text = QPlainTextEdit()
        self.selected_files_text.setMinimumSize(300, 60)
        self.selected_files_text.setMaximumSize(1920, 120)
        layout.addWidget(self.selected_files_text)

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self, 'Wähle den Template Ordner')
        self.folder_path = folder_path
        if folder_path:
            self.folder_entry.setText(folder_path)
            self.file_area.show()
            self.show_files()

    def select_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self, 'Wähle den Ausgabe Ordner',)
        self.output_folder = folder_path
        if folder_path:
            self.output_folder_entry.setText(folder_path)

    def select_gen_data_file(self):
        file_path = QFileDialog.getOpenFileName(
            self, 'Wähle eine Excel Datei', '', 'Excel files (*.xlsx *.xls);;All Files (*)')
        self.gen_data_file = file_path[0]
        if self.gen_data_file:
            self.button_widget.show()
            self.gen_data_file_entry.setText(self.gen_data_file)

    def export_gen_data_example(self):
        folder_path = QFileDialog.getSaveFileName(
            self, 'Exportieren F:xile', 'Briefbogen_Daten.xlsx', 'Excel files (*.xlsx)')
        if folder_path[0]:
            pkl_allgemein_df = pd.read_pickle(os.path.join(
                basedir, 'pkl', 'general.pkl'), compression='xz')
            pkl_vorstand_df = pd.read_pickle(os.path.join(
                basedir, 'pkl', 'vorstand.pkl'), compression='xz')
            pkl_jugend_df = pd.read_pickle(os.path.join(
                basedir, 'pkl', 'jugend.pkl'), compression='xz')

            with pd.ExcelWriter(folder_path[0]) as writer:
                pkl_allgemein_df.to_excel(
                    writer, sheet_name='Allgemeine Daten', index=False)
                pkl_vorstand_df.to_excel(
                    writer, sheet_name='Vorstand', index=False)
                pkl_jugend_df.to_excel(
                    writer, sheet_name='Jugend', index=False)

            msg_box = QMessageBox()
            msg_box.setWindowTitle('Export erfolgreich')
            msg_box.setText(
                f'Die Beispiel Datei wurde an den ausgewählten Pfad exportiert\n{folder_path[0]}')
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.exec()

    def show_files(self):
        for _, file in enumerate(self.files):
            file[1].setParent(None)
        self.files = []

        files = os.listdir(self.folder_path)
        for file in files:
            if Path(file).suffix == '.docx':
                checkbox = QRadioButton(file)
                self.files.append((file, checkbox))
                self.file_layout.addWidget(checkbox)

    def generate_briefbogen(self):
        if not self.folder_path:
            self.selected_files_text.insertPlainText(
                'Kein Template Odner gewählt!\n')
            return
        if not self.output_folder:
            self.selected_files_text.insertPlainText(
                'Keine Ausgabe Ordner ausgewählt!\n')
            return
        if not self.gen_data_file:
            self.selected_files_text.insertPlainText(
                'Keine Daten ausgewählt!\n')
            return
        if sum([1 for file in self.files if file[1].isChecked()]) == 0:
            self.selected_files_text.insertPlainText(
                'Kein Template ausgewählt!\n')
            return

        self.create_dirs()

        # unterdrückt openpyxl warnung
        warnings.filterwarnings(
            'ignore', category=UserWarning, module='openpyxl')

        df_dict = pd.read_excel(
            self.gen_data_file, sheet_name=None, engine='openpyxl')
        try:
            self.allgemein_df = df_dict.get('Allgemeine Daten').set_index(
                'Variablennamen (nicht ändern)').transpose()

            if self.jugend_radio.isChecked():
                p_df = df_dict.get('Jugend')
            else:
                p_df = df_dict.get('Vorstand')
        except Exception as e:
            self.selected_files_text.insertPlainText(f'Daten Template nicht korrekt.\nDu kannst ein Beispiel in der Toolbar generieren.\n Error:{e}')
            return

        selected_files = []
        for file in self.files:
            if file[1].isChecked():
                file_path = os.path.join(self.folder_path, file[0])
                selected_files.append(file_path)
                self.generate_template(p_df, file_path, file[0])

        msg_box = QMessageBox()
        msg_box.setWindowTitle('Briefbogen erstellt')
        msg_box.setText('Es wurden alle Briefbogen erstellt.')
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.exec()

    def generate_template(self, p_df, template, template_name):
        '''
        Funktion erstellt aus dem gegebenen Template 
        einen fertigen Briefbogen. Dieser wird in 
        dem passenden Ordner gespeichert
        '''
        self.selected_files_text.clear()
        tpl = DocxTemplate(template)

        name = 'Vorstand'
        font = 'DLRG Univers 55 Roman'
        color = '#808080'
        size = 12
        if self.jugend_radio.isChecked():
            name = 'Jugend'
            font = 'Mulish'
            color = '#000000'
            size = 14

        for i, row in p_df.iterrows():
            o_email = self.allgemein_df['o_email'][0]
            o_webseite = self.allgemein_df['o_webseite'][0]
            jo_email = self.allgemein_df['jo_email'][0]
            jo_webseite = self.allgemein_df['jo_webseite'][0]
            p_vorname = row['p_vorname']
            p_nachname = row['p_nachname']
            p_email = row['p_email']
            context = {'organisation': self.allgemein_df['organisation'][0],
                       'o_strasse': self.allgemein_df['o_strasse'][0],
                       'o_plz_ort': self.allgemein_df['o_plz_ort'][0],
                       'o_telefon': self.allgemein_df['o_telefon'][0],
                       'o_fax': self.allgemein_df['o_fax'][0],
                       'o_email': RichText(o_email, url_id=tpl.build_url_id(f'mailto:{o_email}'), font=font, size=size, color=color),
                       'o_webseite': RichText(o_webseite, url_id=tpl.build_url_id(o_webseite), font=font, size=size, color=color),
                       'o_amtsgericht_ort_nummer': self.allgemein_df['o_amtsgericht_ort_nummer'][0],
                       'o_vorsitz': self.allgemein_df['o_vorsitz'][0],
                       'o_stv_vorsitz': self.allgemein_df['o_stv_vorsitz'][0],
                       'o_schatzmeister': self.allgemein_df['o_schatzmeister'][0],
                       'o_bank_1_name': self.allgemein_df['o_bank_1_name'][0],
                       'o_bank_1_iban': self.allgemein_df['o_bank_1_iban'][0],
                       'o_bank_1_bic': self.allgemein_df['o_bank_1_bic'][0],
                       'o_bank_2_name': self.allgemein_df['o_bank_2_name'][0],
                       'o_bank_2_iban': self.allgemein_df['o_bank_2_iban'][0],
                       'o_bank_2_bic': self.allgemein_df['o_bank_2_bic'][0],
                       'o_ust_o_str': self.allgemein_df['o_ust_o_str'][0],
                       'o_var1': self.allgemein_df['o_var1'][0],
                       'o_var2': self.allgemein_df['o_var2'][0],
                       'o_var3': self.allgemein_df['o_var3'][0],
                       'o_var4': self.allgemein_df['o_var4'][0],
                       'o_var5': self.allgemein_df['o_var5'][0],
                       'j_organisation': self.allgemein_df['j_organisation'][0],
                       'jo_strasse': self.allgemein_df['jo_strasse'][0],
                       'jo_plz_ort': self.allgemein_df['jo_plz_ort'][0],
                       'jo_telefon': self.allgemein_df['jo_telefon'][0],
                       'jo_fax': self.allgemein_df['jo_fax'][0],
                       'jo_email': RichText(jo_email, url_id=tpl.build_url_id(f'mailto:{o_email}'), font=font, size=size, color=color),
                       'jo_webseite': RichText(jo_webseite, url_id=tpl.build_url_id(o_webseite), font=font, size=size, color=color),
                       'jo_amtsgericht_ort_nummer': self.allgemein_df['jo_amtsgericht_ort_nummer'][0],
                       'jo_vorsitz': self.allgemein_df['jo_vorsitz'][0],
                       'jo_stv_vorsitz': self.allgemein_df['jo_stv_vorsitz'][0],
                       'jo_schatzmeister': self.allgemein_df['jo_schatzmeister'][0],
                       'jo_bank_1_name': self.allgemein_df['o_bank_1_name'][0],
                       'jo_bank_1_iban': self.allgemein_df['o_bank_1_iban'][0],
                       'jo_bank_1_bic': self.allgemein_df['o_bank_1_bic'][0],
                       'jo_bank_2_name': self.allgemein_df['o_bank_2_name'][0],
                       'jo_bank_2_iban': self.allgemein_df['o_bank_2_iban'][0],
                       'jo_bank_2_bic': self.allgemein_df['o_bank_2_bic'][0],
                       'jo_ust_o_str': self.allgemein_df['jo_ust_o_str'][0],
                       'jo_kreisjugendring': self.allgemein_df['jo_kreisjugendring'][0],
                       'jo_var1': self.allgemein_df['jo_var1'][0],
                       'jo_var2': self.allgemein_df['jo_var2'][0],
                       'jo_var3': self.allgemein_df['jo_var3'][0],
                       'jo_var4': self.allgemein_df['jo_var4'][0],
                       'jo_var5': self.allgemein_df['jo_var5'][0],
                       'p_vorname': p_vorname,
                       'p_nachname': p_nachname,
                       'p_funktion': row['p_funktion'],
                       'p_email': RichText(p_email, url_id=tpl.build_url_id(f'mailto:{p_email}'), font=font, size=size, color=color),
                       }
            tpl.render(context, autoescape=True)

            # Ändere content type zu dotx
            doc_part = tpl.part
            doc_part._content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'

            # Dynamischer name für alleBriefbogen
            document_path = os.path.join(self.output_folder, name, str(
                self.year), f'Briefbogen_{p_vorname}_{p_nachname}_{self.year}_{template_name[:-5]}.dotx')
            tpl.save(document_path)

            self.selected_files_text.insertPlainText(
                f'{str(i).zfill(2)}  Generiere {name} Briefbogen für {p_vorname} {p_nachname}\n')

    def create_dirs(self):
        '''
        Funktion erstellt Ausgabeordner falls diese
        noch nicht existieren.
        '''
        if self.vorstand_radio.isChecked():
            path = os.path.join(self.output_folder, 'Vorstand')
            if not os.path.exists(path):
                os.makedirs(path)
            path = os.path.join(path, str(self.year))
            if not os.path.exists(path):
                os.makedirs(path)

        if self.jugend_radio.isChecked():
            path = os.path.join(self.output_folder, 'Jugend')
            if not os.path.exists(path):
                os.makedirs(path)
            path = os.path.join(path, str(self.year))
            if not os.path.exists(path):
                os.makedirs(path)

    def show_info(self):
        msg_box = QMessageBox()
        msg_box.setWindowTitle('Info')
        msg_box.setText(
            f'Version {VERSION.replace("v", "")}\nJoel Klein\nhttps://github.com/joe2824/dlrg_briefbogen_generator')
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.exec()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(os.path.join(basedir, 'icons', 'icon.ico')))
    selector = DLRGBriefbogenGenerator()
    selector.show()
    sys.exit(app.exec())
