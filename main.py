# Modulos

import os
import io
import sys
import time
import smtplib
import winsound
import threading
import configparser
from os import listdir
from os.path import isfile, join
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from PyQt5 import uic, QtGui
from PyQt5.QtCore import QEventLoop
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog
import xlrd


# Rutas Archivos

directory_path = os.path.dirname(__file__)
gui_path = os.path.join(directory_path, "mainWindow.ui")
config_path = os.path.join(directory_path, "config.ini")
icon_path = os.path.join(directory_path, "LogoNA.png")
html_path = os.path.join(directory_path, "BodyInvoiceNA.html")
with io.open(file=html_path, mode="r", encoding="utf-8") as html_file:
    message = html_file.read()
    html_file.close()

config = configparser.ConfigParser()
config.read(config_path)

email_from = config["DEFAULT"]["email_from"]
email_pass = config["DEFAULT"]["email_pass"]
excel_path = config["DEFAULT"]["excel_path"]
invoices_path = config["DEFAULT"]["invoices_path"]


# Funciones

def make_json_from_data(column_names, row_data): # Convierte el Diccioanrio en Json
    row_list = []
    for item in row_data:
        json_obj = {}
        for i in range(0, column_names.__len__()):
            json_obj[column_names[i]] = item[i]
        row_list.append(json_obj)
    return row_list

def xls_to_dict(workbook_url):  # Convierte todas las Hojas del libro Excel en Diccionario
    workbook_dict = {}
    book = xlrd.open_workbook(workbook_url)
    sheets = book.sheets()
    for sheet in sheets:
        workbook_dict[sheet.name] = {}
        columns = sheet.row_values(0)
        rows = []
        for row_index in range(1, sheet.nrows):
            row = sheet.row_values(row_index)
            rows.append(row)
        sheet_data = make_json_from_data(columns, rows)
        workbook_dict[sheet.name] = sheet_data
    return workbook_dict

def sound_system(type_sound):
    # type = "SystemAsterisk" ---> Información
    # type = "SystemHand" ---> Error
    winsound.PlaySound(type_sound, winsound.SND_ALIAS)

def list_directory(path):
    return [arch for arch in listdir(path) if isfile(join(path, arch))]


# Clase Inicio Ventana

class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(gui_path, self)
        self.setWindowIcon(QtGui.QIcon(icon_path))
        self.txtEmail.setText(email_from)
        self.txtPass.setText(email_pass)
        self.txtExcelPath.setText(excel_path)
        self.txtInvoicesPath.setText(invoices_path)
        self.txtEmail.textChanged.connect(self.changed_text_email)
        self.txtPass.textChanged.connect(self.changed_text_pass)
        self.txtExcelPath.textChanged.connect(self.changed_text_excel_path)
        self.txtInvoicesPath.textChanged.connect(self.changed_text_invoices_path)
        self.btnShowPass.clicked.connect(self.show_password)
        self.btnExcelPath.clicked.connect(self.get_excel_dialog_path)
        self.btnInvoicesPath.clicked.connect(self.get_invoices_dialog_path)
        self.btnClearListProcess.clicked.connect(self.clear_list_process)
        self.btnRunProcess.clicked.connect(self.run_process)

    # Ver Contraseña del Email
    def show_password(self):
        if self.btnShowPass.text() == "Ver":
            self.txtPass.setEchoMode(0)
            self.btnShowPass.setText("Ocultar")
        else:
            self.txtPass.setEchoMode(2)
            self.btnShowPass.setText("Ver")

    # Cuadro de Diálogo para seleccionar Ruta Excel
    def get_excel_dialog_path(self):
        excel_dialog_path = QFileDialog.getOpenFileName(
            self,
            "Seleccione el Archivo Excel",
            "",
            "*.xlsx"
        )
        if excel_dialog_path[0] != "":
            self.txtExcelPath.setText(excel_dialog_path[0])

    # Cuadro de Diálogo para seleccionar Ruta Recibos
    def get_invoices_dialog_path(self):
        invoices_dialog_path = QFileDialog.getExistingDirectory(
            self,
            "Seleccione la Carpeta de Recibos"
        )
        if invoices_dialog_path != "":
            self.txtInvoicesPath.setText(invoices_dialog_path)

    # Al detectar un cambio en el txtEmail
    def changed_text_email(self):
        new_email_from = self.txtEmail.text()
        config.set("DEFAULT", "email_from", new_email_from)
        with open(file=config_path, mode="w", encoding="utf-8") as save:
            config.write(save)

    # Al detectar un cambio en el txtPass
    def changed_text_pass(self):
        new_email_pass = self.txtPass.text()
        config.set("DEFAULT", "email_pass", new_email_pass)
        with open(file=config_path, mode="w", encoding="utf-8") as save:
            config.write(save)

    # Al detectar un cambio en el txtExcelPath
    def changed_text_excel_path(self):
        new_excel_path = self.txtExcelPath.text()
        config.set("DEFAULT", "ruta_excel", new_excel_path)
        with open(file=config_path, mode="w", encoding="utf-8") as save:
            config.write(save)

    # Al detectar un cambio en el txtInvoicesPath
    def changed_text_invoices_path(self):
        new_invoices_path = self.txtInvoicesPath.text()
        config.set("DEFAULT", "ruta_recibos", new_invoices_path)
        with open(file=config_path, mode="w", encoding="utf-8") as save:
            config.write(save)

    # Limpiar Lista de Proceso
    def clear_list_process(self):
        self.listProcess.clear()

    # Ejecutar Proceso
    def run_process(self):
        # Variables de los valores de los textbox
        email_from_process = str(self.txtEmail.text())
        email_pass_process = str(self.txtPass.text())
        excel_path_process = str(self.txtExcelPath.text())
        invoices_path_process = str(self.txtInvoicesPath.text())
        # Validar Cuadros vacios y Rutas
        validation_fields = (
            email_from_process,
            email_pass_process,
            excel_path_process,
            invoices_path_process,
        )
        if "" in validation_fields:
            threading.Thread(target = sound_system, args = ("SystemHand",)).start()
            QMessageBox.critical(self, "Error!", "Una o varias Cajas de Texto están Vacías!")
            return
        if not os.path.exists(excel_path_process):
            threading.Thread(target = sound_system, args = ("SystemHand",)).start()
            QMessageBox.critical(self, "Error!", "Ruta de Excel no existe!")
            return
        if not os.path.exists(invoices_path_process):
            threading.Thread(target = sound_system, args = ("SystemHand",)).start()
            QMessageBox.critical(self, "Error!", "Ruta de Recibos no existe!")
            return
        try:
            invoices_dict = xls_to_dict(excel_path_process)
            invoices_dict = invoices_dict["ClienteNA"]
            invoices_list = list_directory(invoices_path_process)
            # Crear Servidor e iniciar sesión
            try:
                server = smtplib.SMTP_SSL("smtpout.secureserver.net", 465)
                server.ehlo()
                server.login(email_from_process, email_pass_process)
            except:
                threading.Thread(target = sound_system, args = ("SystemHand",)).start()
                QMessageBox.critical(self, "Error!", "Ha fallado el inicio de sesión!")
                return
            # Crear bucle para envío de email
            i = 1
            len_dict = len(invoices_dict)
            if len(invoices_list) != 0:
                for row in invoices_dict:
                    num_invoice = str(row["n_factura"]).replace(".0", "")
                    name = str(row["nombre_mayus"])
                    email_to = str(row["email"])
                    progress_percen = str(round((i/len_dict)*100, 2)) + "%"
                    msg_insert_item = f"Error ({progress_percen}) ---> ({num_invoice}) {name}"
                    i = i + 1
                    if num_invoice + ".pdf" in invoices_list:
                        if email_to == "":
                            self.listProcess.insertItem(
                                0,
                                f"{msg_insert_item} (Email vacío)"
                            )
                        else:
                            # Objeto y parámetros del mensaje
                            msg = MIMEMultipart()
                            bcc = ["cerm1202@gmail.com"]
                            msg["From"] = f"Neurocognitive Academy <{email_from_process}>"
                            msg["To"] = email_to
                            msg["Subject"] = "Recibo de Pago de Neurocognitive Academy"
                            # Agreagar el mensaje al cuerpo
                            body = MIMEText(message, "html")
                            msg.attach(body)
                            # Adjuntar archivo
                            file_invoice_path = f"{invoices_path_process}/{num_invoice}.pdf"
                            attached = MIMEApplication(open(file_invoice_path,"rb").read())
                            attached.add_header(
                                "Content-Disposition",
                                "attachment",
                                filename=f"{num_invoice}.pdf"
                            )
                            msg.attach(attached)
                            # Enviar email
                            try:
                                to_addresses = [email_to] + bcc
                                server.sendmail(msg["From"], to_addresses, msg.as_string())
                            except:
                                self.listProcess.insertItem(
                                    0,
                                    f"{msg_insert_item} (Fallo de Envío)"
                                )
                                err = sys.exc_info()[1]
                                err = err.args[0]
                                if err in (550, 552):
                                    threading.Thread(
                                        target=sound_system,
                                        args=("SystemHand",)
                                    ).start()
                                    QMessageBox.critical(
                                        self,
                                        "Error!",
                                        f"Error {err}: Ha superado su cuota de envíos de 250!"
                                    )
                                    return
                            else:
                                self.listProcess.insertItem(
                                    0,
                                    f"Succes ({msg_insert_item} (Email Enviado)"
                                )
                                # Actualiza con eventos en proceso
                                QApplication.processEvents(QEventLoop.ExcludeUserInputEvents)
                                self.listProcess.repaint()
                                time.sleep(1)
                    else:
                        self.listProcess.insertItem(
                            0,
                            f"Error ({msg_insert_item} (Sin Recibo)"
                        )
            else:
                threading.Thread(target = sound_system, args = ("SystemExclamation",)).start()
                QMessageBox.critical(self, "Atención!", "Directorio de recibos vacío!")
                return
        except:
            threading.Thread(target = sound_system, args = ("SystemHand",)).start()
            QMessageBox.critical(self, "Error!", "Ha Ocurrido un error!")
        else:
            server.quit()
            threading.Thread(target = sound_system, args = ("SystemAsterisk",)).start()
            QMessageBox.information(self, "Información!", "El proceso ha finalizado con Éxito!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
