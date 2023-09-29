from email.mime.application import MIMEApplication

from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QMessageBox
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import yagmail
from PyQt5.QtWidgets import QFileDialog
import os


class Ui_FrmEnvioEmail(object):
    def setupUi(self, FrmEnvioEmail):
        FrmEnvioEmail.setObjectName("FrmEnvioEmail")
        FrmEnvioEmail.resize(661, 576)

        self.centralwidget = QtWidgets.QWidget(FrmEnvioEmail)
        self.centralwidget.setObjectName("centralwidget")
        self.EdtServidorSMTP = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtServidorSMTP.setGeometry(QtCore.QRect(119, 50, 191, 20))
        self.EdtServidorSMTP.setObjectName("EdtServidorSMTP")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(119, 30, 71, 16))
        self.label.setObjectName("label")
        self.EdtPorta = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtPorta.setGeometry(QtCore.QRect(319, 50, 71, 20))
        self.EdtPorta.setObjectName("EdtPorta")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(319, 30, 47, 13))
        self.label_2.setObjectName("label_2")
        self.RdSSL = QtWidgets.QRadioButton(self.centralwidget)
        self.RdSSL.setGeometry(QtCore.QRect(409, 50, 51, 17))
        self.RdSSL.setObjectName("RdSSL")
        self.RdTLS = QtWidgets.QRadioButton(self.centralwidget)
        self.RdTLS.setGeometry(QtCore.QRect(469, 50, 61, 17))
        self.RdTLS.setObjectName("RdTLS")
        self.EdtUsuario = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtUsuario.setGeometry(QtCore.QRect(119, 100, 241, 20))
        self.EdtUsuario.setObjectName("EdtUsuario")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(119, 80, 41, 16))
        self.label_3.setObjectName("label_3")
        self.EdtSenha = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtSenha.setGeometry(QtCore.QRect(369, 100, 161, 20))
        self.EdtSenha.setObjectName("EdtSenha")
        self.EdtSenha.setEchoMode(QtWidgets.QLineEdit.Password)  # Configurando o modo de exibição da senha
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(369, 80, 47, 13))
        self.label_4.setObjectName("label_4")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(44, 130, 581, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.EdtDe = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtDe.setGeometry(QtCore.QRect(100, 200, 231, 20))
        self.EdtDe.setObjectName("EdtDe")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(100, 180, 21, 16))
        self.label_5.setObjectName("label_5")
        self.EdtPara = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtPara.setGeometry(QtCore.QRect(350, 200, 221, 20))
        self.EdtPara.setObjectName("EdtPara")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(350, 180, 40, 13))
        self.label_6.setObjectName("label_6")
        self.EdtCopia = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtCopia.setGeometry(QtCore.QRect(100, 250, 471, 20))
        self.EdtCopia.setObjectName("EdtCopia")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(100, 230, 47, 13))
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(100, 280, 47, 13))
        self.label_8.setObjectName("label_8")
        self.EdtAssunto = QtWidgets.QLineEdit(self.centralwidget)
        self.EdtAssunto.setGeometry(QtCore.QRect(100, 300, 251, 20))
        self.EdtAssunto.setObjectName("EdtAssunto")
        self.EdtTexto = QtWidgets.QTextEdit(self.centralwidget)
        self.EdtTexto.setGeometry(QtCore.QRect(100, 350, 471, 161))
        self.EdtTexto.setObjectName("EdtTexto")
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(100, 330, 47, 13))
        self.label_9.setObjectName("label_9")
        self.BtnEnviar = QtWidgets.QPushButton(self.centralwidget)
        self.BtnEnviar.setGeometry(QtCore.QRect(497, 520, 75, 23))
        self.BtnEnviar.setObjectName("BtnEnviar")
        self.BtnAnexo = QtWidgets.QPushButton(self.centralwidget)
        self.BtnAnexo.setGeometry(QtCore.QRect(360, 300, 75, 23))
        self.BtnAnexo.setObjectName("BtnAnexo")
        FrmEnvioEmail.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(FrmEnvioEmail)
        self.statusbar.setObjectName("statusbar")
        FrmEnvioEmail.setStatusBar(self.statusbar)

        self.retranslateUi(FrmEnvioEmail)
        QtCore.QMetaObject.connectSlotsByName(FrmEnvioEmail)

        # Defina o ícone da janela
        icon_path = r'C:\Anderson\Python\HelpTecEnvioEmail\Icone.png'
        icon = QtGui.QIcon(icon_path)
        FrmEnvioEmail.setWindowIcon(icon)

        # Impedir maximização
        FrmEnvioEmail.setWindowFlags(FrmEnvioEmail.windowFlags() & ~QtCore.Qt.WindowMaximizeButtonHint)

        # Defina o fundo azul claro usando folha de estilo
        style_sheet = """
                            QWidget {
                                background-color: #E6F7FF; /* Cor de fundo azul claro */
                            }
                        """
        FrmEnvioEmail.setStyleSheet(style_sheet)

    def retranslateUi(self, FrmEnvioEmail):
        _translate = QtCore.QCoreApplication.translate
        FrmEnvioEmail.setWindowTitle(_translate("FrmEnvioEmail", "HelpTech - Teste envio de e-mail"))
        self.label.setText(_translate("FrmEnvioEmail", "Servidor SMTP"))
        self.label_2.setText(_translate("FrmEnvioEmail", "Porta"))
        self.RdSSL.setText(_translate("FrmEnvioEmail", "SSL"))
        self.RdTLS.setText(_translate("FrmEnvioEmail", "TLS 1.2"))
        self.label_3.setText(_translate("FrmEnvioEmail", "Usuário"))
        self.label_4.setText(_translate("FrmEnvioEmail", "Senha"))
        self.label_5.setText(_translate("FrmEnvioEmail", "De"))
        self.label_6.setText(_translate("FrmEnvioEmail", "Para"))
        self.label_7.setText(_translate("FrmEnvioEmail", "Cópia"))
        self.label_8.setText(_translate("FrmEnvioEmail", "Assunto"))
        self.label_9.setText(_translate("FrmEnvioEmail", "Texto"))
        self.BtnEnviar.setText(_translate("FrmEnvioEmail", "Enviar"))
        self.BtnAnexo.setText(_translate("FrmEnvioEmail", "Anexo"))

class MinhaJanelaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()

        # Crie uma instância da classe Ui_MainWindow
        self.ui = Ui_FrmEnvioEmail()
        self.ui.setupUi(self)

        # Conecte a função enviar_email ao botão de enviar
        self.ui.BtnEnviar.clicked.connect(self.enviar_email)

        # Conectar o botão BtnAnexo
        self.ui.BtnAnexo.clicked.connect(self.anexar_arquivo)

        # Lista para armazenar os caminhos dos arquivos anexados
        self.arquivos_anexados = []

    def enviar_email(self):
        servidor_smtp = self.ui.EdtServidorSMTP.text()
        porta = self.ui.EdtPorta.text()
        ssl = self.ui.RdSSL.isChecked()
        tls = self.ui.RdTLS.isChecked()
        usuario = self.ui.EdtUsuario.text()
        senha = self.ui.EdtSenha.text()
        de = self.ui.EdtDe.text()
        para = self.ui.EdtPara.text()
        copia = self.ui.EdtCopia.text()
        assunto = self.ui.EdtAssunto.text()
        texto = self.ui.EdtTexto.toPlainText()

        # Configurar a mensagem de e-mail
        mensagem = MIMEMultipart()
        mensagem['From'] = de
        mensagem['To'] = para
        mensagem['Cc'] = copia
        mensagem['Subject'] = assunto

        corpo_email = MIMEText(texto, 'plain')
        mensagem.attach(corpo_email)

        # Adicionar os arquivos anexados à mensagem de e-mail
        for arquivo in self.arquivos_anexados:
            with open(arquivo, "rb") as anexo:
                part = MIMEApplication(anexo.read(), Name=os.path.basename(arquivo))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo)}"'
                mensagem.attach(part)

        # Conectar ao servidor SMTP e enviar o e-mail
        try:
            servidor = smtplib.SMTP(servidor_smtp, porta)
            servidor.starttls()  # Ativar a criptografia TLS
            servidor.login(usuario, senha)
            servidor.sendmail(de, [para, copia], mensagem.as_string())
            servidor.quit()
            QMessageBox.information(self, "Sucesso", "E-mail enviado com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao enviar o e-mail:\n{str(e)}")

    def anexar_arquivo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly  # Se quiser apenas arquivos de leitura
        file_names, _ = QFileDialog.getOpenFileNames(self, "Selecionar Arquivos para Anexar", "",
                                                     "Todos os Arquivos (*)", options=options)

        if file_names:
            # Os arquivos selecionados são armazenados em 'file_names'
            print("Arquivos selecionados:", file_names)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = MinhaJanelaPrincipal()
    janela.show()
    sys.exit(app.exec_())

