import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QProgressBar, QTextEdit
)
from PyQt5 import QtCore
from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import ssl

# vCenter bağlantısı
def connect_to_vcenter(host, user, password):
    try:
        context = ssl._create_unverified_context()
        si = SmartConnect(host=host, user=user, pwd=password, sslContext=context)
        return si
    except Exception as e:
        return str(e)

# VM'yi recursive olarak bul
def find_vm_by_name(folder, vm_name):
    for entity in folder.childEntity:
        if isinstance(entity, vim.VirtualMachine) and entity.name == vm_name:
            return entity
        elif hasattr(entity, 'childEntity'):  # Eğer klasör varsa içine gir
            vm = find_vm_by_name(entity, vm_name)
            if vm:
                return vm
    return None

# Excel'den okunan verilerle VM'lere Custom Attributes ekleme
def process_excel_and_add_attributes(excel_file, si, progress_bar, log_text_edit):
    try:
        df = pd.read_excel(excel_file, dtype=str).fillna('')  # Boş hücreler için boş string
        total_vms = len(df)
        progress_bar.setMaximum(total_vms)
        progress_bar.setValue(0)

        for index, row in df.iterrows():
            vm_name = row['VM Name']
            attributes = row.drop(labels=['VM Name']).to_dict()

            content = si.RetrieveContent()
            vm = None

            for datacenter in content.rootFolder.childEntity:
                if isinstance(datacenter, vim.Datacenter):
                    vm_folder = datacenter.vmFolder
                    vm = find_vm_by_name(vm_folder, vm_name)
                    if vm:
                        break

            if vm is None:
                log_text_edit.append(f"VM '{vm_name}' bulunamadı!")
                continue

            for key, value in attributes.items():
                log_text_edit.append(f"'{vm.name}' VM'si için '{key}' alanına '{value}' atanıyor...")
                try:
                    vm.SetCustomValue(key=key, value=value)
                except Exception as e:
                    log_text_edit.append(f"'{vm.name}' VM'ye Custom Attributes eklenirken hata oluştu: {e}")
            
            progress_bar.setValue(index + 1)
            QApplication.processEvents()  # GUI'yi güncelleme

        log_text_edit.append("İşlem Tamamlandı!")
    except Exception as e:
        log_text_edit.append(f"Excel dosyasını işlerken hata oluştu: {e}")

# PyQt5 Arayüz
class VCenterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("vCenter Bilgi Girişi")
        self.layout = QVBoxLayout()

        # vCenter adresi
        self.layout.addWidget(QLabel("vCenter Adresi:"))
        self.vcenter_host = QLineEdit()
        self.layout.addWidget(self.vcenter_host)

        # vCenter kullanıcı adı
        self.layout.addWidget(QLabel("Kullanıcı Adı:"))
        self.vcenter_user = QLineEdit()
        self.layout.addWidget(self.vcenter_user)

        # vCenter şifresi
        self.layout.addWidget(QLabel("Şifre:"))
        self.vcenter_password = QLineEdit()
        self.vcenter_password.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(self.vcenter_password)

        # Excel dosyası
        self.layout.addWidget(QLabel("Excel Dosyası:"))
        self.excel_file_path = QLineEdit()
        self.layout.addWidget(self.excel_file_path)
        self.browse_button = QPushButton("Gözat")
        self.browse_button.clicked.connect(self.browse_file)
        self.layout.addWidget(self.browse_button)

        # Gönder butonu
        self.submit_button = QPushButton("Bağlan ve Göm")
        self.submit_button.clicked.connect(self.submit)
        self.layout.addWidget(self.submit_button)

        # İlerleme çubuğu
        self.progress_bar = QProgressBar()
        self.layout.addWidget(self.progress_bar)

        # İşlem adımlarını gösteren metin alanı
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        self.layout.addWidget(self.log_text_edit)

        # İmza ekleme
        self.signature_label = QLabel("Mehmet GÖZLEMECİ", self)
        self.signature_label.setAlignment(QtCore.Qt.AlignCenter)  # Ortalanmış metin
        self.layout.addWidget(self.signature_label)

        self.setLayout(self.layout)

    def browse_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Excel Dosyasını Seç", "", "Excel Files (*.xlsx)")
        if file_path:
            self.excel_file_path.setText(file_path)

    def submit(self):
        vcenter_host = self.vcenter_host.text()
        vcenter_user = self.vcenter_user.text()
        vcenter_password = self.vcenter_password.text()
        excel_file = self.excel_file_path.text()

        if not vcenter_host or not vcenter_user or not vcenter_password or not excel_file:
            QMessageBox.warning(self, "Eksik Bilgi", "Lütfen tüm alanları doldurun!")
            return

        # vCenter'a bağlan
        si = connect_to_vcenter(vcenter_host, vcenter_user, vcenter_password)
        if isinstance(si, str):  # Hata varsa
            QMessageBox.critical(self, "Bağlantı Hatası", f"vCenter'a bağlanırken hata oluştu: {si}")
        else:
            QMessageBox.information(self, "Başarılı", "vCenter'a başarılı bir şekilde bağlanıldı.")
            try:
                process_excel_and_add_attributes(excel_file, si, self.progress_bar, self.log_text_edit)
                QMessageBox.information(self, "Başarılı", "Excel dosyasındaki veriler işlendi.")
            except Exception as e:
                QMessageBox.critical(self, "İşleme Hatası", f"Veriler işlenirken hata oluştu: {e}")
            finally:
                Disconnect(si)

# Uygulama çalıştırma
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VCenterApp()
    window.show()
    sys.exit(app.exec_())
