import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QLineEdit, QGroupBox
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

class FileComparer(QWidget):
    def __init__(self):
        super().__init__()

        self.file1 = None
        self.file2 = None

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Tablo Karşılaştırıcı")
        self.resize(600, 450)  # Başlangıç boyutu

        # Ana layout
        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignCenter)  # Tüm öğeleri ortala

        # Font ayarları
        font = QFont()
        font.setPointSize(12)  # Yazı boyutunu büyüt

        # Bölüm 1: Tablo Dosyalarını Seç
        file_section = QGroupBox("1. Tablo Dosyalarını Seç")
        file_section_layout = QVBoxLayout()

        self.label_file1 = QLabel("Tablo 1: Henüz seçilmedi")
        self.label_file1.setFont(font)
        file_section_layout.addWidget(self.label_file1, alignment=Qt.AlignCenter)

        button_file1 = QPushButton("Tablo 1 Dosya Seç")
        button_file1.setFont(font)
        button_file1.setFixedSize(400, 50)
        file_section_layout.addWidget(button_file1, alignment=Qt.AlignCenter)
        button_file1.clicked.connect(self.select_file1)

        self.label_file2 = QLabel("Tablo 2: Henüz seçilmedi")
        self.label_file2.setFont(font)
        file_section_layout.addWidget(self.label_file2, alignment=Qt.AlignCenter)

        button_file2 = QPushButton("Tablo 2 Dosya Seç")
        button_file2.setFont(font)
        button_file2.setFixedSize(400, 50)
        file_section_layout.addWidget(button_file2, alignment=Qt.AlignCenter)
        button_file2.clicked.connect(self.select_file2)

        file_section.setLayout(file_section_layout)
        main_layout.addWidget(file_section)

        # Bölüm 2: Çıktı Dosyasını Belirt
        output_section = QGroupBox("2. Çıktı Dosyasını Belirt")
        output_section_layout = QVBoxLayout()

        self.output_file_input = QLineEdit()
        self.output_file_input.setFont(font)
        self.output_file_input.setFixedSize(400, 50)
        self.output_file_input.setPlaceholderText("Çıktı dosyası ismini girin")
        output_section_layout.addWidget(self.output_file_input, alignment=Qt.AlignCenter)

        output_section.setLayout(output_section_layout)
        main_layout.addWidget(output_section)

        # Bölüm 3: Karşılaştır
        compare_section = QGroupBox("3. Karşılaştır")
        compare_section_layout = QVBoxLayout()

        button_process = QPushButton("Karşılaştır")
        button_process.setFont(font)
        button_process.setFixedSize(200, 50)
        compare_section_layout.addWidget(button_process, alignment=Qt.AlignCenter)
        button_process.clicked.connect(self.process_files)

        compare_section.setLayout(compare_section_layout)
        main_layout.addWidget(compare_section)

        self.setLayout(main_layout)

    def select_file1(self):
        self.file1, _ = QFileDialog.getOpenFileName(self, "Tablo 1 dosyasını seçin", "", "Excel Files (*.xlsx *.xls)")
        if self.file1:
            self.label_file1.setText(f"Tablo 1: {self.file1}")

    def select_file2(self):
        self.file2, _ = QFileDialog.getOpenFileName(self, "Tablo 2 dosyasını seçin", "", "Excel Files (*.xlsx *.xls)")
        if self.file2:
            self.label_file2.setText(f"Tablo 2: {self.file2}")

    def process_files(self):
        if not self.file1 or not self.file2:
            print("Lütfen her iki dosyayı da seçin.")
            return

        output_file = self.output_file_input.text().strip()
        if not output_file:
            print("Lütfen sonuç dosyası adını girin.")
            return

        if not output_file.endswith(".xlsx"):
            output_file += ".xlsx"

        df1 = pd.read_excel(self.file1)
        df2 = pd.read_excel(self.file2)

        df1.columns = df1.columns.str.strip()
        df2.columns = df2.columns.str.strip()

        df1["TARİH"] = pd.to_datetime(df1["TARİH"], dayfirst=True).dt.strftime("%d/%m/%Y")
        df2["TARİH"] = pd.to_datetime(df2["TARİH"], dayfirst=True).dt.strftime("%d/%m/%Y")

        df1["BORÇ"] = df1["BORÇ"].fillna(0).map(lambda x: f"{float(x):.2f}")
        df1["ALACAK"] = df1["ALACAK"].fillna(0).map(lambda x: f"{float(x):.2f}")
        df2["BORÇ"] = df2["BORÇ"].fillna(0).map(lambda x: f"{float(x):.2f}")
        df2["ALACAK"] = df2["ALACAK"].fillna(0).map(lambda x: f"{float(x):.2f}")

        merged_df = pd.merge(
            df1, df2,
            on=["TARİH", "BORÇ", "ALACAK"],
            how="outer",
            indicator=True,
            suffixes=('_df1', '_df2')
        )

        diff1 = merged_df[merged_df["_merge"] == "left_only"]
        diff1_columns = ["TARİH", "BORÇ", "ALACAK"] + [col for col in merged_df.columns if col.endswith('_df1')]
        diff1 = diff1[diff1_columns]
        diff1.columns = [col.replace('_df1', '') for col in diff1.columns]

        diff2 = merged_df[merged_df["_merge"] == "right_only"]
        diff2_columns = ["TARİH", "BORÇ", "ALACAK"] + [col for col in merged_df.columns if col.endswith('_df2')]
        diff2 = diff2[diff2_columns]
        diff2.columns = [col.replace('_df2', '') for col in diff2.columns]

        with pd.ExcelWriter(output_file) as writer:
            diff1.to_excel(writer, sheet_name="Tablo 1'de Olup Tablo 2'de Yok", index=False)
            diff2.to_excel(writer, sheet_name="Tablo 2'de Olup Tablo 1'de Yok", index=False)

        print(f"Farklı kayıtlar '{output_file}' dosyasına kaydedildi.")
        QApplication.quit()

app = QApplication(sys.argv)
window = FileComparer()
window.show()
sys.exit(app.exec_())
