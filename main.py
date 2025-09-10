# このスクリプトは、zipファイル内に含まれるファイルリストをCSVにエクスポートするためのものです。
# main.py
import os
import sys
import glob
import zipfile

from PyQt6 import QtCore
from PyQt6.QtWidgets import (QApplication, QWidget,
                             QMainWindow, QLabel,
                             QLineEdit, QPushButton,
                             QHBoxLayout, QVBoxLayout,
                             QCheckBox, QFrame,
                             QSpacerItem, QSizePolicy,
                             QTextEdit,
                             QFileDialog, QMessageBox,
                             QComboBox)
from PyQt6.QtGui import (QIcon, QDragEnterEvent, QDropEvent)
from PyQt6.QtCore import Qt

class MainWindow(QMainWindow):
    dir_path = ""
    zip_file_paths = []
    file_names = []
    step = 0

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # ドロップイベントを受け付ける
        self.setAcceptDrops(True)

        # ウィンドウの設定
        self.setWindowTitle("zipファイル内ファイルリストをエクスポート")
        self.setWindowIcon(QIcon("app.png"))
        self.setFixedSize(600,400)
        # MainWindowの中央ウィジェットを設定
        self.centralWidget = QWidget(parent=self)
        self.centralWidget.setObjectName("centralWidget")
        #中央ウィジェットにメインになる垂直レイアウトウィジェットを配置
        self.verticalLayoutWidget = QWidget(parent=self.centralWidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(10, 10, 580, 380))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 10)
        self.verticalLayout.setObjectName("verticalLayout")

        # 1 垂直レイアウトウィジェットの1行目の要素
        self.label_log = QLabel(parent=self.verticalLayoutWidget)
        self.label_log.setObjectName("label_log")
        self.label_log.setText("ここにzipファイルまたはzipファイルをまとめたフォルダをドラッグ＆ドロップしてください")
        self.verticalLayout.addWidget(self.label_log)
        # 2 垂直レイアウトウィジェットの2行目の要素
        self.textEdit_log = QTextEdit(parent=self.verticalLayoutWidget)
        self.textEdit_log.setObjectName("textEdit_log")
        self.verticalLayout.addWidget(self.textEdit_log)
        self.textEdit_log.setAcceptDrops(False)  # QTextEdit自体はドロップを受け付けないようにする  

        # 3 垂直レイアウトウィジェットの3行目の要素（横レイアウトに3要素配置）
        self.horizontalLayout_3 = QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        # 3-1 垂直レイアウトウィジェットの3行目の要素（横レイアウトの1要素目）
        spacerItem = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem)
        # 3-2 垂直レイアウトウィジェットの3行目の要素（横レイアウトの2要素目）
        self.btn_run = QPushButton(parent=self.verticalLayoutWidget)
        self.btn_run.setObjectName("btn_run")
        self.btn_run.setText("実行")
        self.btn_run.clicked.connect(self.export_files_in_zip)
        self.horizontalLayout_3.addWidget(self.btn_run)
        # 3-3 垂直レイアウトウィジェットの3行目の要素（横レイアウトの3要素目）
        self.btn_quit = QPushButton(parent=self.verticalLayoutWidget)
        self.btn_quit.setObjectName("btn_quit")
        self.btn_quit.setText("終了")
        self.btn_quit.clicked.connect(self.close)
        self.horizontalLayout_3.addWidget(self.btn_quit)
        self.verticalLayout.addLayout(self.horizontalLayout_3)

        # MainWindowに中央ウィジェットをセット
        self.setCentralWidget(self.centralWidget)
        # メニューバーの設定
        self.menubar = self.menuBar()
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 22))
        self.menubar.setObjectName("menubar")
        # ステータスバーの設定
        self.statusbar = self.statusBar()
        self.statusbar.setObjectName("statusbar")
        self.statusbar.showMessage("Ready")
        
        self.show()

    # ドラッグイベント
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls:
            event.acceptProposedAction()
        else:
            self.textEdit_log.append("無効なパスです。")
            event.ignore()

    # ドロップイベント
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
            links = []
            for url in event.mimeData().urls():
                links.append(str(url.toLocalFile()))
            print(links)
            self.zip_file_paths = []
            for path in links:
                # zipファイルを受け付ける
                if os.path.isfile(path) and path.lower().endswith(".zip"):
                    self.zip_file_paths.append(path)
                # フォルダを受け付ける
                elif os.path.isdir(path):
                    self.zip_file_paths.extend(glob.glob(os.path.join(path, "*.zip")))
            if len(self.zip_file_paths) == 0:
                self.textEdit_log.append("zipファイルが見つかりません。")
                return
            self.textEdit_log.append(f"zipファイルを{len(self.zip_file_paths)}件読み込みました。")
            for zp in self.zip_file_paths:
                self.textEdit_log.append(zp)
    
    def export_files_in_zip(self):
        if not self.zip_file_paths:
            self.textEdit_log.append("zipファイルが指定されていません。")
            return
        file_names = []
        all_file_names = []
        for zip_path in self.zip_file_paths:
            file_names = self.get_file_names_in_zip(zip_path)
            all_file_names.extend(file_names)
    
        if not all_file_names:
            self.textEdit_log.append("zipファイル内にファイルが見つかりません。")
            return
        output_path, _ = QFileDialog.getSaveFileName(self, "ファイルを保存", "", "csv Files (*.csv);;All Files (*)")
        if output_path:
            self.save_to_csv(all_file_names, output_path)
            self.textEdit_log.append("処理が完了しました。")


    def get_file_names_in_zip(self, zip_path):
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                files_in_zip = zip_ref.namelist()
                file_names = []
                file_names_list = []
                for f in files_in_zip:
                    if os.path.basename(f):  # 空の名前を除外
                        file_names = (os.path.basename(zip_path) , os.path.basename(f))
                        file_names_list.append(file_names)
                return file_names_list
        except zipfile.BadZipFile:
            self.textEdit_log.append(f"{zip_path} は無効なzipファイルです。")
            return []

    def save_to_csv(self, file_names, output_path):
        import csv
        try:
            with open(output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['zipファイル', 'ファイル名'])
                for name in file_names:
                    writer.writerow([name[0], name[1]])
            self.textEdit_log.append(f"ファイルリストを {output_path} にエクスポートしました。")
        except Exception as e:
            self.textEdit_log.append(f"CSVのエクスポートに失敗しました: {e}")


def main(args):
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())

if __name__ == "__main__":
    main(sys.argv)