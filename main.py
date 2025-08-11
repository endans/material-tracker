#!/usr/bin/env python3
# main.py - Material Tracker (versi dikembangkan)
import sys
import csv
import os
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QSpinBox, QFileDialog, QGroupBox, QMessageBox, QListWidget,
    QMenu, QAbstractItemView, QSizePolicy, QSpacerItem, QDialog, QDateEdit,
    QCompleter
)
from PySide6.QtGui import QIcon, QDesktopServices, QAction, QPalette
from PySide6.QtCore import Qt, QUrl, QDate

import json

APP_NAME = "Log Material Gudang CKT Purwokerto"
VERSION = "1.3.1"

DATA_DIR = os.path.join(os.path.expanduser("~"), ".material_tracker")
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

def data_path(filename):
    return os.path.join(DATA_DIR, filename)

def load_json(filename, default):
    try:
        with open(data_path(filename), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def save_json(filename, data):
    with open(data_path(filename), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

# Default containers
DEFAULT_DIVISI = []
DEFAULT_TIM = []
DEFAULT_MATERIAL = []
DEFAULT_AKSESORI = []

DIVISI = load_json("divisi.json", DEFAULT_DIVISI)
TIM = load_json("tim.json", DEFAULT_TIM)
MATERIAL = load_json("material.json", DEFAULT_MATERIAL)
AKSESORI = load_json("aksesori.json", DEFAULT_AKSESORI)

HISTORI_MAT = load_json("histori_kabel_aksesori.json", [])
HISTORI_ONT = load_json("histori_ont.json", [])
STOCK_ENTRIES = load_json("stock_entries.json", [])

LAST_SELECTION = load_json("last_selection.json", {"divisi": "", "tim": ""})

def apply_theme(app, dark=False):
    if dark:
        app.setStyle("Fusion")
        dark_palette = app.palette()
        dark_palette.setColor(QPalette.Window, Qt.black)
        dark_palette.setColor(QPalette.WindowText, Qt.white)
        dark_palette.setColor(QPalette.Base, Qt.darkGray)
        dark_palette.setColor(QPalette.AlternateBase, Qt.gray)
        dark_palette.setColor(QPalette.Text, Qt.white)
        dark_palette.setColor(QPalette.Button, Qt.gray)
        dark_palette.setColor(QPalette.ButtonText, Qt.white)
        app.setPalette(dark_palette)
    else:
        app.setStyle("Fusion")
        app.setPalette(app.style().standardPalette())

class EditableListBox(QWidget):
    def __init__(self, items, label):
        super().__init__()
        self.items = items
        self.layout = QVBoxLayout(self)
        self.grp = QGroupBox(label)
        vbox = QVBoxLayout()
        self.grp.setLayout(vbox)
        self.list_widget = QListWidget()
        self.list_widget.addItems(self.items)
        vbox.addWidget(self.list_widget)
        form_hbox = QHBoxLayout()
        self.edit = QLineEdit()
        form_hbox.addWidget(self.edit)
        self.btn_add = QPushButton("Tambah")
        self.btn_add.clicked.connect(self.add_item)
        self.btn_edit = QPushButton("Edit")
        self.btn_edit.clicked.connect(self.edit_item)
        self.btn_del = QPushButton("Hapus")
        self.btn_del.clicked.connect(self.del_item)
        form_hbox.addWidget(self.btn_add)
        form_hbox.addWidget(self.btn_edit)
        form_hbox.addWidget(self.btn_del)
        vbox.addLayout(form_hbox)
        self.layout.addWidget(self.grp)

    def add_item(self):
        val = self.edit.text().strip()
        if val and val not in self.items:
            self.items.append(val)
            self.list_widget.addItem(val)
            self.edit.clear()

    def del_item(self):
        idx = self.list_widget.currentRow()
        if idx >= 0:
            self.items.pop(idx)
            self.list_widget.takeItem(idx)

    def edit_item(self):
        idx = self.list_widget.currentRow()
        val = self.edit.text().strip()
        if idx >= 0 and val and val not in self.items:
            self.items[idx] = val
            self.list_widget.item(idx).setText(val)
            self.edit.clear()

    def reload(self):
        self.list_widget.clear()
        self.list_widget.addItems(self.items)

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Pengaturan")
        self.setMinimumWidth(500)
        layout = QVBoxLayout(self)

        self.divisi_box = EditableListBox(DIVISI, "Divisi")
        layout.addWidget(self.divisi_box)
        self.tim_box = EditableListBox(TIM, "Tim")
        layout.addWidget(self.tim_box)
        self.mat_box = EditableListBox(MATERIAL, "Material (Kabel)")
        layout.addWidget(self.mat_box)
        self.aks_box = EditableListBox(AKSESORI, "Aksesoris")
        layout.addWidget(self.aks_box)

        btn_layout = QHBoxLayout()
        btn_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.btn_save = QPushButton("Simpan Semua Pengaturan")
        self.btn_save.setStyleSheet("background:#2196F3;color:white;")
        self.btn_save.clicked.connect(self.save)
        btn_layout.addWidget(self.btn_save)
        layout.addLayout(btn_layout)

    def save(self):
        save_json("divisi.json", DIVISI)
        save_json("tim.json", TIM)
        save_json("material.json", MATERIAL)
        save_json("aksesori.json", AKSESORI)
        QMessageBox.information(self, "Berhasil", "Pengaturan disimpan.")
        self.accept()

    def reload(self):
        self.divisi_box.reload()
        self.tim_box.reload()
        self.mat_box.reload()
        self.aks_box.reload()

class TelegramReportTab(QWidget):
    """Generik tab laporan (MyRepublic, Asianet, Oxygen)."""
    def __init__(self, report_type):
        super().__init__()
        self.report_type = report_type
        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)

        if report_type == 'myrepublic':
            self.display_name = "Laporan MyRepublic"
            self.csv_file = os.path.expanduser('~/Reports/wifi_reports.csv')
        elif report_type == 'asianet':
            self.display_name = "Laporan Asianet"
            self.csv_file = os.path.expanduser('~/Reports/asianet_reports.csv')
        elif report_type == 'oxygen' or report_type == 'oxy':
            self.display_name = "Laporan IKR Oxygen"
            self.csv_file = os.path.expanduser('~/Reports/Oxy_Reports.csv')
        else:
            self.display_name = f"Laporan {report_type}"
            self.csv_file = os.path.expanduser(f'~/Reports/{report_type}_reports.csv')

        self.columns = ["Timestamp", "Subscription ID", "Customer", "SN", "Team"]
        self.init_ui()

    def init_ui(self):
        filter_layout = QHBoxLayout()
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Cari ... (nama pelanggan / SN / tim)")
        self.search_field.textChanged.connect(self.filter_table)
        filter_layout.addWidget(QLabel("🔍"))
        filter_layout.addWidget(self.search_field)
        filter_layout.addWidget(QLabel("Dari:"))
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addMonths(-1))
        filter_layout.addWidget(self.date_from)
        filter_layout.addWidget(QLabel("s/d"))
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        filter_layout.addWidget(self.date_to)
        self.date_from.dateChanged.connect(self.filter_table)
        self.date_to.dateChanged.connect(self.filter_table)
        self.layout.addLayout(filter_layout)

        self.tbl = QTableWidget()
        self.tbl.setColumnCount(len(self.columns))
        self.tbl.setHorizontalHeaderLabels(self.columns)
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        self.layout.addWidget(self.tbl)

        btn_layout = QHBoxLayout()
        reload_btn = QPushButton("Muat Ulang Laporan")
        reload_btn.setStyleSheet("background:#2196F3;color:white")
        reload_btn.clicked.connect(self.load_reports)
        btn_layout.addWidget(reload_btn)
        export_btn = QPushButton("Ekspor ke CSV")
        export_btn.clicked.connect(self.export_csv)
        btn_layout.addWidget(export_btn)
        btn_layout.addItem(QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.layout.addLayout(btn_layout)

        # Muat laporan pertama kali
        self.load_reports()

    def load_reports(self):
        self.raw_rows = []
        self.tbl.setRowCount(0)
        if os.path.exists(self.csv_file):
            try:
                with open(self.csv_file, newline='', encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        if self.report_type == 'asianet':
                            mapped = {
                                "Timestamp": row.get("Timestamp", "") or row.get("Tanggal", ""),
                                "Subscription ID": row.get("ID Pelanggan", ""),
                                "Customer": row.get("Nama Pelanggan", ""),
                                "SN": row.get("SN", ""),
                                "Team": row.get("Nama Teknisi", "")
                            }
                        else:
                            mapped = {
                                "Timestamp": row.get("Timestamp", "") or row.get("tanggal", ""),
                                "Subscription ID": row.get("Subscription ID", "") or row.get("ID Pelanggan", ""),
                                "Customer": row.get("Customer", "") or row.get("Nama Pelanggan", ""),
                                "SN": row.get("SN", "") or row.get("Serial Number", ""),
                                "Team": row.get("Team", "") or row.get("Nama Teknisi", "")
                            }
                        self.raw_rows.append(mapped)
                self.filter_table()
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Gagal memuat CSV:\n{self.csv_file}\n\n{e}")
        else:
            # kalau file tidak ada, jangan spam alert saat load pertama: cukup tunjukkan saat klik reload
            pass

    def filter_table(self):
        keyword = self.search_field.text().strip().lower()
        date_from = self.date_from.date()
        date_to = self.date_to.date()
        filtered = []
        for r in self.raw_rows:
            search_str = f"{r.get('Customer','')} {r.get('SN','')} {r.get('Team','')}".lower()
            if keyword and keyword not in search_str:
                continue
            date_str = (r.get("Timestamp", "") or "")[:10].replace('/', '-')
            tgl_dt = QDate.fromString(date_str, "yyyy-MM-dd")
            if not tgl_dt.isValid():
                tgl_dt = QDate.fromString(date_str, "dd-MM-yyyy")
            if not tgl_dt.isValid():
                # jika timestamp tidak dapat diparse, masukkan ke hasil agar tidak hilang
                filtered.append(r)
                continue
            if tgl_dt < date_from or tgl_dt > date_to:
                continue
            filtered.append(r)
        self.populate_table(filtered)

    def populate_table(self, rows):
        self.tbl.setRowCount(0)
        for row_data in rows:
            row = self.tbl.rowCount()
            self.tbl.insertRow(row)
            for i, col in enumerate(self.columns):
                val = row_data.get(col, "")
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                self.tbl.setItem(row, i, item)

    def export_csv(self):
        if not self.raw_rows:
            QMessageBox.information(self, "Info", "Tidak ada data untuk diekspor.")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self, f"Simpan Laporan sebagai CSV",
            f"{self.display_name.replace(' ', '_').lower()}_{datetime.now().strftime('%Y%m%d')}.csv",
            "CSV Files (*.csv)"
        )
        if filename:
            try:
                with open(filename, 'w', newline='', encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(self.columns)
                    for r in self.raw_rows:
                        writer.writerow([r.get(col, "") for col in self.columns])
                QMessageBox.information(self, "Berhasil", f"Berhasil menyimpan ke {filename}")
            except Exception as e:
                QMessageBox.warning(self, "Gagal", f"Gagal menyimpan CSV:\n{e}")

    def get_all_sn(self):
        sn_set = set()
        for r in getattr(self, "raw_rows", []):
            sn = r.get("SN", "")
            if sn:
                sn_set.add(sn.strip())
        return sn_set

class FormPengambilan(QWidget):
    def __init__(self, main):
        super().__init__()
        self.main = main
        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)

        # Meta group
        self.grp_meta = QGroupBox("Pengambilan")
        meta_layout = QHBoxLayout()
        self.grp_meta.setLayout(meta_layout)
        self.cmb_divisi = QComboBox()
        self.cmb_divisi.setEditable(True)
        self.cmb_divisi.setInsertPolicy(QComboBox.NoInsert)
        self.cmb_divisi.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.cmb_tim = QComboBox()
        self.cmb_tim.setEditable(True)
        self.cmb_tim.setInsertPolicy(QComboBox.NoInsert)
        self.cmb_tim.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        meta_layout.addWidget(QLabel("Divisi:"))
        meta_layout.addWidget(self.cmb_divisi)
        meta_layout.addWidget(QLabel("Tim:"))
        meta_layout.addWidget(self.cmb_tim)
        meta_layout.addStretch()
        self.layout.addWidget(self.grp_meta)

        # Kabel & Aksesori
        self.grp_kabel = QGroupBox("Kabel dan Aksesori")
        kabel_layout = QVBoxLayout()
        self.grp_kabel.setLayout(kabel_layout)
        self.tbl_kabel = QTableWidget(0, 2)
        self.tbl_kabel.setHorizontalHeaderLabels(["Nama Material", "Jumlah"])
        self.tbl_kabel.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        kabel_layout.addWidget(self.tbl_kabel)
        btn_row = QHBoxLayout()
        btn_kabel_add = QPushButton("Tambah Data")
        btn_kabel_add.clicked.connect(self.add_kabel_row)
        btn_row.addWidget(btn_kabel_add)
        btn_row.addStretch()
        kabel_layout.addLayout(btn_row)
        self.layout.addWidget(self.grp_kabel)

        # SN ONT
        self.grp_ont = QGroupBox("SN ONT")
        ont_layout = QVBoxLayout()
        self.grp_ont.setLayout(ont_layout)
        self.tbl_ont = QTableWidget(0, 1)
        self.tbl_ont.setHorizontalHeaderLabels(["Serial Number"])
        self.tbl_ont.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        ont_layout.addWidget(self.tbl_ont)
        btn_row2 = QHBoxLayout()
        btn_ont_add = QPushButton("Tambah Data")
        btn_ont_add.clicked.connect(self.add_ont_row)
        btn_ont_import = QPushButton("Import CSV SN ONT")
        btn_ont_import.clicked.connect(self.import_ont_csv)
        btn_row2.addWidget(btn_ont_add)
        btn_row2.addWidget(btn_ont_import)
        btn_row2.addStretch()
        ont_layout.addLayout(btn_row2)
        self.layout.addWidget(self.grp_ont)

        # Buttons
        button_layout = QHBoxLayout()
        self.btn_clear = QPushButton("Reset")
        self.btn_clear.setStyleSheet("background:#f44336;color:white;")
        self.btn_clear.clicked.connect(self.clear_form)
        button_layout.addWidget(self.btn_clear)
        button_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.btn_submit = QPushButton("Submit")
        self.btn_submit.setStyleSheet("background:#2196F3;color:white;")
        self.btn_submit.clicked.connect(self.submit)
        button_layout.addWidget(self.btn_submit, alignment=Qt.AlignRight)
        self.layout.addLayout(button_layout)

        self.reload_options()

    def create_completer_for(self, combo: QComboBox):
        # QCompleter dari daftar items, supaya pencarian lebih nyaman
        comp = QCompleter(combo.model(), combo)
        comp.setCompletionMode(QCompleter.PopupCompletion)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        combo.setCompleter(comp)

    def reload_options(self):
        # Divisi & Tim
        self.cmb_divisi.clear()
        self.cmb_divisi.addItems(DIVISI)
        if LAST_SELECTION.get("divisi"):
            idx = self.cmb_divisi.findText(LAST_SELECTION["divisi"])
            if idx >= 0:
                self.cmb_divisi.setCurrentIndex(idx)
            else:
                self.cmb_divisi.setEditText(LAST_SELECTION["divisi"])
        self.create_completer_for(self.cmb_divisi)

        self.cmb_tim.clear()
        self.cmb_tim.addItems(TIM)
        if LAST_SELECTION.get("tim"):
            idx = self.cmb_tim.findText(LAST_SELECTION["tim"])
            if idx >= 0:
                self.cmb_tim.setCurrentIndex(idx)
            else:
                self.cmb_tim.setEditText(LAST_SELECTION["tim"])
        self.create_completer_for(self.cmb_tim)

        # Reset tabel input
        self.tbl_kabel.setRowCount(0)
        self.tbl_ont.setRowCount(0)

    def clear_form(self):
        self.tbl_kabel.setRowCount(0)
        self.tbl_ont.setRowCount(0)

    def add_kabel_row(self):
        row = self.tbl_kabel.rowCount()
        self.tbl_kabel.insertRow(row)
        cmb = QComboBox()
        cmb.setEditable(True)
        cmb.setInsertPolicy(QComboBox.NoInsert)
        cmb.addItems(MATERIAL + AKSESORI)
        # set completer
        comp = QCompleter([*MATERIAL, *AKSESORI], self)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        cmb.setCompleter(comp)
        self.tbl_kabel.setCellWidget(row, 0, cmb)
        spin = QSpinBox()
        spin.setRange(1, 1000)
        self.tbl_kabel.setCellWidget(row, 1, spin)

    def add_ont_row(self):
        row = self.tbl_ont.rowCount()
        self.tbl_ont.insertRow(row)
        line = QLineEdit()
        line.setPlaceholderText("Scan/masukkan serial number ONT")
        self.tbl_ont.setCellWidget(row, 0, line)

    def import_ont_csv(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Pilih file CSV SN ONT untuk diimport", os.path.expanduser("~"), "CSV Files (*.csv);;All Files (*)")
        if not filename:
            return
        try:
            added = 0
            with open(filename, newline='', encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # coba ambil kolom SN atau Serial Number
                    sn = row.get("SN") or row.get("Serial Number") or row.get("sn") or row.get("serial_number") or ""
                    if not sn:
                        # jika baris tidak punya kolom SN, coba gabungkan semua kolom dan skip jika tidak ada
                        continue
                    tgl = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    entry = {
                        "tanggal": tgl,
                        "sn": str(sn).strip(),
                        "tim": LAST_SELECTION.get("tim", ""),
                        "divisi": LAST_SELECTION.get("divisi", "")
                    }
                    HISTORI_ONT.append(entry)
                    added += 1
            if added > 0:
                save_json("histori_ont.json", HISTORI_ONT)
                QMessageBox.information(self, "Import Selesai", f"Berhasil menambahkan {added} SN dari {filename}")
                self.main.reload_all()
            else:
                QMessageBox.information(self, "Import", "Tidak ada SN ditemukan di file.")
        except Exception as e:
            QMessageBox.warning(self, "Gagal Import", f"Gagal mengimpor CSV:\n{e}")

    def submit(self):
        divisi = self.cmb_divisi.currentText().strip()
        tim = self.cmb_tim.currentText().strip()
        tgl = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        mat_entries = []
        for row in range(self.tbl_kabel.rowCount()):
            widget_mat = self.tbl_kabel.cellWidget(row, 0)
            widget_qty = self.tbl_kabel.cellWidget(row, 1)
            if widget_mat and widget_qty:
                nama = widget_mat.currentText().strip()
                try:
                    qty = int(widget_qty.value())
                except Exception:
                    qty = 0
                if nama and qty > 0:
                    mat_entries.append({
                        "tanggal": tgl,
                        "deskripsi": nama,
                        "qty": qty,
                        "tim": tim,
                        "divisi": divisi
                    })

        ont_entries = []
        for row in range(self.tbl_ont.rowCount()):
            widget_sn = self.tbl_ont.cellWidget(row, 0)
            if widget_sn:
                sn = widget_sn.text().strip()
                if sn:
                    ont_entries.append({
                        "tanggal": tgl,
                        "sn": sn,
                        "tim": tim,
                        "divisi": divisi
                    })

        if not mat_entries and not ont_entries:
            QMessageBox.warning(self, "Validasi Gagal", "Masukkan minimal satu data material/ONT!")
            return

        global HISTORI_MAT, HISTORI_ONT, LAST_SELECTION
        if mat_entries:
            HISTORI_MAT += mat_entries
            save_json("histori_kabel_aksesori.json", HISTORI_MAT)
        if ont_entries:
            HISTORI_ONT += ont_entries
            save_json("histori_ont.json", HISTORI_ONT)

        # Simpan pilihan terakhir
        LAST_SELECTION["divisi"] = divisi
        LAST_SELECTION["tim"] = tim
        save_json("last_selection.json", LAST_SELECTION)

        # Jika divisi atau tim baru, tambahkan ke list dan simpan pengaturan
        if divisi and divisi not in DIVISI:
            DIVISI.append(divisi)
            save_json("divisi.json", DIVISI)
        if tim and tim not in TIM:
            TIM.append(tim)
            save_json("tim.json", TIM)

        QMessageBox.information(self, "Berhasil", "Data berhasil disimpan.")
        self.main.reload_all()
        self.clear_form()

class Resume(QWidget):
    def __init__(self, main, laporan_tabs):
        super().__init__()
        self.main = main
        self.laporan_tabs = laporan_tabs
        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)
        self.tabs = QTabWidget()
        self.tab_stock = QWidget()
        self.tab_kabel = QWidget()
        self.tab_ont = QWidget()
        self.tabs.addTab(self.tab_stock, "Stock")
        self.tabs.addTab(self.tab_kabel, "Material")
        self.tabs.addTab(self.tab_ont, "ONT")
        self.layout.addWidget(self.tabs)
        self.init_stock_tab()
        self.init_kabel_tab()
        self.init_ont_tab()
        self.reload_data()

    def init_stock_tab(self):
        stock_layout = QVBoxLayout(self.tab_stock)
        form_grp = QGroupBox("Tambah Stock Material Masuk")
        form_layout = QHBoxLayout()
        form_grp.setLayout(form_layout)
        self.stock_desc = QComboBox()
        self.stock_desc.setEditable(True)
        self.stock_desc.setInsertPolicy(QComboBox.NoInsert)
        self.stock_desc.addItems(MATERIAL + AKSESORI)
        # completer
        comp = QCompleter(MATERIAL + AKSESORI, self)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        self.stock_desc.setCompleter(comp)

        self.stock_qty = QSpinBox()
        self.stock_qty.setRange(1, 10000)
        form_layout.addWidget(QLabel("Nama Item:"))
        form_layout.addWidget(self.stock_desc)
        form_layout.addWidget(QLabel("Qty Masuk:"))
        form_layout.addWidget(self.stock_qty)
        self.btn_stock_add = QPushButton("Tambah Stock")
        self.btn_stock_add.setStyleSheet("background:#4CAF50;color:white;")
        self.btn_stock_add.clicked.connect(self.add_stock)
        form_layout.addWidget(self.btn_stock_add)
        stock_layout.addWidget(form_grp)

        self.tbl_stock = QTableWidget()
        self.tbl_stock.setColumnCount(7)
        self.tbl_stock.setHorizontalHeaderLabels([
            "No", "Tanggal Masuk", "Nama Item", "Qty Masuk", "Diambil Teknisi", "Stock Awal", "Aksi"
        ])
        self.tbl_stock.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_stock.setColumnWidth(0, 40)
        for i in range(1, 7):
            self.tbl_stock.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
        stock_layout.addWidget(self.tbl_stock)

    def init_kabel_tab(self):
        kabel_layout = QVBoxLayout(self.tab_kabel)
        search_row = QHBoxLayout()
        self.search_kabel = QLineEdit()
        self.search_kabel.setPlaceholderText("Cari nama material atau tim...")
        self.search_kabel.textChanged.connect(self.filter_kabel)
        search_row.addWidget(QLabel("🔍"))
        search_row.addWidget(self.search_kabel)
        kabel_layout.addLayout(search_row)
        self.tbl_kabel = QTableWidget()
        self.tbl_kabel.setColumnCount(6)
        self.tbl_kabel.setHorizontalHeaderLabels(["No", "Tanggal", "Deskripsi", "Qty", "Nama Tim", "Aksi"])
        self.tbl_kabel.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for i in range(1, 6):
            self.tbl_kabel.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
        kabel_layout.addWidget(self.tbl_kabel)
        btn_row = QHBoxLayout()
        self.btn_download_kabel = QPushButton("Download Data")
        self.btn_download_kabel.clicked.connect(self.download_kabel)
        self.btn_import_kabel = QPushButton("Import CSV ke Histori Material")
        self.btn_import_kabel.clicked.connect(self.import_kabel_csv)
        btn_row.addWidget(self.btn_import_kabel)
        btn_row.addWidget(self.btn_download_kabel)
        btn_row.addStretch()
        kabel_layout.addLayout(btn_row)

    def init_ont_tab(self):
        ont_layout = QVBoxLayout(self.tab_ont)
        search_row = QHBoxLayout()
        self.search_ont = QLineEdit()
        self.search_ont.setPlaceholderText("Cari serial number atau tim...")
        self.search_ont.textChanged.connect(self.filter_ont)
        search_row.addWidget(QLabel("🔍"))
        search_row.addWidget(self.search_ont)
        ont_layout.addLayout(search_row)
        self.tbl_ont = QTableWidget()
        self.tbl_ont.setColumnCount(6)
        self.tbl_ont.setHorizontalHeaderLabels(["No", "Tanggal", "Serial Number", "Nama Tim", "Status", "Aksi"])
        self.tbl_ont.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for i in range(1, 6):
            self.tbl_ont.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
        ont_layout.addWidget(self.tbl_ont)
        btn_row = QHBoxLayout()
        self.btn_import_ont = QPushButton("Import CSV ONT")
        self.btn_import_ont.clicked.connect(self.import_ont_csv_from_resume)
        self.btn_download_ont = QPushButton("Download Data")
        self.btn_download_ont.clicked.connect(self.download_ont)
        btn_row.addWidget(self.btn_import_ont)
        btn_row.addWidget(self.btn_download_ont)
        btn_row.addStretch()
        ont_layout.addLayout(btn_row)

    def show_stock(self):
        self.tbl_stock.setRowCount(0)
        # hitung jumlah diambil per item per tanggal
        taken_per_item = {}
        for entry in HISTORI_MAT:
            key = (entry['deskripsi'], entry['tanggal'][:10])
            taken_per_item[key] = taken_per_item.get(key, 0) + int(entry.get('qty', 0))

        for i, entry in enumerate(STOCK_ENTRIES):
            row = self.tbl_stock.rowCount()
            self.tbl_stock.insertRow(row)
            self.tbl_stock.setItem(row, 0, QTableWidgetItem(str(row + 1)))
            self.tbl_stock.item(row, 0).setTextAlignment(Qt.AlignCenter)
            self.tbl_stock.setItem(row, 1, QTableWidgetItem(entry["tanggal"]))
            self.tbl_stock.setItem(row, 2, QTableWidgetItem(entry["deskripsi"]))
            self.tbl_stock.setItem(row, 3, QTableWidgetItem(str(entry["qty"])))
            key = (entry["deskripsi"], entry["tanggal"][:10])
            diambil = taken_per_item.get(key, 0)
            self.tbl_stock.setItem(row, 4, QTableWidgetItem(str(diambil)))
            stock_awal = int(entry["qty"]) - diambil
            self.tbl_stock.setItem(row, 5, QTableWidgetItem(str(stock_awal)))
            btn_del = QPushButton("🗑")
            btn_del.setFixedSize(22, 22)
            btn_del.setStyleSheet("font-size:12px;min-width:20px;max-width:22px;background:#f44336;color:white;")
            btn_del.clicked.connect(lambda _, idx=i: self.hapus_stock(idx))
            self.tbl_stock.setCellWidget(row, 6, btn_del)

    def show_kabel(self, filter_txt=""):
        self.tbl_kabel.setRowCount(0)
        for i, entry in enumerate(HISTORI_MAT):
            if filter_txt and (filter_txt.lower() not in entry["deskripsi"].lower() and filter_txt.lower() not in entry["tim"].lower()):
                continue
            row = self.tbl_kabel.rowCount()
            self.tbl_kabel.insertRow(row)
            self.tbl_kabel.setItem(row, 0, QTableWidgetItem(str(row + 1)))
            self.tbl_kabel.setItem(row, 1, QTableWidgetItem(entry["tanggal"]))
            self.tbl_kabel.setItem(row, 2, QTableWidgetItem(entry["deskripsi"]))
            self.tbl_kabel.setItem(row, 3, QTableWidgetItem(str(entry["qty"])))
            self.tbl_kabel.setItem(row, 4, QTableWidgetItem(entry["tim"]))
            btn_del = QPushButton("🗑")
            btn_del.setFixedSize(22, 22)
            btn_del.setStyleSheet("font-size:12px;min-width:20px;max-width:22px;background:#f44336;color:white;")
            btn_del.clicked.connect(lambda _, idx=i: self.hapus_kabel(idx))
            self.tbl_kabel.setCellWidget(row, 5, btn_del)

    def show_ont(self, filter_txt=""):
        telegram_sn_set = set()
        for tab in self.laporan_tabs:
            telegram_sn_set.update(tab.get_all_sn())
        self.tbl_ont.setRowCount(0)
        for i, entry in enumerate(HISTORI_ONT):
            if filter_txt and (filter_txt.lower() not in entry["sn"].lower() and filter_txt.lower() not in entry["tim"].lower()):
                continue
            row = self.tbl_ont.rowCount()
            self.tbl_ont.insertRow(row)
            self.tbl_ont.setItem(row, 0, QTableWidgetItem(str(row + 1)))
            self.tbl_ont.setItem(row, 1, QTableWidgetItem(entry["tanggal"]))
            self.tbl_ont.setItem(row, 2, QTableWidgetItem(entry["sn"]))
            self.tbl_ont.setItem(row, 3, QTableWidgetItem(entry["tim"]))
            status = "Terpakai" if entry["sn"] in telegram_sn_set else "Kosong"
            status_item = QTableWidgetItem(status)
            # warna status: hijau untuk terpakai, merah untuk kosong
            from PySide6.QtGui import QColor
            status_item.setForeground(QColor("green") if status == "Terpakai" else QColor("red"))
            self.tbl_ont.setItem(row, 4, status_item)
            btn_del = QPushButton("🗑")
            btn_del.setFixedSize(22, 22)
            btn_del.setStyleSheet("font-size:12px;min-width:20px;max-width:22px;background:#f44336;color:white;")
            btn_del.clicked.connect(lambda _, idx=i: self.hapus_ont(idx))
            self.tbl_ont.setCellWidget(row, 5, btn_del)

    def filter_kabel(self):
        txt = self.search_kabel.text()
        self.show_kabel(txt)

    def filter_ont(self):
        txt = self.search_ont.text()
        self.show_ont(txt)

    def hapus_kabel(self, idx):
        global HISTORI_MAT
        if 0 <= idx < len(HISTORI_MAT):
            del HISTORI_MAT[idx]
            save_json("histori_kabel_aksesori.json", HISTORI_MAT)
            self.reload_data()

    def hapus_ont(self, idx):
        global HISTORI_ONT
        if 0 <= idx < len(HISTORI_ONT):
            del HISTORI_ONT[idx]
            save_json("histori_ont.json", HISTORI_ONT)
            self.reload_data()

    def add_stock(self):
        desc = self.stock_desc.currentText().strip()
        qty = self.stock_qty.value()
        tgl = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if not desc or qty <= 0:
            QMessageBox.warning(self, "Validasi", "Isi nama item dan qty dengan benar.")
            return
        global STOCK_ENTRIES
        STOCK_ENTRIES.append({
            "tanggal": tgl,
            "deskripsi": desc,
            "qty": qty
        })
        save_json("stock_entries.json", STOCK_ENTRIES)
        # jika item baru, tambahkan ke MATERIAL/AKSESORI (sederhana: tambahkan ke MATERIAL)
        if desc not in MATERIAL and desc not in AKSESORI:
            MATERIAL.append(desc)
            save_json("material.json", MATERIAL)
        self.show_stock()
        self.stock_qty.setValue(1)
        QMessageBox.information(self, "Berhasil", "Stock berhasil ditambahkan.")

    def hapus_stock(self, idx):
        global STOCK_ENTRIES
        if 0 <= idx < len(STOCK_ENTRIES):
            del STOCK_ENTRIES[idx]
            save_json("stock_entries.json", STOCK_ENTRIES)
            self.show_stock()

    def download_kabel(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, "Simpan CSV Kabel/Aksesori",
            f"rekap_kabel_aksesori_{datetime.now().strftime('%Y%m%d')}.csv",
            "CSV Files (*.csv)"
        )
        if filename:
            try:
                with open(filename, "w", newline='', encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Tanggal", "Deskripsi", "Qty", "Nama Tim"])
                    for entry in HISTORI_MAT:
                        writer.writerow([entry["tanggal"], entry["deskripsi"], entry["qty"], entry.get("tim","")])
                QMessageBox.information(self, "Download Berhasil", f"Berhasil mengunduh ke {filename}")
            except Exception as e:
                QMessageBox.warning(self, "Gagal", f"Gagal menyimpan CSV:\n{e}")

    def download_ont(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, "Simpan CSV SN ONT",
            f"rekap_ont_{datetime.now().strftime('%Y%m%d')}.csv",
            "CSV Files (*.csv)"
        )
        if filename:
            try:
                with open(filename, "w", newline='', encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Tanggal", "Serial Number", "Nama Tim", "Status"])
                    telegram_sn_set = set()
                    for tab in self.laporan_tabs:
                        telegram_sn_set.update(tab.get_all_sn())
                    for entry in HISTORI_ONT:
                        status = "Terpakai" if entry["sn"] in telegram_sn_set else "Kosong"
                        writer.writerow([entry["tanggal"], entry["sn"], entry.get("tim",""), status])
                QMessageBox.information(self, "Download Berhasil", f"Berhasil mengunduh ke {filename}")
            except Exception as e:
                QMessageBox.warning(self, "Gagal", f"Gagal menyimpan CSV:\n{e}")

    def import_kabel_csv(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Pilih file CSV Material untuk diimport", os.path.expanduser("~"), "CSV Files (*.csv);;All Files (*)")
        if not filename:
            return
        try:
            added = 0
            with open(filename, newline='', encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # cari field yang mungkin mengindikasikan material
                    desc = row.get("Deskripsi") or row.get("deskripsi") or row.get("Nama Item") or row.get("Item") or ""
                    qty = row.get("Qty") or row.get("qty") or row.get("Jumlah") or row.get("jumlah") or "0"
                    tim = row.get("Tim") or row.get("tim") or row.get("Nama Tim") or LAST_SELECTION.get("tim","")
                    if not desc:
                        # jika tidak dapat menemukan deskripsi, skip
                        continue
                    try:
                        q = int(float(qty))
                    except Exception:
                        q = 0
                    if q <= 0:
                        # jika qty nol, skip
                        continue
                    entry = {
                        "tanggal": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "deskripsi": desc.strip(),
                        "qty": q,
                        "tim": tim,
                        "divisi": LAST_SELECTION.get("divisi","")
                    }
                    HISTORI_MAT.append(entry)
                    added += 1
                    # tambahkan ke master list jika belum ada
                    if desc.strip() not in MATERIAL and desc.strip() not in AKSESORI:
                        MATERIAL.append(desc.strip())
                # selesai baca
            if added > 0:
                save_json("histori_kabel_aksesori.json", HISTORI_MAT)
                save_json("material.json", MATERIAL)
                QMessageBox.information(self, "Import Selesai", f"Berhasil menambahkan {added} entri material dari {filename}")
                self.reload_data()
            else:
                QMessageBox.information(self, "Import", "Tidak ada entri material valid ditemukan di file.")
        except Exception as e:
            QMessageBox.warning(self, "Gagal Import", f"Gagal mengimpor CSV:\n{e}")

    def import_ont_csv_from_resume(self):
        # alias untuk import ONT dari tab Resume
        filename, _ = QFileDialog.getOpenFileName(self, "Pilih file CSV SN ONT untuk diimport", os.path.expanduser("~"), "CSV Files (*.csv);;All Files (*)")
        if not filename:
            return
        try:
            added = 0
            with open(filename, newline='', encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    sn = row.get("SN") or row.get("Serial Number") or row.get("sn") or row.get("serial_number") or ""
                    if not sn:
                        continue
                    entry = {
                        "tanggal": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "sn": str(sn).strip(),
                        "tim": row.get("Tim") or row.get("tim") or LAST_SELECTION.get("tim",""),
                        "divisi": row.get("Divisi") or row.get("divisi") or LAST_SELECTION.get("divisi","")
                    }
                    HISTORI_ONT.append(entry)
                    added += 1
            if added > 0:
                save_json("histori_ont.json", HISTORI_ONT)
                QMessageBox.information(self, "Import Selesai", f"Berhasil menambahkan {added} SN dari {filename}")
                self.reload_data()
            else:
                QMessageBox.information(self, "Import", "Tidak ada SN ditemukan di file.")
        except Exception as e:
            QMessageBox.warning(self, "Gagal Import", f"Gagal mengimpor CSV:\n{e}")

    def reload_data(self):
        # refresh semua tampilan tabel
        self.show_stock()
        self.show_kabel(self.search_kabel.text())
        self.show_ont(self.search_ont.text())

class MaterialTracker(QMainWindow):
    def __init__(self):
        super().__init__()

        icon_path = os.path.join(os.path.dirname(__file__), "icon", "material_tracker.png")
        if os.path.isfile(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            self.setWindowIcon(QIcon.fromTheme("applications-system"))

        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.setMinimumSize(940, 620)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Inisialisasi laporan tabs (MyRepublic, Asianet, IKR Oxygen)
        self.laporan_tabs = [
            TelegramReportTab('myrepublic'),
            TelegramReportTab('asianet'),
            TelegramReportTab('oxygen')
        ]

        self.form_pengambilan = FormPengambilan(self)
        self.resume = Resume(self, self.laporan_tabs)
        self.laporan_tab_wrapper = LaporanTab(self.laporan_tabs)

        self.tabs.addTab(self.form_pengambilan, "Pengambilan")
        self.tabs.addTab(self.resume, "Resume")
        self.tabs.addTab(self.laporan_tab_wrapper, "Laporan")

        self.init_menu()

    def init_menu(self):
        menubar = self.menuBar()
        pref_menu = QMenu("&Preferences", self)
        self.action_light = QAction("Tema Terang", self, checkable=True)
        self.action_dark = QAction("Tema Gelap", self, checkable=True)
        self.action_light.setChecked(True)
        self.action_dark.triggered.connect(lambda: self.set_theme(dark=True))
        self.action_light.triggered.connect(lambda: self.set_theme(dark=False))
        pref_menu.addAction(self.action_light)
        pref_menu.addAction(self.action_dark)

        self.action_settings = QAction("Pengaturan...", self)
        self.action_settings.triggered.connect(self.show_settings)
        pref_menu.addSeparator()
        pref_menu.addAction(self.action_settings)
        menubar.addMenu(pref_menu)

        help_menu = QMenu("&Help", self)
        about_action = QAction("Tentang", self)
        about_action.triggered.connect(self.about)
        help_action = QAction("Bantuan Online", self)
        help_action.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/endans/material-tracker")))
        help_menu.addAction(about_action)
        help_menu.addAction(help_action)
        menubar.addMenu(help_menu)

    def set_theme(self, dark):
        self.action_dark.setChecked(dark)
        self.action_light.setChecked(not dark)
        apply_theme(QApplication.instance(), dark=dark)

    def about(self):
        QMessageBox.information(
            self, "Tentang",
            f"{APP_NAME} v{VERSION}\n\nAplikasi pelacak material untuk tim CKT Purwokerto.\nDikembangkan - fitur: pencarian cepat, simpan pilihan terakhir, import CSV, dan laporan Oxygen."
        )

    def show_settings(self):
        dlg = SettingsDialog(self)
        dlg.reload()
        if dlg.exec():
            self.reload_all()

    def reload_all(self):
        # reload opsi form pengambilan & resume data
        self.form_pengambilan.reload_options()
        self.resume.reload_data()
        for tab in self.laporan_tabs:
            tab.load_reports()

class LaporanTab(QWidget):
    def __init__(self, laporan_tabs):
        super().__init__()
        self.tabs = QTabWidget()
        self.laporan_tabs = laporan_tabs
        for tab in self.laporan_tabs:
            self.tabs.addTab(tab, tab.display_name)
        layout = QVBoxLayout(self)
        layout.addWidget(self.tabs)

def main():
    app = QApplication(sys.argv)
    apply_theme(app, dark=False)
    mw = MaterialTracker()
    mw.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
