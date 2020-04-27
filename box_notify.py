import os.path
import pickle
import string
import sys

import simpleaudio as sa
from PySide2.QtCore import (QTimer, QSettings)
from PySide2.QtWidgets import (QPushButton, QApplication,
                               QVBoxLayout, QHBoxLayout, QLineEdit, QLabel, QDialog, QTableWidget, QTableWidgetItem,
                               QAbstractItemView, QFileDialog)

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '12goY1Z2Qn5zpWkjbUpGgxusp8stgiKksib96rB24_X4'
SAMPLE_RANGE_NAME = 'App Sales'


class SettingsForm(QDialog):
    def __init__(self, parent=None):
        super(SettingsForm, self).__init__(parent)
        self.settings = QSettings()

        layout = QVBoxLayout()

        # Notification sound
        wav_file = self.settings.value("notification_sound", "when.wav")
        self.sound_lbl = QLabel("Notification sound:")
        self.sound_val = QLineEdit(wav_file)
        self.sound_btn = QPushButton("Browse")
        hl = QHBoxLayout()
        hl.addWidget(self.sound_lbl)
        hl.addWidget(self.sound_val)
        hl.addWidget(self.sound_btn)
        layout.addLayout(hl)

        # Spread sheet ID
        sheet_id = self.settings.value("sheet_id", SAMPLE_SPREADSHEET_ID)
        self.sheet_id_lbl = QLabel("Spread sheet ID: ")
        self.sheet_id_val = QLineEdit(sheet_id)
        hl = QHBoxLayout()
        hl.addWidget(self.sheet_id_lbl)
        hl.addWidget(self.sheet_id_val)
        layout.addLayout(hl)

        # Control buttons
        self.save_btn = QPushButton("Save")
        self.cancel_btn = QPushButton("Cancel")
        hl = QHBoxLayout()
        hl.addWidget(self.save_btn)
        hl.addWidget(self.cancel_btn)
        layout.addLayout(hl)

        self.setLayout(layout)

        self.sound_btn.clicked.connect(self.sound_browse)
        self.save_btn.clicked.connect(self.on_save)
        self.cancel_btn.clicked.connect(self.on_cancel)

    def sound_browse(self):
        filename, filter = QFileDialog.getOpenFileName(parent=self, caption='Select sound file for notifications',
                                                       dir='.',
                                                       filter='Waveform Audio File Format (*.wav)')
        if filename:
            self.sound_val.setText(filename)

    def on_save(self):
        self.settings.setValue("notification_sound", self.sound_val.text())
        self.settings.setValue("sheet_id", self.sheet_id_val.text())
        self.hide()

    def on_cancel(self):
        self.hide()


class Form(QDialog):
    creds: None
    service: None
    sheet: None

    def __init__(self, parent=None):
        super(Form, self).__init__(parent)
        self.settings = QSettings()

        # Create widgets
        self.settings_btn = QPushButton("Settings")
        self.next = QPushButton("Next")
        self.table = QTableWidget(0, 5)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers);
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setHorizontalHeaderLabels(["Name", "Item", "Amount", "Time", "Kitchen"])
        # self.table.resizeColumnsToContents()

        # Create layout and add widgets
        layout = QVBoxLayout()
        layout.addWidget(self.table)

        h = QHBoxLayout()
        h.addWidget(self.settings_btn)
        h.addWidget(self.next)
        layout.addLayout(h)
        self.setLayout(layout)

        self.settings_btn.clicked.connect(self.show_settings_form)

        self.table.itemSelectionChanged.connect(self.greetings)
        self.orders = list()
        self.old_val = list()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_google_sheet)

        while True:
            try:
                self.setup_google_access()
                self.timer.start(5000)
                break
            except:
                settigns_form = SettingsForm(self)
                settigns_form.exec_()

    def setup_google_access(self):
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('sheets', 'v4', credentials=self.creds)

        # Call the Sheets API
        self.sheet = self.service.spreadsheets()
        id = self.settings.value("sheet_id", SAMPLE_SPREADSHEET_ID)

        result = self.sheet.values().get(spreadsheetId=id,
                                         range=SAMPLE_RANGE_NAME).execute()

        self.old_val = result.get('values', [])

    def check_google_sheet(self):
        id = self.settings.value("sheet_id", SAMPLE_SPREADSHEET_ID)
        result = self.sheet.values().get(spreadsheetId=id,
                                         range=SAMPLE_RANGE_NAME).execute()
        new_val = result.get('values', [])

        if len(new_val) != len(self.old_val):
            d = new_val[len(self.old_val):]
            for item in d:

                order = dict(purchaser_name="undefined",
                             purchaser_email="undefined",
                             purchased_item="undefined",
                             item_sku="undefined",
                             amount="undefined",
                             paid_at="undefined",
                             shipping_city="undefined",
                             shipping_country="undefined",
                             shipping_address_line_1="undefined",
                             shipping_address_line_2="undefined",
                             shipping_postal_code="undefined",
                             shipping_state="undefined",
                             processor_link="undefined",
                             paid_time="undefined",
                             kitchen_time="undefined",
                             )

                try:
                    cell_value = lambda x: item[string.ascii_uppercase.index(x)]

                    order["purchaser_name"] = cell_value('A')
                    order["purchaser_email"] = cell_value('B')
                    order["purchased_item"] = cell_value('C')
                    order["item_sku"] = cell_value('D')
                    order["amount"] = cell_value('E')
                    order["paid_at"] = cell_value('F')
                    order["shipping_city"] = cell_value('G')
                    order["shipping_country"] = cell_value('H')
                    order["shipping_address_line_1"] = cell_value('I')
                    order["shipping_address_line_2"] = cell_value('J')
                    order["shipping_postal_code"] = cell_value('K')
                    order["shipping_state"] = cell_value('L')
                    order["processor_link"] = cell_value('M')
                    order["paid_time"] = cell_value('N')
                    order["kitchen_time"] = cell_value('O')
                except IndexError:
                    print('OMG...')

                self.orders.append(order)
                table_idx = self.table.rowCount()

                self.table.insertRow(table_idx)

                self.table.setItem(table_idx, 0, QTableWidgetItem(order["purchaser_name"]))
                self.table.setItem(table_idx, 1, QTableWidgetItem(order["purchased_item"]))
                self.table.setItem(table_idx, 2, QTableWidgetItem(order["amount"]))
                self.table.setItem(table_idx, 3, QTableWidgetItem(order["paid_time"]))
                self.table.setItem(table_idx, 3, QTableWidgetItem(order["kitchen_time"]))

            wav_file = self.settings.value("notification_sound", "when.wav")
            wave_obj = sa.WaveObject.from_wave_file(wav_file)
            play_obj = wave_obj.play()
            self.old_val = new_val

    # Greets the user
    def greetings(self):
        print("Hello")

    def show_settings_form(self):
        settigns_form = SettingsForm(self)
        settigns_form.exec_()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setOrganizationName("BurgerBox");
    app.setOrganizationDomain("io32.com");
    app.setApplicationName("BurgerBox notifier");

    form = Form()
    form.resize(600, 300)
    form.show()
    sys.exit(app.exec_())
