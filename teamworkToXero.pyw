import sys
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QVBoxLayout, QComboBox, QMessageBox

class InvoiceGenerator(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Select CSV File
        self.csv_label = QLabel('Select Teamwork CSV File', self)
        layout.addWidget(self.csv_label)

        self.csv_input = QLineEdit(self)
        layout.addWidget(self.csv_input)

        self.csv_button = QPushButton('Browse', self)
        self.csv_button.clicked.connect(self.browseFile)
        layout.addWidget(self.csv_button)

        # Invoice Number
        self.invoice_label = QLabel('Enter Invoice Number', self)
        layout.addWidget(self.invoice_label)

        self.invoice_input = QLineEdit(self)
        layout.addWidget(self.invoice_input)

        # Unit Amount
        self.unit_label = QLabel('Enter Unit Amount', self)
        layout.addWidget(self.unit_label)

        self.unit_input = QLineEdit(self)
        layout.addWidget(self.unit_input)

        # Currency Selection
        self.currency_label = QLabel('Select Currency', self)
        layout.addWidget(self.currency_label)

        self.currency_combo = QComboBox(self)
        self.currency_combo.addItems(['AUD', 'NZD'])
        layout.addWidget(self.currency_combo)

        # Aggregation Type
        self.aggregate_label = QLabel('Aggregate by', self)
        layout.addWidget(self.aggregate_label)

        self.aggregate_combo = QComboBox(self)
        self.aggregate_combo.addItems(['Task', 'Date'])
        layout.addWidget(self.aggregate_combo)

        # Generate Button
        self.generate_button = QPushButton('Generate Invoice', self)
        self.generate_button.clicked.connect(self.generateInvoice)
        layout.addWidget(self.generate_button)

        # Set layout and window properties
        self.setLayout(layout)
        self.setWindowTitle('Teamwork to Xero Invoice Generator')
        self.setGeometry(300, 300, 400, 250)

    def browseFile(self):
        # Open file dialog to select CSV file
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Select Teamwork CSV File", "", "CSV Files (*.csv)", options=options)
        if fileName:
            self.csv_input.setText(fileName)

    def generateInvoice(self):
        csv_file = self.csv_input.text()
        invoice_number = self.invoice_input.text()
        unit_amount = self.unit_input.text()
        currency = self.currency_combo.currentText()
        aggregate_by = self.aggregate_combo.currentText()

        # Ensure all fields are filled
        if not csv_file or not invoice_number or not unit_amount:
            QMessageBox.warning(self, 'Input Error', 'Please fill in all fields!')
            return

        # Process the CSV
        try:
            df = pd.read_csv(csv_file, index_col=False)
            df['Decimal Hours'] = df['Hours'] + df['Minutes'] / 60  # Convert time to decimal hours

            # Handle missing task and description, label "Other" when both are missing
            df['Task or Description'] = df.apply(
                lambda row: row['Task'] if pd.notnull(row['Task']) and row['Task'] else
                            row['Description'] if pd.notnull(row['Description']) and row['Description'] else
                            'Other', axis=1)

            if aggregate_by == 'Task':
                # Group by task
                grouped = df.groupby('Task or Description').agg({
                    'Company': 'first',  # Assuming all rows within a task have the same company
                    'Decimal Hours': 'sum'  # Sum the hours worked per task
                }).reset_index()

                # Prepare the Xero invoice format (grouped by task)
                xero_df = pd.DataFrame({
                    '*ContactName': grouped['Company'],
                    'EmailAddress': '',
                    'POAddressLine1': '',
                    'POAddressLine2': '',
                    'POAddressLine3': '',
                    'POAddressLine4': '',
                    'POCity': '',
                    'PORegion': '',
                    'POPostalCode': '',
                    'POCountry': '',
                    '*InvoiceNumber': invoice_number,  # Use the user-provided invoice number
                    'Reference': '',
                    '*InvoiceDate': datetime.now().strftime('%d/%m/%Y'),
                    '*DueDate': (datetime.now().replace(day=20) if datetime.now().day < 20 else datetime.now().replace(day=20, month=datetime.now().month + 1)).strftime('%d/%m/%Y'),
                    'InventoryItemCode': '',
                    '*Description': grouped['Task or Description'],
                    '*Quantity': grouped['Decimal Hours'],
                    '*UnitAmount': unit_amount,
                    'Discount': '',
                    '*AccountCode': '200',
                    '*TaxType': 'NONE',
                    'TrackingName1': '',
                    'TrackingOption1': '',
                    'TrackingName2': '',
                    'TrackingOption2': '',
                    'Currency': currency,
                    'BrandingTheme': ''
                })
            else:
                # Group by date and concatenate tasks under the same date
                df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
                df['Date Label'] = df['Date'].apply(lambda x: x.strftime('%d %B %Y'))  # Format as "DD MONTH YEAR"
                
                # Group by date, and concatenate tasks on the same date
                grouped = df.groupby('Date Label').agg({
                    'Company': 'first',  # Assuming all rows within a date have the same company
                    'Decimal Hours': 'sum',  # Sum the hours worked per date
                    'Task or Description': lambda x: '\n'.join(x)  # Concatenate tasks for the same date
                }).reset_index()

                # Prepare the Xero invoice format (grouped by date)
                xero_df = pd.DataFrame({
                    '*ContactName': grouped['Company'],
                    'EmailAddress': '',
                    'POAddressLine1': '',
                    'POAddressLine2': '',
                    'POAddressLine3': '',
                    'POAddressLine4': '',
                    'POCity': '',
                    'PORegion': '',
                    'POPostalCode': '',
                    'POCountry': '',
                    '*InvoiceNumber': invoice_number,  # Use the user-provided invoice number
                    'Reference': '',
                    '*InvoiceDate': datetime.now().strftime('%d/%m/%Y'),
                    '*DueDate': (datetime.now().replace(day=20) if datetime.now().day < 20 else datetime.now().replace(day=20, month=datetime.now().month + 1)).strftime('%d/%m/%Y'),
                    'InventoryItemCode': '',
                    '*Description': grouped['Date Label'] + ':\n' + grouped['Task or Description'],  # Prepend date to concatenated tasks
                    '*Quantity': grouped['Decimal Hours'],
                    '*UnitAmount': unit_amount,
                    'Discount': '',
                    '*AccountCode': '200',
                    '*TaxType': 'NONE',
                    'TrackingName1': '',
                    'TrackingOption1': '',
                    'TrackingName2': '',
                    'TrackingOption2': '',
                    'Currency': currency,
                    'BrandingTheme': ''
                })

            # Save the result to CSV
            xero_df.to_csv('xero_invoice.csv', index=False)
            QMessageBox.information(self, 'Success', 'Invoice CSV created successfully!')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error processing file: {e}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = InvoiceGenerator()
    ex.show()
    sys.exit(app.exec_())
