import sys
import os
import glob
import numpy as np
import pandas as pd
import xlsxwriter
import datetime
import matplotlib.pyplot as plt
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QHBoxLayout
import pyqtgraph as pg
from PyQt6.QtWidgets import QVBoxLayout
from pyqtgraph.Qt import QtGui

class FileSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        mainLayout = QVBoxLayout(self)

        self.selectButton = QPushButton('Select Folder with CSV Files', self)
        self.selectButton.clicked.connect(self.showDialog)

        self.filesLabel = QLabel('No folder selected', self)
        
        buttonLayout = QHBoxLayout()
        buttonLayout.addWidget(self.selectButton)
        
        labelLayout = QHBoxLayout()
        labelLayout.addWidget(self.filesLabel)

        mainLayout.addLayout(buttonLayout)
        mainLayout.addLayout(labelLayout)

        self.setLayout(mainLayout)
        self.setWindowTitle('CSV File Merger and Plotter')
        self.setGeometry(500, 500, 400, 300)

    def showDialog(self):
        options = QFileDialog.Option.ShowDirsOnly
        folderPath = QFileDialog.getExistingDirectory(self, 'Select Folder', options=options)
        if folderPath:
            self.filesLabel.setText('Processing...')
            data = self.combineCSVFiles(folderPath)
            if not data.empty:
                self.saveData(folderPath, data)
                self.filesLabel.setText('Processing Complete! Check the folder for results and combined CSV file.')
            else:
                self.filesLabel.setText('No CSV files found or combined data is empty.')

    def combineCSVFiles(self, folderPath):
        filePattern = os.path.join(folderPath, '*.csv')
        csvFiles = glob.glob(filePattern)
        combinedDf = pd.DataFrame()

        for file in csvFiles:
            try:
                # Use header=None to prevent automatic header detection, then skiprows to start from the correct data row
                df = pd.read_csv(file, header=None, skiprows=6)  
                df.columns = df.iloc[0]  # Set the column headers from the first row (index 0 after skiprows)
                df = df[1:]  # Remove the header row from the data
                
                # Preserve 'Date/Time' as datetime, convert other columns to numeric
                df['Date/Time'] = pd.to_datetime(df['Date/Time'], errors='coerce')  # Ensure Date/Time is converted
                numeric_cols = df.columns.drop('Date/Time')  # All columns except Date/Time
                df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce')  # Convert numeric columns
            
                combinedDf = pd.concat([combinedDf, df], ignore_index=True)           
            except Exception as e:
                print(f"Could not read {file}: {e}")

        return combinedDf

    def saveData(self, folderPath, data):
        outputPath = os.path.join(folderPath, 'combinedCSV.xlsx')
        writer = pd.ExcelWriter(outputPath, engine='xlsxwriter')
        data.to_excel(writer, sheet_name='Data', index=False)

        # Creating a simple macro
        vba_code = """
    Sub Auto_Open()
        MsgBox "Welcome to your combined CSV data!", vbInformation
    End Sub
        """

        # Adding the macro to the workbook
        writer.book.filename = outputPath
        writer.book.add_vba_project()
        writer.book.add_vba_code(vba_code)
        
        writer.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FileSelector()
    ex.show()
    sys.exit(app.exec())
