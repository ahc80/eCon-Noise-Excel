import sys
import os
import glob
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QHBoxLayout

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
        self.setWindowTitle('CSV File Merger')
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
                combinedDf = pd.concat([combinedDf, df], ignore_index=True)
            except Exception as e:
                print(f"Could not read {file}: {e}")

        return combinedDf

    def saveData(self, folderPath, data):
        output_path = os.path.join(folderPath, 'combinedCSV.csv')
        data.to_csv(output_path, index=False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FileSelector()
    ex.show()
    sys.exit(app.exec())
