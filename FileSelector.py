import sys
import os
import glob
import numpy as np
import pandas as pd
import xlsxwriter
import datetime
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

        fig, ax = plt.subplots()
        colormap = plt.get_cmap('viridis')
        colors = colormap(np.linspace(0, 1, len(data.columns) - 1))

        # Plot all AI columns on one graph
        for idx, column in enumerate([col for col in data.columns if 'AI' in col]):
            data.plot(x='Date/Time', y=column, ax=ax, color=colors[idx], label=column)

        ax.set_title('Combined Plot')
        ax.set_ylabel('mOhms')
        ax.set_xlabel('Date/Time')
        ax.legend(title='Legend')

        ax.set_ylim(0, 12)

        dateFormat = "%Y-%m-%d %H:%M:%S"  # Adjust this format to match your 'Date/Time' column
        startDate = datetime.datetime.strptime('2024-01-01 00:00:00', dateFormat)
        endDate = datetime.datetime.strptime('2024-12-31 23:59:59', dateFormat)
        ax.set_xlim(startDate, endDate)
        ax.xaxis.set_major_formatter(plt.FixedFormatter(dateFormat))

        image_path = os.path.join(folderPath, 'combined_plot.png')
        plt.savefig(image_path)
        plt.close()

        # Create a worksheet for the combined plot
        worksheet = writer.book.add_worksheet(name='Combined Plot')
        worksheet.insert_image('B2', image_path)

        writer.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FileSelector()
    ex.show()
    sys.exit(app.exec())
