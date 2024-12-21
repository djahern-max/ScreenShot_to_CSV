import sys
import json
import os
import platform
import subprocess
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QTableWidget, QTableWidgetItem, QHeaderView,
                           QPushButton, QLabel, QFileDialog, QMessageBox, 
                           QInputDialog, QDialog)
from PyQt6.QtCore import Qt
from PIL import Image
import pytesseract
from datetime import datetime
import time
import re

class ErrorCorrectionDialog(QDialog):
    def __init__(self, expenses, parent=None):
        super().__init__(parent)
        self.expenses = expenses.copy()
        self.setWindowTitle("Error Correction")
        self.setModal(True)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        instruction_label = QLabel(
            "Please review and correct any errors in the amounts below. "
            "Common issues:\n- Missing digits (e.g., $1.34 instead of $51.34)\n"
            "- Incorrect decimal places\n- OCR misreading numbers"
        )
        instruction_label.setWordWrap(True)
        layout.addWidget(instruction_label)
        
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Entry #", "Amount", "Remark"])
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        
        self.table.setRowCount(len(self.expenses))
        for i, expense in enumerate(self.expenses):
            entry_item = QTableWidgetItem(str(i + 1))
            entry_item.setFlags(entry_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, entry_item)
            
            amount_item = QTableWidgetItem(f"{expense['amount']:.2f}")
            self.table.setItem(i, 1, amount_item)
            
            remark_item = QTableWidgetItem(expense['remark'])
            remark_item.setFlags(remark_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 2, remark_item)
        
        layout.addWidget(self.table)
        
        button_layout = QHBoxLayout()
        save_button = QPushButton("Save Corrections")
        save_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
    def get_corrected_expenses(self):
        corrected_expenses = []
        for i in range(self.table.rowCount()):
            try:
                amount = float(self.table.item(i, 1).text())
                remark = self.table.item(i, 2).text()
                corrected_expenses.append({
                    "amount": round(amount, 2),
                    "remark": remark
                })
            except (ValueError, AttributeError):
                continue
        return corrected_expenses

class FinancialScreenshotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize Tesseract path
        if platform.system().lower() == 'darwin':  # Mac
            pytesseract.pytesseract.tesseract_cmd = '/opt/homebrew/bin/tesseract'
        else:  # Windows - adjust path as needed
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        
        # Initialize main window properties
        self.setWindowTitle("Expense Screenshot to JSON Converter")
        self.setGeometry(100, 100, 500, 400)
        
        # Initialize variables
        self.expenses = []
        self.total_processed = 0
        
        # Create and set up the UI
        self.create_ui()
    
    def create_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Instructions label
        instructions = QLabel(
            "Instructions:\n\n"
            "1. Click 'Capture Region'\n"
            "2. Click 'OK' when prompted to start the capture process\n"
            "3. Select the region containing Amount and Remark columns\n"
            "4. Verify the captured data\n"
            "5. Repeat if more data needs to be captured\n"
            "6. Save to JSON when finished"
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)
        
        # Buttons
        self.capture_btn = QPushButton("Capture Region")
        self.capture_btn.clicked.connect(self.capture_screenshot)
        layout.addWidget(self.capture_btn)
        
        self.save_btn = QPushButton("Save to JSON")
        self.save_btn.clicked.connect(self.save_json)
        layout.addWidget(self.save_btn)
        
        # Status labels
        self.counter_label = QLabel("Processed expenses: 0")
        layout.addWidget(self.counter_label)
        
        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

    def process_text_to_json(self, text):
        """Convert OCR text to structured JSON data with improved parsing"""
        try:
            # Split text into lines and filter empty lines
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            new_expenses = []
            amount_pattern = r'^(\d+\.?\d*)'  # More strict pattern for positive numbers with optional decimals
            
            for line in lines:
                # Split line into components
                parts = line.split()
                if not parts:
                    continue
                
                # Try to find amount
                amount_match = re.match(amount_pattern, parts[0])
                if amount_match:
                    try:
                        amount = float(amount_match.group().replace(',', ''))
                        
                        # Look for remark at the end of the line
                        remark = "No Remark Available"
                        
                        # Check if there's a potential remark at the end
                        last_part = parts[-1]
                        if last_part.isdigit() and len(last_part) <= 6:
                            remark = last_part
                            
                        new_expenses.append({
                            "amount": round(amount, 2),
                            "remark": remark
                        })
                    except ValueError:
                        continue
            
            return new_expenses
            
        except Exception as e:
            print(f"Error processing text: {str(e)}")
            return []

    def capture_screenshot(self):
        filename = None
        try:
            # Show processing cursor
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            
            self.showMinimized()
            QApplication.processEvents()
            time.sleep(0.5)
            
            # Inform user
            QMessageBox.information(
                self, 
                "Region Selection",
                "Click OK to start the screenshot capture process.\n\n"
                "After clicking OK, select the region containing both Amount and Remark columns."
            )
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"expense_capture_{timestamp}.png"
            
            if platform.system().lower() == 'darwin':  # Mac
                try:
                    result = subprocess.run(
                        ['screencapture', '-i', filename],
                        check=True,
                        timeout=30,  # 30 second timeout
                        capture_output=True
                    )
                except subprocess.TimeoutExpired:
                    raise Exception("Screenshot capture timed out. Please try again.")
            else:  # Windows
                # ... existing Windows code ...
                pass

            if os.path.exists(filename) and os.path.getsize(filename) > 0:
                try:
                    with Image.open(filename) as img:
                        text = pytesseract.image_to_string(img, config='--psm 6', timeout=10)
                    new_expenses = self.process_text_to_json(text)
                    
                    if new_expenses:
                        dialog = ErrorCorrectionDialog(new_expenses, self)
                        if dialog.exec() == QDialog.DialogCode.Accepted:
                            corrected_expenses = dialog.get_corrected_expenses()
                            self.expenses.extend(corrected_expenses)
                            self.total_processed += len(corrected_expenses)
                            self.counter_label.setText(f"Processed expenses: {self.total_processed}")
                    else:
                        QMessageBox.warning(self, "Warning", 
                                          "No valid expense data found in selection.\n"
                                          "Make sure Amount and Remark columns are clearly visible.")
                finally:
                    if os.path.exists(filename):
                        os.remove(filename)
            else:
                QMessageBox.warning(self, "Warning", "No region was selected or capture failed.")
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to capture/process region: {str(e)}")
        
        finally:
            # Cleanup
            try:
                if filename and os.path.exists(filename):
                    os.remove(filename)
                self.showNormal()
                QApplication.restoreOverrideCursor()
                QApplication.processEvents()
            except Exception as cleanup_error:
                print(f"Cleanup error: {str(cleanup_error)}")

    def save_json(self):
        if not self.expenses:
            QMessageBox.warning(self, "Warning", "No expense data to save.")
            return

        total = sum(expense["amount"] for expense in self.expenses)
        
        msg = QMessageBox()
        msg.setWindowTitle("Verify Total")
        msg.setText(f"Total amount: ${total:.2f}\n\nIs this total correct?")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | 
                             QMessageBox.StandardButton.No)
        msg.setDefaultButton(QMessageBox.StandardButton.Yes)
        
        if msg.exec() == QMessageBox.StandardButton.No:
            QMessageBox.information(self, "Save Cancelled", 
                                  "Please review and recapture the data if needed.")
            return
            
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Save JSON File",
            f"expenses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            "JSON Files (*.json)"
        )
        
        if not file_name:
            return
            
        try:
            final_json = {
                "expenses": sorted(self.expenses, key=lambda x: x['amount'])
            }
            
            with open(file_name, 'w') as f:
                json.dump(final_json, f, indent=1)
                
            QMessageBox.information(
                self,
                "Success", 
                f"Successfully saved {len(self.expenses)} expenses to JSON!\n"
                f"File: {file_name}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save JSON: {str(e)}")

    def cleanup(self):
        try:
            self.showNormal()
            QApplication.processEvents()
            QApplication.restoreOverrideCursor()
        except Exception as e:
            print(f"Cleanup error: {str(e)}")


def main():
    app = QApplication(sys.argv)
    window = FinancialScreenshotApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()