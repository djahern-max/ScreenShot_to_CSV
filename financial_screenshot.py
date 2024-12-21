import sys
import json
import os
import platform
import subprocess
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QPushButton,
    QLabel,
    QFileDialog,
    QMessageBox,
    QInputDialog,
    QDialog,
    QLineEdit,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from PIL import Image
import pytesseract
from datetime import datetime
import time
import re
import io


class TotalInputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Enter Expected Total")
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        instruction = QLabel("Please enter the expected total amount:")
        layout.addWidget(instruction)

        self.total_input = QLineEdit()
        self.total_input.setPlaceholderText("Enter amount (e.g., 1234.56)")
        layout.addWidget(self.total_input)

        buttons = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        buttons.addWidget(ok_button)
        buttons.addWidget(cancel_button)
        layout.addLayout(buttons)

    def get_total(self):
        try:
            return float(self.total_input.text())
        except ValueError:
            return None


class ErrorCorrectionDialog(QDialog):
    def __init__(self, expenses, expected_total=None, parent=None):
        super().__init__(parent)
        self.expenses = expenses.copy()
        self.expected_total = expected_total
        self.setWindowTitle("Error Correction")
        self.setModal(True)
        self.setup_ui()
        self.filter_suspicious_entries()
        self.update_totals()

    def is_suspicious(self, expense):
        """Check if an expense entry might need review"""
        amount = expense["amount"]

        if amount == int(amount):
            return True
        if amount % 1 in [0.99, 0.95]:
            return False
        if amount < 10:
            return True

        cents = int((amount % 1) * 100)
        if cents not in [0, 50, 95, 99] and cents < 90:
            return True

        return False

    def filter_suspicious_entries(self):
        """Hide rows that don't need review"""
        for row in range(self.table.rowCount()):
            amount = float(self.table.item(row, 1).text())
            expense = {"amount": amount, "remark": self.table.item(row, 2).text()}

            if not self.is_suspicious(expense):
                self.table.hideRow(row)

    def setup_ui(self):
        layout = QVBoxLayout(self)

        totals_layout = QHBoxLayout()
        self.expected_total_label = QLabel(
            f"Expected Total: ${self.expected_total:.2f}"
            if self.expected_total
            else "Expected Total: Not Set"
        )
        self.current_total_label = QLabel("Current Total: $0.00")
        self.difference_label = QLabel("Difference: $0.00")

        totals_layout.addWidget(self.expected_total_label)
        totals_layout.addWidget(self.current_total_label)
        totals_layout.addWidget(self.difference_label)
        layout.addLayout(totals_layout)

        instruction_label = QLabel(
            "Review highlighted entries that might need correction.\n"
            "Common issues:\n"
            "- Missing digits (e.g., $1.34 instead of $51.34)\n"
            "- Incorrect decimal places\n"
            "- OCR misreading numbers\n\n"
            "Non-suspicious entries are hidden automatically."
        )
        instruction_label.setWordWrap(True)
        layout.addWidget(instruction_label)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Entry #", "Amount", "Remark"])
        self.table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self.table.itemChanged.connect(self.update_totals)

        self.table.setRowCount(len(self.expenses))
        for i, expense in enumerate(self.expenses):
            entry_item = QTableWidgetItem(str(i + 1))
            entry_item.setFlags(entry_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, entry_item)

            amount_item = QTableWidgetItem(f"{expense['amount']:.2f}")
            self.table.setItem(i, 1, amount_item)

            remark_item = QTableWidgetItem(expense["remark"])
            remark_item.setFlags(remark_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 2, remark_item)

        layout.addWidget(self.table)

        button_layout = QHBoxLayout()

        show_all_button = QPushButton("Show All Entries")
        show_all_button.clicked.connect(self.show_all_entries)
        button_layout.addWidget(show_all_button)

        save_button = QPushButton("Save Corrections")
        save_button.clicked.connect(self.check_total_match)
        button_layout.addWidget(save_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

    def show_all_entries(self):
        """Show all rows, including non-suspicious ones"""
        for row in range(self.table.rowCount()):
            self.table.showRow(row)

    def update_totals(self):
        """Update the total and difference labels"""
        try:
            current_total = sum(
                float(self.table.item(row, 1).text())
                for row in range(self.table.rowCount())
                if self.table.item(row, 1) is not None
            )

            self.current_total_label.setText(f"Current Total: ${current_total:.2f}")

            if self.expected_total is not None:
                difference = self.expected_total - current_total
                self.difference_label.setText(f"Difference: ${difference:.2f}")

                if abs(difference) < 0.01:
                    self.difference_label.setStyleSheet("color: green;")
                else:
                    self.difference_label.setStyleSheet("color: red;")
        except Exception as e:
            print(f"Error updating totals: {str(e)}")
            self.current_total_label.setText("Current Total: $0.00")
            self.difference_label.setText("Difference: N/A")

    def check_total_match(self):
        """Check if totals match before accepting"""
        if self.expected_total is not None:
            current_total = sum(
                float(self.table.item(row, 1).text())
                for row in range(self.table.rowCount())
            )
            difference = abs(self.expected_total - current_total)

            if difference > 0.01:
                reply = QMessageBox.question(
                    self,
                    "Totals Don't Match",
                    f"There is still a difference of ${difference:.2f}. "
                    f"Do you want to continue reviewing?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.Yes:
                    return

        self.accept()

    def get_corrected_expenses(self):
        corrected_expenses = []
        for row in range(self.table.rowCount()):
            try:
                amount = float(self.table.item(row, 1).text())
                remark = self.table.item(row, 2).text()
                corrected_expenses.append(
                    {"amount": round(amount, 2), "remark": remark}
                )
            except (ValueError, AttributeError):
                continue
        return corrected_expenses


class FinancialScreenshotApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Initialize Tesseract path
        if platform.system().lower() == "darwin":  # Mac
            pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"
        else:  # Windows - adjust path as needed
            pytesseract.pytesseract.tesseract_cmd = (
                r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            )

        # Initialize main window properties
        self.setWindowTitle("Expense Screenshot to JSON Converter")
        self.setGeometry(100, 100, 500, 400)

        # Initialize variables
        self.expenses = []
        self.total_processed = 0
        self.expected_total = None

        # Create and set up the UI
        self.create_ui()

    def create_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Instructions label
        instructions = QLabel(
            "Instructions:\n\n"
            "1. Click 'Set Expected Total' to enter the known total\n"
            "2. Click 'Capture Region'\n"
            "3. Click 'OK' when prompted to start the capture process\n"
            "4. Select the region containing Amount and Remark columns\n"
            "5. Review and correct any highlighted entries\n"
            "6. Repeat if more data needs to be captured\n"
            "7. Save to JSON when finished"
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        # Add Set Total button
        self.set_total_btn = QPushButton("Set Expected Total")
        self.set_total_btn.clicked.connect(self.set_expected_total)
        layout.addWidget(self.set_total_btn)

        # Buttons
        self.capture_btn = QPushButton("Capture Region")
        self.capture_btn.clicked.connect(self.capture_screenshot)
        layout.addWidget(self.capture_btn)

        self.save_btn = QPushButton("Save to JSON")
        self.save_btn.clicked.connect(self.save_json)
        layout.addWidget(self.save_btn)

        # Status labels
        self.total_label = QLabel("Expected Total: Not Set")
        layout.addWidget(self.total_label)

        self.counter_label = QLabel("Processed expenses: 0")
        layout.addWidget(self.counter_label)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        # Add Exit button
        self.exit_btn = QPushButton("Exit")
        self.exit_btn.clicked.connect(self.confirm_exit)
        self.exit_btn.setStyleSheet("background-color: #ff4444; color: white;")
        layout.addWidget(self.exit_btn)

    def set_expected_total(self):
        dialog = TotalInputDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.expected_total = dialog.get_total()
            if self.expected_total is not None:
                self.total_label.setText(f"Expected Total: ${self.expected_total:.2f}")
            else:
                QMessageBox.warning(
                    self, "Invalid Input", "Please enter a valid number."
                )

    def confirm_exit(self):
        if self.expenses and len(self.expenses) > 0:
            reply = QMessageBox.question(
                self,
                "Confirm Exit",
                "You have unsaved expenses. Are you sure you want to exit?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.close()
        else:
            self.close()

    def capture_screenshot(self):
        try:
            # Show the message box first, before minimizing
            QMessageBox.information(
                self,
                "Region Selection",
                "Click OK to start the screenshot capture process.\n\n"
                "After clicking OK, select the region containing both Amount and Remark columns.",
            )

            # Hide the main window and give time for UI to update
            self.hide()
            QApplication.processEvents()
            time.sleep(1)  # Increased delay to ensure window is hidden

            if platform.system().lower() == "darwin":  # Mac
                result = subprocess.run(
                    ["screencapture", "-i", "-"], capture_output=True, check=True
                )
                img = Image.open(io.BytesIO(result.stdout))
            else:  # Windows
                result = subprocess.run(
                    ["snippingtool", "/clip"], capture_output=True, check=True
                )
                img = Image.open(io.BytesIO(result.stdout))

            text = pytesseract.image_to_string(img, config="--psm 6")
            new_expenses = self.process_text_to_json(text)

            if new_expenses:
                dialog = ErrorCorrectionDialog(new_expenses, self.expected_total, self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    corrected_expenses = dialog.get_corrected_expenses()
                    self.expenses.extend(corrected_expenses)
                    self.total_processed += len(corrected_expenses)
                    self.counter_label.setText(
                        f"Processed expenses: {self.total_processed}"
                    )
            else:
                QMessageBox.warning(
                    self,
                    "Warning",
                    "No valid expense data found in selection.\n"
                    "Make sure Amount and Remark columns are clearly visible.",
                )

        except Exception as e:
            QMessageBox.critical(
                self, "Error", f"Failed to capture/process region: {str(e)}"
            )
        finally:
            # Always make sure the window is shown again
            self.show()
            QApplication.processEvents()

    def process_text_to_json(self, text):
        """Convert OCR text to structured JSON data with improved parsing"""
        try:
            # Split text into lines and filter empty lines
            lines = [line.strip() for line in text.split("\n") if line.strip()]

            new_expenses = []
            # More strict pattern for positive numbers with optional decimals
            amount_pattern = r"^(\d+\.?\d*)"

            for line in lines:
                # Split line into components
                parts = line.split()
                if not parts:
                    continue

                # Try to find amount
                amount_match = re.match(amount_pattern, parts[0])
                if amount_match:
                    try:
                        amount = float(amount_match.group().replace(",", ""))

                        # Look for remark at the end of the line
                        remark = "No Remark Available"

                        # Check if there's a potential remark at the end
                        last_part = parts[-1]
                        if last_part.isdigit() and len(last_part) <= 6:
                            remark = last_part

                        new_expenses.append(
                            {"amount": round(amount, 2), "remark": remark}
                        )
                    except ValueError:
                        continue

            return new_expenses

        except Exception as e:
            print(f"Error processing text: {str(e)}")
            return []

    def save_json(self):
        if not self.expenses:
            QMessageBox.warning(self, "Warning", "No expense data to save.")
            return

        total = sum(expense["amount"] for expense in self.expenses)

        if self.expected_total is not None:
            difference = abs(self.expected_total - total)
            if difference > 0.01:  # Allow for floating point imprecision
                reply = QMessageBox.question(
                    self,
                    "Totals Don't Match",
                    f"Current total (${
                        total:.2f}) differs from expected total "
                    f"(${
                        self.expected_total:.2f}) by ${
                        difference:.2f}.\n\n"
                    f"Do you want to save anyway?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.No:
                    return

        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Save JSON File",
            f"expenses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            "JSON Files (*.json)",
        )

        if not file_name:
            return

        try:
            final_json = {"expenses": sorted(self.expenses, key=lambda x: x["amount"])}

            with open(file_name, "w") as f:
                json.dump(final_json, f, indent=1)

            QMessageBox.information(
                self,
                "Success",
                f"Successfully saved {len(self.expenses)} expenses to JSON!\n"
                f"File: {file_name}",
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to save JSON: {
                    str(e)}",
            )


def capture_screenshot(self):
    try:
        # Show the message box first, before minimizing
        QMessageBox.information(
            self,
            "Region Selection",
            "Click OK to start the screenshot capture process.\n\n"
            "After clicking OK, select the region containing both Amount and Remark columns.",
        )

        # Hide the main window and give time for UI to update
        self.hide()
        QApplication.processEvents()
        time.sleep(1)  # Increased delay to ensure window is hidden

        if platform.system().lower() == "darwin":  # Mac
            result = subprocess.run(
                ["screencapture", "-i", "-"], capture_output=True, check=True
            )
            img = Image.open(io.BytesIO(result.stdout))
        else:  # Windows
            # Similar modification for Windows...
            pass

        text = pytesseract.image_to_string(img, config="--psm 6")
        new_expenses = self.process_text_to_json(text)

        if new_expenses:
            dialog = ErrorCorrectionDialog(new_expenses, self.expected_total, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                corrected_expenses = dialog.get_corrected_expenses()
                self.expenses.extend(corrected_expenses)
                self.total_processed += len(corrected_expenses)
                self.counter_label.setText(
                    f"Processed expenses: {self.total_processed}"
                )
        else:
            QMessageBox.warning(
                self,
                "Warning",
                "No valid expense data found in selection.\n"
                "Make sure Amount and Remark columns are clearly visible.",
            )

    except Exception as e:
        QMessageBox.critical(
            self, "Error", f"Failed to capture/process region: {str(e)}"
        )
    finally:
        # Always make sure the window is shown again
        self.show()
        QApplication.processEvents()


def main():
    app = QApplication(sys.argv)
    window = FinancialScreenshotApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
