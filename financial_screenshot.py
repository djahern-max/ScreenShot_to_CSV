"""
Expense Screenshot Processor
A PyQt6 application for capturing and processing expense data from screenshots.
"""

import sys
import json
import os
import platform
from datetime import datetime
import time
import re
import io
from dataclasses import dataclass
from typing import List, Optional, Dict, Any
from decimal import Decimal, InvalidOperation as DecimalException

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
    QDialog,
    QLineEdit,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from PIL import Image
import pytesseract
import subprocess
import tempfile
from PIL import ImageGrab


# Data Models
@dataclass
class Expense:
    amount: Decimal
    remark: str

    def to_dict(self) -> Dict[str, Any]:
        return {"amount": float(self.amount), "remark": self.remark}


class OCRProcessor:
    """Handles OCR processing and text extraction"""

    def __init__(self):
        self._setup_tesseract()

    def _setup_tesseract(self):
        """Configure Tesseract based on operating system"""
        if platform.system().lower() == "darwin":
            pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"
        else:
            pytesseract.pytesseract.tesseract_cmd = (
                r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            )

    def process_image(self, image: Image.Image) -> List[Expense]:
        """Process image and extract expenses"""
        text = pytesseract.image_to_string(image, config="--psm 6")
        return self._parse_text(text)

    def _parse_text(self, text: str) -> List[Expense]:
        """Parse OCR text into expense objects"""
        expenses = []
        amount_pattern = r"^(\d+\.?\d*)"

        for line in text.split("\n"):
            if not line.strip():
                continue

            parts = line.split()
            amount_match = re.match(amount_pattern, parts[0])

            if amount_match:
                try:
                    amount = Decimal(amount_match.group().replace(",", ""))
                    remark = (
                        parts[-1]
                        if parts[-1].isdigit() and len(parts[-1]) <= 6
                        else "No Remark Available"
                    )
                    expenses.append(
                        Expense(amount=amount.quantize(Decimal("0.01")), remark=remark)
                    )
                except (ValueError, DecimalException):
                    continue

        return expenses


class TotalInputDialog(QDialog):
    """Dialog for entering expected total amount"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Enter Expected Total")
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # Add instruction label
        layout.addWidget(QLabel("Please enter the expected total amount:"))

        # Add input field
        self.total_input = QLineEdit()
        self.total_input.setPlaceholderText("Enter amount (e.g., 1234.56)")
        layout.addWidget(self.total_input)

        # Add buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

    def get_total(self) -> Optional[Decimal]:
        """Get the entered total amount"""
        try:
            return Decimal(self.total_input.text())
        except (ValueError, DecimalException):
            return None


class ErrorCorrectionDialog(QDialog):
    """Dialog for reviewing and correcting expense entries"""

    def __init__(
        self, expenses: List[Expense], expected_total: Optional[Decimal], parent=None
    ):
        super().__init__(parent)
        self.expenses = expenses.copy()
        self.expected_total = expected_total
        self.setWindowTitle("Error Correction")
        self.setModal(True)
        self._setup_ui()
        self._filter_suspicious_entries()
        self._update_totals()

    def _is_suspicious(self, expense: Expense) -> bool:
        """Determine if an expense entry needs review"""
        amount = float(expense.amount)
        cents = int((amount % 1) * 100)

        return any(
            [
                amount == int(amount),
                amount < 10,
                cents not in [0, 50, 95, 99] and cents < 90,
            ]
        )

    def _setup_ui(self):
        """Set up the dialog's user interface"""
        layout = QVBoxLayout(self)

        # Add totals display
        self._setup_totals_display(layout)

        # Add instructions
        self._setup_instructions(layout)

        # Add expense table
        self._setup_expense_table(layout)

        # Add buttons
        self._setup_buttons(layout)

    def _setup_totals_display(self, layout: QVBoxLayout):
        totals_layout = QHBoxLayout()

        self.expected_total_label = QLabel(
            f"Expected Total: ${float(self.expected_total):.2f}"
            if self.expected_total
            else "Expected Total: Not Set"
        )
        self.current_total_label = QLabel("Current Total: $0.00")
        self.difference_label = QLabel("Difference: $0.00")

        for label in [
            self.expected_total_label,
            self.current_total_label,
            self.difference_label,
        ]:
            totals_layout.addWidget(label)

        layout.addLayout(totals_layout)

    def _setup_instructions(self, layout: QVBoxLayout):
        instructions = QLabel(
            "Review highlighted entries that might need correction.\n"
            "Common issues:\n"
            "- Missing digits (e.g., $1.34 instead of $51.34)\n"
            "- Incorrect decimal places\n"
            "- OCR misreading numbers\n\n"
            "Non-suspicious entries are hidden automatically."
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

    def _setup_expense_table(self, layout: QVBoxLayout):
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Entry #", "Amount", "Remark"])
        self.table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self.table.itemChanged.connect(self._update_totals)

        # Populate table
        self.table.setRowCount(len(self.expenses))
        for i, expense in enumerate(self.expenses):
            # Entry number
            entry_item = QTableWidgetItem(str(i + 1))
            entry_item.setFlags(entry_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, entry_item)

            # Amount
            amount_item = QTableWidgetItem(f"{float(expense.amount):.2f}")
            self.table.setItem(i, 1, amount_item)

            # Remark
            remark_item = QTableWidgetItem(expense.remark)
            remark_item.setFlags(remark_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 2, remark_item)

        layout.addWidget(self.table)

    def _setup_buttons(self, layout: QVBoxLayout):
        button_layout = QHBoxLayout()

        show_all_button = QPushButton("Show All Entries")
        show_all_button.clicked.connect(self.show_all_entries)

        save_button = QPushButton("Save Corrections")
        save_button.clicked.connect(self._check_total_match)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        for button in [show_all_button, save_button, cancel_button]:
            button_layout.addWidget(button)

        layout.addLayout(button_layout)

    def _filter_suspicious_entries(self):
        """Hide rows that don't need review"""
        for row in range(self.table.rowCount()):
            amount = Decimal(self.table.item(row, 1).text())
            expense = Expense(amount=amount, remark=self.table.item(row, 2).text())

            if not self._is_suspicious(expense):
                self.table.hideRow(row)

    def show_all_entries(self):
        """Show all rows, including non-suspicious ones"""
        for row in range(self.table.rowCount()):
            self.table.showRow(row)

    def _update_totals(self):
        """Update the total and difference labels"""
        try:
            current_total = Decimal("0")
            for row in range(self.table.rowCount()):
                if self.table.item(row, 1) is not None:
                    current_total += Decimal(self.table.item(row, 1).text())

            self.current_total_label.setText(
                f"Current Total: ${float(current_total):.2f}"
            )

            if self.expected_total is not None:
                difference = self.expected_total - current_total
                self.difference_label.setText(f"Difference: ${float(difference):.2f}")

                if abs(difference) < Decimal("0.01"):
                    self.difference_label.setStyleSheet("color: green;")
                else:
                    self.difference_label.setStyleSheet("color: red;")
        except Exception as e:
            print(f"Error updating totals: {str(e)}")
            self.current_total_label.setText("Current Total: $0.00")
            self.difference_label.setText("Difference: N/A")

    def _check_total_match(self):
        """Check if totals match before accepting"""
        if self.expected_total is not None:
            current_total = Decimal("0")
            for row in range(self.table.rowCount()):
                current_total += Decimal(self.table.item(row, 1).text())

            difference = abs(self.expected_total - current_total)

            if difference > Decimal("0.01"):
                reply = QMessageBox.question(
                    self,
                    "Totals Don't Match",
                    f"There is still a difference of ${float(difference):.2f}. "
                    f"Do you want to continue reviewing?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.Yes:
                    return

        self.accept()

    def get_corrected_expenses(self) -> List[Expense]:
        """Get the list of corrected expenses"""
        corrected_expenses = []
        for row in range(self.table.rowCount()):
            try:
                amount = Decimal(self.table.item(row, 1).text())
                remark = self.table.item(row, 2).text()
                corrected_expenses.append(
                    Expense(amount=amount.quantize(Decimal("0.01")), remark=remark)
                )
            except (ValueError, DecimalException, AttributeError):
                continue
        return corrected_expenses


class ExpenseApp(QMainWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()
        self.expenses: List[Expense] = []
        self.total_processed = 0
        self.expected_total: Optional[Decimal] = None
        self.ocr_processor = OCRProcessor()

        self._setup_window()
        self._create_ui()

    def _setup_window(self):
        """Configure main window properties"""
        self.setWindowTitle("Expense Screenshot to JSON Converter")
        self.setGeometry(100, 100, 500, 400)

    def _create_ui(self):
        """Create the main user interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self._add_instructions(layout)
        self._add_buttons(layout)
        self._add_status_labels(layout)

    def _add_instructions(self, layout: QVBoxLayout):
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

    def _add_buttons(self, layout: QVBoxLayout):
        # Set Total button
        self.set_total_btn = QPushButton("Set Expected Total")
        self.set_total_btn.clicked.connect(self._set_expected_total)
        layout.addWidget(self.set_total_btn)

        # Capture button
        self.capture_btn = QPushButton("Capture Region")
        self.capture_btn.clicked.connect(self._capture_screenshot)
        layout.addWidget(self.capture_btn)

        # Save button
        self.save_btn = QPushButton("Save to JSON")
        self.save_btn.clicked.connect(self._save_json)
        layout.addWidget(self.save_btn)

        # Exit button
        self.exit_btn = QPushButton("Exit")
        self.exit_btn.clicked.connect(self._confirm_exit)
        self.exit_btn.setStyleSheet("background-color: #ff4444; color: white;")
        layout.addWidget(self.exit_btn)

    def _add_status_labels(self, layout: QVBoxLayout):
        self.total_label = QLabel("Expected Total: Not Set")
        self.counter_label = QLabel("Processed expenses: 0")
        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)

        for label in [self.total_label, self.counter_label, self.status_label]:
            layout.addWidget(label)

    def _set_expected_total(self):
        """Set the expected total amount"""
        dialog = TotalInputDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.expected_total = dialog.get_total()
            if self.expected_total is not None:
                self.total_label.setText(
                    f"Expected Total: ${float(self.expected_total):.2f}"
                )
            else:
                QMessageBox.warning(
                    self, "Invalid Input", "Please enter a valid number."
                )

    def _capture_screenshot(self):
        """Capture and process screenshot region"""
        try:
            # Show instruction message
            QMessageBox.information(
                self,
                "Region Selection",
                "Click OK to start the screenshot capture process.\n\n"
                "After clicking OK, select the region containing both Amount and Remark columns.",
            )

            # Hide window and wait
            self.hide()
            QApplication.processEvents()
            time.sleep(1)

            img = None
            if platform.system().lower() == "darwin":  # Mac
                # Create a temporary file for the screenshot
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                try:
                    result = subprocess.run(
                        ["screencapture", "-i", temp_file.name], check=True
                    )
                    img = Image.open(temp_file.name)
                finally:
                    os.unlink(temp_file.name)  # Clean up the temporary file
            else:  # Windows
                # Use ImageGrab for Windows
                subprocess.run(["snippingtool", "/clip"], check=True)
                time.sleep(0.5)  # Give some time for the clipboard to update
                img = ImageGrab.grabclipboard()

            if img is None:
                raise Exception("Failed to capture screenshot")

            # Process the image
            new_expenses = self.ocr_processor.process_image(img)

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
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
        finally:
            self.show()
            QApplication.processEvents()

    def _save_json(self):
        """Save processed expenses to JSON file"""
        if not self.expenses:
            QMessageBox.warning(self, "Warning", "No expense data to save.")
            return

        total = sum(expense.amount for expense in self.expenses)

        # Check if totals match
        if self.expected_total is not None:
            difference = abs(self.expected_total - total)
            if difference > Decimal("0.01"):
                reply = QMessageBox.question(
                    self,
                    "Totals Don't Match",
                    f"Current total (${float(total):.2f}) differs from expected total "
                    f"(${float(self.expected_total):.2f}) by ${float(difference):.2f}.\n\n"
                    f"Do you want to save anyway?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.No:
                    return

        # Get save location
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Save JSON File",
            f"expenses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            "JSON Files (*.json)",
        )

        if not file_name:
            return

        try:
            # Convert expenses to dictionary format
            expenses_dict = {
                "expenses": [
                    expense.to_dict()
                    for expense in sorted(self.expenses, key=lambda x: x.amount)
                ]
            }

            # Save to file
            with open(file_name, "w") as f:
                json.dump(expenses_dict, f, indent=2)

            QMessageBox.information(
                self,
                "Success",
                f"Successfully saved {len(self.expenses)} expenses to JSON!\n"
                f"File: {file_name}",
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save JSON: {str(e)}")

    def _confirm_exit(self):
        """Confirm before exiting if there are unsaved changes"""
        if self.expenses:
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


def main():
    app = QApplication(sys.argv)
    window = ExpenseApp()  # Changed from FinancialScreenshotApp
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
