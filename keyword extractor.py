import sys
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QFileDialog, QTextEdit, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import docx2txt
from pdfminer.high_level import extract_text
import re
import unicodedata
import pandas as pd

class CVSearchApp(QWidget):
    def __init__(self):
        super().__init__()

        self.keywords = []
        self.cv_texts = []
        self.cv_names = []
        self.turkish_names = []

        self.initUI()

    def initUI(self):
        self.setWindowTitle('CV Keyword Hunter')
        self.setGeometry(100, 100, 600, 400)
        self.setWindowIcon(QIcon(r'C:\Users\berke\OneDrive\Belgeler\Projeler\keyword extractor\app_icon.ico'))

        layout = QVBoxLayout()

        # Keyword entry layout
        keyword_layout = QHBoxLayout()
        self.keyword_label = QLabel("Keywords:")
        self.keyword_entry = QLineEdit()
        self.add_button = QPushButton("Add")
        self.add_button.clicked.connect(self.add_keyword)

        keyword_layout.addWidget(self.keyword_label)
        keyword_layout.addWidget(self.keyword_entry)
        keyword_layout.addWidget(self.add_button)

        # Keywords display
        self.keywords_display = QLabel("Added Keywords: None")
        self.keywords_display.setWordWrap(True)

        # Reset keywords button
        self.reset_keywords_button = QPushButton("Reset Keywords")
        self.reset_keywords_button.clicked.connect(self.reset_keywords)

        # Buttons layout
        buttons_layout = QHBoxLayout()
        self.cv_button = QPushButton("Upload CVs")
        self.cv_button.clicked.connect(self.upload_cvs)
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_cv)
        self.reset_cvs_button = QPushButton("Reset CVs")
        self.reset_cvs_button.clicked.connect(self.reset_cvs)

        buttons_layout.addWidget(self.cv_button)
        buttons_layout.addWidget(self.search_button)
        buttons_layout.addWidget(self.reset_cvs_button)

        # Report button
        self.report_button = QPushButton("Generate Report")
        self.report_button.clicked.connect(self.generate_report)
        buttons_layout.addWidget(self.report_button)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)

        # Text display area
        self.cv_textbox = QTextEdit()
        self.cv_textbox.setReadOnly(True)

        layout.addLayout(keyword_layout)
        layout.addWidget(self.keywords_display)
        layout.addWidget(self.reset_keywords_button)
        layout.addLayout(buttons_layout)
        layout.addWidget(self.status_label)
        layout.addWidget(self.cv_textbox)

        self.setLayout(layout)

    def add_keyword(self):
        keyword = self.keyword_entry.text()
        if keyword:
            if keyword in self.keywords:
                QMessageBox.warning(self, "Warning", "This keyword has already been added.")
            else:
                self.keywords.append(keyword)
                self.keyword_entry.clear()
                self.update_keywords_display()

    def update_keywords_display(self):
        if self.keywords:
            keywords_text = "Added Keywords: " + ", ".join(self.keywords)
        else:
            keywords_text = "Added Keywords: None"
        self.keywords_display.setText(keywords_text)

    def reset_keywords(self):
        self.keywords = []
        self.update_keywords_display()

    def upload_cvs(self):
        options = QFileDialog.Options()
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Upload CVs", "", "Word Files (*.docx);;PDF Files (*.pdf)", options=options)
        if file_paths:
            for file_path in file_paths:
                if file_path.endswith('.pdf'):
                    text = extract_text(file_path)
                elif file_path.endswith('.docx'):
                    text = docx2txt.process(file_path)
                else:
                    text = ""

                self.cv_texts.append(text)
                self.cv_names.append(os.path.basename(file_path))
                self.cv_textbox.append(f"CV Uploaded: {os.path.basename(file_path)}")

    def reset_cvs(self):
        self.cv_texts = []
        self.cv_names = []
        self.cv_textbox.clear()
        self.status_label.setText("")
        self.cv_textbox.append("CVs reset.")

    def load_turkish_names(self, file_path):
        if file_path:
            with open(file_path, 'r', encoding='utf-8') as file:
                self.turkish_names = [line.strip() for line in file.readlines()]

    def normalize_text(self, text):
        # Normalizes text by removing accents and converting to lowercase
        return ''.join(
            c for c in unicodedata.normalize('NFD', text)
            if unicodedata.category(c) != 'Mn'
        ).lower()

    def title_case_name(self, name):
        # Converts a name to title case
        return ' '.join(word.capitalize() for word in name.split())

    def search_cv(self):
        if not self.keywords:
            QMessageBox.warning(self, "Warning", "Please add at least one keyword.")
            return

        # Load Turkish names from the specified text file
        self.load_turkish_names(r"C:\Users\berke\OneDrive\Belgeler\Projeler\keyword extractor\turkish_names.txt")

        self.status_label.setText("Loading...")
        QApplication.processEvents()  # Update the GUI to show the loading status

        matches = []
        match_counts = {}

        for i, cv_text in enumerate(self.cv_texts):
            if not cv_text:
                continue

            match_count = 0
            cv_info = ""

            # Normalize the CV text
            normalized_cv_text = self.normalize_text(cv_text)

            # Check if any keyword is present in the CV
            keywords_found = any(re.search(r'\b%s\b' % re.escape(self.normalize_text(keyword)), normalized_cv_text) for keyword in self.keywords)
            if not keywords_found:
                continue

            # Search for keywords in the CV
            for keyword in self.keywords:
                matches_in_cv = re.findall(r'\b%s\b' % re.escape(self.normalize_text(keyword)), normalized_cv_text)
                if matches_in_cv:
                    match_count += len(matches_in_cv)
                    cv_info += f"'{keyword}' found {len(matches_in_cv)} times.\n"

            # Search for names in the CV
            lines = normalized_cv_text.split('\n')
            for idx, line in enumerate(lines):
                line = line.strip()
                if line and any(self.normalize_text(name) == line.split()[0] for name in self.turkish_names):
                    next_line = lines[idx + 1].strip() if idx + 1 < len(lines) else ""
                    if next_line and len(next_line.split()) == 1:
                        full_name = line + " " + next_line
                        cv_info += f"Name: {self.title_case_name(full_name)}\n"
                    else:
                        cv_info += f"Name: {self.title_case_name(line)}\n"

            if match_count > 0 or cv_info:
                matches.append((i, match_count, cv_info))
                match_counts[self.cv_names[i]] = match_count

        matches.sort(key=lambda x: x[1], reverse=True)

        result_text = ""
        if matches:
            for match in matches:
                cv_index, match_count, cv_info = match
                cv_name = self.cv_names[cv_index]
                result_text += f"{cv_name} - {match_count} matches\n{cv_info}\n\n"
            self.status_label.setText("Completed")
        else:
            result_text = "Not found"
            self.status_label.setText("Completed")

        self.cv_textbox.append(result_text)

    def generate_report(self):
        if not self.keywords:
            QMessageBox.warning(self, "Warning", "Please add at least one keyword.")
            return

        report_data = []
        for i, cv_text in enumerate(self.cv_texts):
            if not cv_text:
                continue

            normalized_cv_text = self.normalize_text(cv_text)
            name_found = ""
            keywords_found = []

            lines = normalized_cv_text.split('\n')
            for idx, line in enumerate(lines):
                line = line.strip()
                if line and any(self.normalize_text(name) == line.split()[0] for name in self.turkish_names):
                    next_line = lines[idx + 1].strip() if idx + 1 < len(lines) else ""
                    if next_line and len(next_line.split()) == 1:
                        name_found = self.title_case_name(line + " " + next_line)
                    else:
                        name_found = self.title_case_name(line)
                    break

            for keyword in self.keywords:
                if re.search(r'\b%s\b' % re.escape(self.normalize_text(keyword)), normalized_cv_text):
                    keywords_found.append(keyword)

            for keyword in keywords_found:
                report_data.append([name_found, keyword, self.cv_names[i]])

        df = pd.DataFrame(report_data, columns=["Name", "Keyword", "CV Name"])
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Report", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_path:
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Information", "Report successfully saved.")
        else:
            QMessageBox.warning(self, "Warning", "Failed to save the report.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CVSearchApp()
    ex.show()
    sys.exit(app.exec_())
