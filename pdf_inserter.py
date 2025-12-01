import sys
import os
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices, QFont, QIcon, QPalette, QColor
import fitz  # PyMuPDF
from pathlib import Path


class PDFMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf1_path = ""
        self.pdf2_path = ""
        self.pdf1_pages = 0
        self.pdf2_pages = 0
        self.initUI()

    def initUI(self):
        self.setWindowTitle(
            "Advanced PDF Merger - Insert Pages at Specific Positions")
        self.setGeometry(100, 100, 1100, 900)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("Advanced PDF Page Inserter")
        title_font = QFont("Arial", 18, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin: 20px;")
        main_layout.addWidget(title_label)

        # Description
        desc_label = QLabel(
            "Insert specific pages from PDF2 into specific positions in PDF1")
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setStyleSheet("color: #7f8c8d; margin-bottom: 20px;")
        main_layout.addWidget(desc_label)

        # PDF 1 Section (Main PDF)
        pdf1_group = self.create_pdf_section("Main PDF (Base Document)", 1)
        main_layout.addWidget(pdf1_group)

        # PDF 2 Section (Source PDF for pages)
        pdf2_group = self.create_pdf_section("Source PDF (Pages to Insert)", 2)
        main_layout.addWidget(pdf2_group)

        # Page Insertion Rules Section
        rules_group = QGroupBox("Page Insertion Rules")
        rules_group.setStyleSheet(
            "QGroupBox { font-weight: bold; margin-top: 15px; }")
        rules_layout = QVBoxLayout()

        # Explanation
        explanation = QLabel(
            "Define which pages to insert where. Example: '1,2' means insert PDF2 page 1 at PDF1 position 1, PDF2 page 2 at PDF1 position 2")
        explanation.setWordWrap(True)
        explanation.setStyleSheet(
            "color: #7f8c8d; padding: 5px; background-color: #f8f9fa; border-radius: 5px;")
        rules_layout.addWidget(explanation)

        # Rules input grid
        grid_layout = QGridLayout()

        # PDF2 Pages to Insert
        grid_layout.addWidget(QLabel("PDF2 Pages to Insert:"), 0, 0)
        self.pdf2_pages_to_insert = QLineEdit()
        self.pdf2_pages_to_insert.setPlaceholderText(
            "e.g., 1,3,5 or 1-3,5,7 (pages from Source PDF)")
        grid_layout.addWidget(self.pdf2_pages_to_insert, 0, 1)

        # PDF1 Positions to Insert At
        grid_layout.addWidget(QLabel("Insert at PDF1 Positions:"), 1, 0)
        self.pdf1_insert_positions = QLineEdit()
        self.pdf1_insert_positions.setPlaceholderText(
            "e.g., 2,4,6 (positions in Main PDF where to insert)")
        grid_layout.addWidget(self.pdf1_insert_positions, 1, 1)

        # Insertion Mode
        grid_layout.addWidget(QLabel("Insertion Mode:"), 2, 0)
        self.insertion_mode = QComboBox()
        self.insertion_mode.addItems(
            ["Replace existing pages", "Insert before position", "Insert after position", "Append at end"])
        grid_layout.addWidget(self.insertion_mode, 2, 1)

        rules_layout.addLayout(grid_layout)

        # Examples
        examples_label = QLabel(
            "Examples:\n• '1,3' and '2,4' → PDF2 pages 1,3 replace PDF1 pages 2,4\n• '1-3' and '5' → PDF2 pages 1-3 inserted at PDF1 position 5\n• Leave empty for 'append at end' mode")
        examples_label.setWordWrap(True)
        examples_label.setStyleSheet(
            "color: #3498db; font-size: 11px; padding: 10px; background-color: #ebf5fb; border-radius: 5px;")
        rules_layout.addWidget(examples_label)

        rules_group.setLayout(rules_layout)
        main_layout.addWidget(rules_group)

        # Quick Actions (Common Scenarios)
        actions_group = QGroupBox("Quick Actions")
        actions_group.setStyleSheet(
            "QGroupBox { font-weight: bold; margin-top: 15px; }")
        actions_layout = QHBoxLayout()

        # Quick action buttons
        replace_first_btn = QPushButton("Replace First Page")
        replace_first_btn.clicked.connect(
            lambda: self.set_quick_action("1", "1"))
        replace_first_btn.setStyleSheet(
            "padding: 8px; background-color: #f39c12; color: white;")

        insert_at_middle_btn = QPushButton("Insert at Middle")
        insert_at_middle_btn.clicked.connect(
            lambda: self.set_quick_action("1-3", "mid"))
        insert_at_middle_btn.setStyleSheet(
            "padding: 8px; background-color: #9b59b6; color: white;")

        append_all_btn = QPushButton("Append All Pages")
        append_all_btn.clicked.connect(
            lambda: self.set_quick_action("all", "end"))
        append_all_btn.setStyleSheet(
            "padding: 8px; background-color: #2ecc71; color: white;")

        actions_layout.addWidget(replace_first_btn)
        actions_layout.addWidget(insert_at_middle_btn)
        actions_layout.addWidget(append_all_btn)
        actions_layout.addStretch()

        actions_group.setLayout(actions_layout)
        main_layout.addWidget(actions_group)

        # Output options
        output_group = QGroupBox("Output Settings")
        output_group.setStyleSheet(
            "QGroupBox { font-weight: bold; margin-top: 15px; }")
        output_layout = QHBoxLayout()

        output_layout.addWidget(QLabel("Output Filename:"))
        self.output_name = QLineEdit()
        self.output_name.setPlaceholderText("merged_output.pdf")
        output_layout.addWidget(self.output_name)

        output_layout.addWidget(QLabel("Save to:"))
        self.output_path = QLineEdit()
        self.output_path.setText(str(Path.home() / "Downloads"))
        output_layout.addWidget(self.output_path)

        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_output_path)
        browse_btn.setStyleSheet("padding: 5px 10px;")
        output_layout.addWidget(browse_btn)

        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)

        # Action buttons
        button_layout = QHBoxLayout()

        self.preview_btn = QPushButton("Preview Result")
        self.preview_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        self.preview_btn.clicked.connect(self.preview_result)
        self.preview_btn.setEnabled(False)

        self.merge_btn = QPushButton("Merge with Page Insertion")
        self.merge_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        self.merge_btn.clicked.connect(self.merge_with_insertion)
        self.merge_btn.setEnabled(False)

        clear_btn = QPushButton("Clear All")
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        clear_btn.clicked.connect(self.clear_all)

        button_layout.addWidget(self.preview_btn)
        button_layout.addWidget(self.merge_btn)
        button_layout.addWidget(clear_btn)

        main_layout.addLayout(button_layout)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Preview area
        self.preview_area = QTextEdit()
        self.preview_area.setReadOnly(True)
        self.preview_area.setMaximumHeight(150)
        self.preview_area.setVisible(False)
        main_layout.addWidget(self.preview_area)

        main_layout.addStretch()

    def create_pdf_section(self, title, pdf_num):
        group = QGroupBox(title)
        group.setStyleSheet("QGroupBox { font-weight: bold; }")
        layout = QHBoxLayout()

        path_label = QLineEdit()
        path_label.setReadOnly(True)
        path_label.setPlaceholderText(f"No PDF selected for {title}")
        path_label.setStyleSheet("padding: 8px; background-color: #f8f9fa;")

        if pdf_num == 1:
            self.pdf1_path_label = path_label
        else:
            self.pdf2_path_label = path_label

        layout.addWidget(path_label)

        select_btn = QPushButton("Select PDF")
        select_btn.clicked.connect(lambda: self.select_pdf(pdf_num))
        select_btn.setStyleSheet("padding: 8px 15px;")

        view_btn = QPushButton("View")
        view_btn.clicked.connect(lambda: self.view_pdf(pdf_num))
        view_btn.setStyleSheet("padding: 8px 15px;")
        view_btn.setEnabled(False)

        if pdf_num == 1:
            self.view_btn1 = view_btn
        else:
            self.view_btn2 = view_btn

        layout.addWidget(select_btn)
        layout.addWidget(view_btn)

        group.setLayout(layout)
        return group

    def select_pdf(self, pdf_num):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"Select PDF File {pdf_num}",
            "",
            "PDF Files (*.pdf);;All Files (*)"
        )

        if file_path:
            try:
                pdf_document = fitz.open(file_path)
                page_count = pdf_document.page_count
                pdf_document.close()

                if pdf_num == 1:
                    self.pdf1_path = file_path
                    self.pdf1_path_label.setText(file_path)
                    self.pdf1_pages = page_count
                    self.view_btn1.setEnabled(True)
                    self.status_bar.showMessage(
                        f"Main PDF loaded: {page_count} pages")
                else:
                    self.pdf2_path = file_path
                    self.pdf2_path_label.setText(file_path)
                    self.pdf2_pages = page_count
                    self.view_btn2.setEnabled(True)
                    self.status_bar.showMessage(
                        f"Source PDF loaded: {page_count} pages")

                if self.pdf1_path and self.pdf2_path:
                    self.merge_btn.setEnabled(True)
                    self.preview_btn.setEnabled(True)

            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Failed to load PDF: {str(e)}")

    def view_pdf(self, pdf_num):
        file_path = self.pdf1_path if pdf_num == 1 else self.pdf2_path
        if file_path and os.path.exists(file_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))

    def browse_output_path(self):
        directory = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory",
            self.output_path.text()
        )
        if directory:
            self.output_path.setText(directory)

    def parse_page_range(self, page_str, max_pages, allow_all=False):
        """Parse page range string like '1-5,7,9-12' or 'all'"""
        if not page_str:
            return [] if not allow_all else list(range(1, max_pages + 1))

        if page_str.lower() == 'all':
            return list(range(1, max_pages + 1))

        pages = set()
        parts = page_str.split(',')

        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = part.split('-')
                try:
                    start_page = max(1, int(start))
                    end_page = min(max_pages, int(end))
                    if start_page <= end_page:
                        pages.update(range(start_page, end_page + 1))
                except ValueError:
                    continue
            else:
                try:
                    page = int(part)
                    if 1 <= page <= max_pages:
                        pages.add(page)
                except ValueError:
                    continue

        return sorted(list(pages))

    def parse_positions(self, positions_str, max_pages):
        """Parse positions string, handle 'mid' and 'end' special values"""
        if not positions_str:
            return []

        positions = []
        parts = positions_str.split(',')

        for part in parts:
            part = part.strip().lower()

            if part == 'mid':
                # Insert at middle position
                mid_pos = max_pages // 2 + 1
                positions.append(mid_pos)
            elif part == 'end':
                # Append at end
                positions.append(max_pages + 1)
            elif '-' in part:
                # Range of positions
                start, end = part.split('-')
                try:
                    start_pos = max(1, int(start))
                    end_pos = min(max_pages + 1, int(end))
                    if start_pos <= end_pos:
                        positions.extend(range(start_pos, end_pos + 1))
                except ValueError:
                    continue
            else:
                # Single position
                try:
                    pos = int(part)
                    if 1 <= pos <= max_pages + 1:
                        positions.append(pos)
                except ValueError:
                    continue

        return positions

    def set_quick_action(self, pages, positions):
        """Set up quick actions for common scenarios"""
        if not self.pdf2_path:
            QMessageBox.warning(
                self, "Warning", "Please select Source PDF first!")
            return

        self.pdf2_pages_to_insert.setText(pages)

        if positions == "1":
            self.pdf1_insert_positions.setText("1")
            self.insertion_mode.setCurrentText("Replace existing pages")
        elif positions == "mid":
            self.pdf1_insert_positions.setText("mid")
            self.insertion_mode.setCurrentText("Insert before position")
        elif positions == "end":
            self.pdf1_insert_positions.setText("")
            self.insertion_mode.setCurrentText("Append at end")

        self.status_bar.showMessage("Quick action applied")

    def resize_page_to_match(self, source_page, target_page_size):
        """Resize source page to match target page size"""
        # Get source and target page rectangles
        source_rect = source_page.rect
        target_rect = fitz.Rect(
            0, 0, target_page_size.width, target_page_size.height)

        # Calculate scaling factors
        scale_x = target_rect.width / source_rect.width
        scale_y = target_rect.height / source_rect.height

        # Use the smaller scale to maintain aspect ratio
        scale = min(scale_x, scale_y)

        # Create a new blank page with target size
        new_page = fitz.open()  # Create a temporary document
        temp_page = new_page.new_page(
            width=target_rect.width, height=target_rect.height)

        # Calculate position to center the scaled page
        scaled_width = source_rect.width * scale
        scaled_height = source_rect.height * scale
        x_pos = (target_rect.width - scaled_width) / 2
        y_pos = (target_rect.height - scaled_height) / 2

        # Create transformation matrix for scaling and positioning
        matrix = fitz.Matrix(scale, scale).pretranslate(x_pos, y_pos)

        # Insert the source page into the new page
        temp_page.show_pdf_page(
            fitz.Rect(0, 0, target_rect.width, target_rect.height),
            source_page.parent,
            source_page.number
        )

        return new_page

    def merge_with_insertion(self):
        """Main function to merge PDFs with page insertion at specific positions"""
        if not self.pdf1_path or not self.pdf2_path:
            QMessageBox.warning(
                self, "Warning", "Please select both PDF files first!")
            return

        # Get output filename
        output_filename = self.output_name.text().strip()
        if not output_filename:
            output_filename = "merged_output.pdf"
        elif not output_filename.endswith('.pdf'):
            output_filename += '.pdf'

        output_dir = self.output_path.text().strip()
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Cannot create output directory: {str(e)}")
                return

        output_path = os.path.join(output_dir, output_filename)

        # Check if file exists
        if os.path.exists(output_path):
            reply = QMessageBox.question(
                self,
                "File Exists",
                f"File '{output_filename}' already exists. Overwrite?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.status_bar.showMessage("Processing PDFs...")

            # Open PDFs
            pdf1 = fitz.open(self.pdf1_path)  # Main PDF
            pdf2 = fitz.open(self.pdf2_path)  # Source PDF

            # Parse which pages to insert from PDF2
            pdf2_pages_to_insert = self.parse_page_range(
                self.pdf2_pages_to_insert.text(),
                self.pdf2_pages,
                allow_all=True
            )

            # Parse positions in PDF1 where to insert
            mode = self.insertion_mode.currentText()
            positions_str = self.pdf1_insert_positions.text()

            if mode == "Append at end":
                # Simple append all pages at the end
                pdf1_pages_to_insert_at = [
                    self.pdf1_pages + 1] * len(pdf2_pages_to_insert)
            else:
                # Parse positions
                pdf1_pages_to_insert_at = self.parse_positions(
                    positions_str,
                    self.pdf1_pages
                )

            # Validate inputs
            if not pdf2_pages_to_insert:
                QMessageBox.warning(
                    self, "Warning", "Please specify which pages to insert from Source PDF!")
                return

            if mode != "Append at end" and not pdf1_pages_to_insert_at:
                QMessageBox.warning(
                    self, "Warning", "Please specify where to insert the pages in Main PDF!")
                return

            # Adjust for 0-based indexing
            pdf2_pages_idx = [p-1 for p in pdf2_pages_to_insert]
            pdf1_positions_idx = [p-1 for p in pdf1_pages_to_insert_at]

            # Create new PDF
            merged_pdf = fitz.open()

            # Handle different insertion modes
            if mode == "Replace existing pages":
                # Replace pages at specified positions
                result_pages = self.replace_pages(
                    pdf1, pdf2, pdf2_pages_idx, pdf1_positions_idx)
            elif mode == "Insert before position":
                # Insert before specified positions
                result_pages = self.insert_pages_before(
                    pdf1, pdf2, pdf2_pages_idx, pdf1_positions_idx)
            elif mode == "Insert after position":
                # Insert after specified positions
                result_pages = self.insert_pages_after(
                    pdf1, pdf2, pdf2_pages_idx, pdf1_positions_idx)
            else:  # Append at end
                result_pages = self.append_pages(pdf1, pdf2, pdf2_pages_idx)

            # Build the final PDF
            for i, page_info in enumerate(result_pages):
                src_pdf, page_idx = page_info
                merged_pdf.insert_pdf(
                    src_pdf, from_page=page_idx, to_page=page_idx)
                self.progress_bar.setValue(
                    int((i + 1) * 100 / len(result_pages)))

            # Save merged PDF
            merged_pdf.save(output_path)
            merged_pdf.close()

            # Close all temporary documents
            for src_pdf, _ in result_pages:
                if src_pdf != pdf1 and src_pdf != pdf2:
                    try:
                        src_pdf.close()
                    except:
                        pass

            pdf1.close()
            pdf2.close()

            self.progress_bar.setValue(100)
            self.status_bar.showMessage(
                f"PDF created successfully! Saved to: {output_path}")

            # Ask if user wants to open the merged PDF
            reply = QMessageBox.question(
                self,
                "Success",
                "PDF created successfully! Do you want to open the file?",
                QMessageBox.Yes | QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(output_path))

            self.progress_bar.setVisible(False)

        except Exception as e:
            QMessageBox.critical(
                self, "Error", f"Failed to process PDFs: {str(e)}")
            self.progress_bar.setVisible(False)
            self.status_bar.showMessage("Error processing PDFs")

    def replace_pages(self, pdf1, pdf2, pdf2_pages_idx, pdf1_positions_idx):
        """Replace pages in PDF1 with pages from PDF2 at specified positions"""
        result_pages = []

        # Ensure we have matching number of pages and positions
        if len(pdf2_pages_idx) != len(pdf1_positions_idx):
            # If positions are fewer, repeat the last position
            if len(pdf1_positions_idx) < len(pdf2_pages_idx):
                last_pos = pdf1_positions_idx[-1] if pdf1_positions_idx else 0
                pdf1_positions_idx = pdf1_positions_idx + \
                    [last_pos] * (len(pdf2_pages_idx) -
                                  len(pdf1_positions_idx))

        # Sort positions and corresponding pages
        sorted_data = sorted(zip(pdf1_positions_idx, pdf2_pages_idx))
        pdf1_positions_idx = [pos for pos, _ in sorted_data]
        pdf2_pages_idx = [page for _, page in sorted_data]

        pdf1_page_idx = 0
        replace_idx = 0

        for i in range(pdf1.page_count):
            if replace_idx < len(pdf1_positions_idx) and i == pdf1_positions_idx[replace_idx]:
                # Get the target page size from PDF1
                target_page = pdf1[i]
                target_size = target_page.rect

                # Get the source page from PDF2
                source_page = pdf2[pdf2_pages_idx[replace_idx]]

                # Create a temporary document with resized page
                temp_doc = self.resize_page_to_match(source_page, target_size)

                # Add the resized page to result
                result_pages.append((temp_doc, 0))

                replace_idx += 1
            else:
                # Keep original PDF1 page
                result_pages.append((pdf1, i))

        return result_pages

    def insert_pages_before(self, pdf1, pdf2, pdf2_pages_idx, pdf1_positions_idx):
        """Insert PDF2 pages before specified positions in PDF1"""
        result_pages = []

        # Get the standard page size from PDF1 (use first page as reference)
        standard_size = pdf1[0].rect if pdf1.page_count > 0 else fitz.Rect(
            0, 0, 595, 842)  # A4 default

        # Sort positions in descending order to maintain correct indices when inserting
        sorted_indices = sorted(
            zip(pdf1_positions_idx, pdf2_pages_idx), reverse=True)

        # Start with all PDF1 pages
        result_pages = [(pdf1, i) for i in range(pdf1.page_count)]

        # Insert PDF2 pages at specified positions
        for pos_idx, pdf2_page_idx in sorted_indices:
            if pos_idx <= len(result_pages):
                # Get the source page from PDF2
                source_page = pdf2[pdf2_page_idx]

                # Create a temporary document with resized page
                temp_doc = self.resize_page_to_match(
                    source_page, standard_size)

                # Insert the resized page
                result_pages.insert(pos_idx, (temp_doc, 0))

        return result_pages

    def insert_pages_after(self, pdf1, pdf2, pdf2_pages_idx, pdf1_positions_idx):
        """Insert PDF2 pages after specified positions in PDF1"""
        result_pages = []

        # Get the standard page size from PDF1 (use first page as reference)
        standard_size = pdf1[0].rect if pdf1.page_count > 0 else fitz.Rect(
            0, 0, 595, 842)  # A4 default

        # Sort positions in descending order to maintain correct indices when inserting
        sorted_indices = sorted(
            zip(pdf1_positions_idx, pdf2_pages_idx), reverse=True)

        # Start with all PDF1 pages
        result_pages = [(pdf1, i) for i in range(pdf1.page_count)]

        # Insert PDF2 pages after specified positions
        for pos_idx, pdf2_page_idx in sorted_indices:
            insert_pos = pos_idx + 1
            if insert_pos <= len(result_pages):
                # Get the source page from PDF2
                source_page = pdf2[pdf2_page_idx]

                # Create a temporary document with resized page
                temp_doc = self.resize_page_to_match(
                    source_page, standard_size)

                # Insert the resized page
                result_pages.insert(insert_pos, (temp_doc, 0))

        return result_pages

    def append_pages(self, pdf1, pdf2, pdf2_pages_idx):
        """Append PDF2 pages at the end of PDF1"""
        result_pages = []

        # Get the standard page size from PDF1 (use first page as reference)
        standard_size = pdf1[0].rect if pdf1.page_count > 0 else fitz.Rect(
            0, 0, 595, 842)  # A4 default

        # Add all PDF1 pages
        result_pages.extend([(pdf1, i) for i in range(pdf1.page_count)])

        # Append PDF2 pages (resized)
        for page_idx in pdf2_pages_idx:
            # Get the source page from PDF2
            source_page = pdf2[page_idx]

            # Create a temporary document with resized page
            temp_doc = self.resize_page_to_match(source_page, standard_size)

            # Append the resized page
            result_pages.append((temp_doc, 0))

        return result_pages

    def preview_result(self):
        """Preview what the result will look like"""
        if not self.pdf1_path or not self.pdf2_path:
            return

        try:
            self.preview_area.setVisible(True)
            self.preview_area.clear()

            pdf1_name = os.path.basename(self.pdf1_path)
            pdf2_name = os.path.basename(self.pdf2_path)

            # Parse inputs
            pdf2_pages_to_insert = self.parse_page_range(
                self.pdf2_pages_to_insert.text(),
                self.pdf2_pages,
                allow_all=True
            )

            mode = self.insertion_mode.currentText()
            positions_str = self.pdf1_insert_positions.text()

            if mode == "Append at end":
                pdf1_pages_to_insert_at = [
                    "end"] * len(pdf2_pages_to_insert) if pdf2_pages_to_insert else []
            else:
                pdf1_pages_to_insert_at = self.parse_positions(
                    positions_str,
                    self.pdf1_pages
                )

            # Build preview text
            preview_text = f"""
            ===== OPERATION PREVIEW =====
            
            MAIN PDF: {pdf1_name}
            Total Pages: {self.pdf1_pages}
            
            SOURCE PDF: {pdf2_name}
            Total Pages: {self.pdf2_pages}
            
            ===== OPERATION DETAILS =====
            Mode: {mode}
            
            Pages to insert from {pdf2_name}: {pdf2_pages_to_insert if pdf2_pages_to_insert else 'ALL PAGES'}
            Insert at positions in {pdf1_name}: {pdf1_pages_to_insert_at if pdf1_pages_to_insert_at else 'APPEND AT END'}
            
            ===== RESULT PREVIEW =====
            Final document will have approximately {self.pdf1_pages + len(pdf2_pages_to_insert)} pages.
            
            """

            # Add specific instructions based on mode
            if mode == "Replace existing pages" and pdf1_pages_to_insert_at:
                preview_text += "\nREPLACEMENT MAPPING:\n"
                for i, (pos, page) in enumerate(zip(pdf1_pages_to_insert_at, pdf2_pages_to_insert)):
                    if i < len(pdf2_pages_to_insert):
                        preview_text += f"  • PDF1 page {pos} ← PDF2 page {page}\n"

            elif mode == "Insert before position" and pdf1_pages_to_insert_at:
                preview_text += "\nINSERTION POINTS:\n"
                for i, (pos, page) in enumerate(zip(pdf1_pages_to_insert_at, pdf2_pages_to_insert)):
                    if i < len(pdf2_pages_to_insert):
                        preview_text += f"  • Insert PDF2 page {page} BEFORE PDF1 page {pos}\n"

            elif mode == "Insert after position" and pdf1_pages_to_insert_at:
                preview_text += "\nINSERTION POINTS:\n"
                for i, (pos, page) in enumerate(zip(pdf1_pages_to_insert_at, pdf2_pages_to_insert)):
                    if i < len(pdf2_pages_to_insert):
                        preview_text += f"  • Insert PDF2 page {page} AFTER PDF1 page {pos}\n"

            elif mode == "Append at end":
                preview_text += "\nAPPEND OPERATION:\n"
                if pdf2_pages_to_insert:
                    preview_text += f"  • Appending {len(pdf2_pages_to_insert)} pages from PDF2 at the end\n"
                else:
                    preview_text += "  • Appending ALL pages from PDF2 at the end\n"

            self.preview_area.setText(preview_text)

        except Exception as e:
            QMessageBox.critical(
                self, "Error", f"Failed to generate preview: {str(e)}")

    def clear_all(self):
        self.pdf1_path = ""
        self.pdf2_path = ""
        self.pdf1_path_label.clear()
        self.pdf2_path_label.clear()
        self.pdf2_pages_to_insert.clear()
        self.pdf1_insert_positions.clear()
        self.output_name.clear()
        self.output_path.setText(str(Path.home() / "Downloads"))
        self.insertion_mode.setCurrentIndex(0)
        self.view_btn1.setEnabled(False)
        self.view_btn2.setEnabled(False)
        self.merge_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self.preview_area.setVisible(False)
        self.preview_area.clear()
        self.status_bar.showMessage("Cleared all fields")
        self.progress_bar.setVisible(False)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = PDFMergerApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
