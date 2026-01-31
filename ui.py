import sys
import os
import subprocess
import shutil
import tempfile
import threading
from PIL import Image, ImageOps
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog,
    QLineEdit, QComboBox, QScrollArea, QFrame, QGridLayout, QHBoxLayout, QMessageBox,
    QSizePolicy, QProgressDialog, QCheckBox
)
from PyQt6.QtGui import QPixmap, QIntValidator, QImage, QIcon, QDrag
from PyQt6.QtCore import Qt, QTranslator, QLibraryInfo, QLocale, QMimeData, QPoint

from workers import ImageImportWorker, PDFCreationWorker

os.environ["LANG"] = "de_DE.UTF-8"


class DraggableLabel(QLabel):
    def __init__(self, parent, file_path, main_window):
        super().__init__(parent)
        self.file_path = file_path
        self.main_window = main_window
        self.unique_id: str | None = None  # Use type annotation to indicate it can be str or None
        self.setAcceptDrops(True)
        self.setStyleSheet("border: 2px solid transparent;")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_start_position = event.pos()

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            return
        if (event.pos() - self.drag_start_position).manhattanLength() < QApplication.startDragDistance():
            return

        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setText(self.file_path)
        drag.setMimeData(mime_data)

        pixmap = self.pixmap()
        if pixmap:
            drag.setPixmap(pixmap.scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio))
            drag.setHotSpot(QPoint(pixmap.width() // 2, pixmap.height() // 2))
        drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
            self.setStyleSheet("border: 2px solid #3498db;")

    def dragLeaveEvent(self, event):
        self.setStyleSheet("border: 2px solid transparent;")

    def dropEvent(self, event):
        source_path = event.mimeData().text()
        target_path = self.file_path
        if source_path != target_path:
            self.main_window.reorderImages(source_path, target_path)
            event.acceptProposedAction()
        self.setStyleSheet("border: 2px solid transparent;")


class ImageUploader(QWidget):
    def __init__(self):
        super().__init__()
        translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
        translator = QTranslator()
        translator.load("qtwidgets_de", translations_path)
        self.setLocale(QLocale(QLocale.Language.German))
        self.initUI()
        self.setWindowTitle("PDF Creator")
        self.images = []
        self.import_worker = None
        self.pdf_worker = None

    def initUI(self):
        layout = QVBoxLayout()
        self.setAcceptDrops(True)
        self.setMinimumSize(750, 600)
        self.setMaximumSize(750, 600)

        self.aktennummer_input = QLineEdit(self)
        self.aktennummer_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.aktennummer_input.setPlaceholderText("Aktennummer")
        self.aktennummer_input.textChanged.connect(self.updatePdfButtonState)

        self.dokumentenkürzel_input = QComboBox(self)
        self.dokumentenkürzel_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.dokumentenkürzel_input.addItems(["(Dokumentenkürzel auswählen oder leer lassen)", "GA", "ST", "PR", "UB", "OT", "BWS"])

        self.dokumentenzahl_input = QLineEdit(self)
        self.dokumentenzahl_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.dokumentenzahl_input.setPlaceholderText("Dokumentenzahl")
        self.dokumentenzahl_input.textChanged.connect(self.updatePdfButtonState)

        layout.addWidget(QLabel("Aktennummer:"))
        layout.addWidget(self.aktennummer_input)
        layout.addWidget(QLabel("Dokumentenkürzel:"))
        layout.addWidget(self.dokumentenkürzel_input)
        layout.addWidget(QLabel("Dokumentenzahl:"))
        layout.addWidget(self.dokumentenzahl_input)

        # Add starting photo number input
        start_number_layout = QHBoxLayout()
        start_number_layout.addWidget(QLabel("Startindex:"))
        self.start_photo_number = QLineEdit(self)
        self.start_photo_number.setText("1")  # Default value is 1 (not 0)
        self.start_photo_number.setValidator(QIntValidator(1, 999))  # Only allow integers
        self.start_photo_number.setFixedWidth(60)
        start_number_layout.addWidget(self.start_photo_number)
        start_number_layout.addStretch()
        layout.addLayout(start_number_layout)

        # Add ZIP saving toggle (default ON)
        self.save_images_as_zip_checkbox = QCheckBox("Bilder als ZIP-Datei speichern", self)
        self.save_images_as_zip_checkbox.setChecked(True)
        layout.addWidget(self.save_images_as_zip_checkbox)

        self.delete_originals_checkbox = QCheckBox("Alte Bilder löschen", self)
        self.delete_originals_checkbox.setChecked(False)
        layout.addWidget(self.delete_originals_checkbox)

        self.add_timestamp_checkbox = QCheckBox("Bilder mit Zeitstempeln versehen", self)
        self.add_timestamp_checkbox.setChecked(False)
        layout.addWidget(self.add_timestamp_checkbox)

        self.upload_button = QPushButton("Dateien hinzufügen", self)
        self.upload_button.clicked.connect(self.openFileDialog)
        layout.addWidget(self.upload_button)

        # Create a container for the image area and counter
        image_container_layout = QVBoxLayout()
        
        # Add image counter label
        counter_layout = QHBoxLayout()
        counter_layout.addStretch()
        self.image_counter_label = QLabel("0 Bilder", self)
        self.image_counter_label.setStyleSheet("color: #888888; font-weight: bold;")
        counter_layout.addWidget(self.image_counter_label)
        image_container_layout.addLayout(counter_layout)
        
        # Add the image area
        self.image_area = QScrollArea()
        self.image_container = QWidget()
        self.image_layout = QGridLayout()
        self.image_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.image_container.setLayout(self.image_layout)
        self.image_area.setWidget(self.image_container)
        self.image_area.setWidgetResizable(True)
        image_container_layout.addWidget(self.image_area)
        
        layout.addLayout(image_container_layout)

        self.empty_label = QLabel("Importiere Fotos oder ziehe sie hierhin", self.image_container)
        self.empty_label.setStyleSheet("color: #888888;")
        self.empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_layout = QVBoxLayout()
        self.empty_layout.addStretch()
        self.empty_layout.addWidget(self.empty_label)
        self.empty_layout.addStretch()
        self.image_layout.addLayout(self.empty_layout, 0, 0)

        pdf_buttons_layout = QHBoxLayout()
        pdf_buttons_layout.setContentsMargins(150, 10, 150, 0)
        pdf_buttons_layout.setSpacing(10)

        self.preview_original_button = QPushButton("Vorschau (Originale)", self)
        self.preview_original_button.setEnabled(False)
        self.preview_original_button.clicked.connect(self.createPreviewOriginal)
        self.preview_original_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        pdf_buttons_layout.addWidget(self.preview_original_button)

        self.pdf_button = QPushButton("PDF erstellen", self)
        self.pdf_button.setEnabled(False)
        self.pdf_button.clicked.connect(self.createPDF)
        self.pdf_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        pdf_buttons_layout.addWidget(self.pdf_button)

        self.reset_button = QPushButton("Zurücksetzen", self)
        self.reset_button.setStyleSheet("color: #ff453a;")
        self.reset_button.clicked.connect(self.resetApp)
        self.reset_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        pdf_buttons_layout.addWidget(self.reset_button)

        layout.addLayout(pdf_buttons_layout)
        self.setLayout(layout)

        self.overlay = QLabel(self)
        self.overlay.setStyleSheet("background-color: rgba(255, 255, 255, 0.3);")
        self.overlay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.overlay.setText("Drag & Drop Here")
        self.overlay.setVisible(False)

    def rearrangeImages(self):
        for i in reversed(range(self.image_layout.count())):
            item = self.image_layout.itemAt(i)
            if item:
                widget = item.widget()
                if widget and widget != self.empty_label:
                    self.image_layout.removeWidget(widget)
                    widget.setParent(None)
        for idx, (frame, _, _) in enumerate(self.images):
            row = idx // 4
            col = idx % 4
            self.image_layout.addWidget(frame, row, col)
        self.empty_label.setVisible(len(self.images) == 0)

    def showProgress(self, maximum, text):
        progress_dialog = QProgressDialog(text, "Cancel", 0, maximum, self)
        progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        progress_dialog.setMinimumDuration(0)
        progress_dialog.setAutoClose(False)
        progress_dialog.setAutoReset(False)
        progress_dialog.show()
        return progress_dialog

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

    def resizeEvent(self, event):
        self.overlay.setGeometry(self.rect())

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.overlay.setVisible(True)
            self.raise_()
            self.activateWindow()

    def dragLeaveEvent(self, event):
        self.overlay.setVisible(False)

    def dropEvent(self, event):
        self.overlay.setVisible(False)
        if event.mimeData().hasUrls():
            valid_files = []
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp")):
                    valid_files.append(file_path)
            if valid_files:
                self.startImageImport(valid_files)
        event.acceptProposedAction()

    def openFileDialog(self):
        homedir = os.environ.get('HOME', '')
        dialog = QFileDialog()
        files, _ = dialog.getOpenFileNames(
            parent=self,
            caption="Bilder auswählen",
            directory=homedir,
            filter="Bilder (*.png *.jpg *.jpeg *.bmp);;Alle Dateien (*)"
        )
        if files:
            self.startImageImport(files)

    def startImageImport(self, file_paths):
        self.import_progress_dialog = self.showProgress(len(file_paths), "Importing images...")
        self.import_worker = ImageImportWorker(file_paths)
        self.import_worker.imageImported.connect(self.addImageFromWorker)
        self.import_worker.progress.connect(lambda val: self.import_progress_dialog.setValue(val))
        self.import_progress_dialog.canceled.connect(self.import_worker.cancel)
        self.import_worker.finished.connect(self.importFinished)
        self.import_worker.start()

    def updateImageCounter(self):
        count = len(self.images)
        self.image_counter_label.setText(f"{count} {'Bild' if count == 1 else 'Bilder'}")

    def importFinished(self):
        self.import_progress_dialog.close()
        self.updateImageCounter()  # Update the counter when import is finished

    def addImageFromWorker(self, file_path, qimage):
        pixmap = QPixmap.fromImage(qimage)
        frame = QFrame(self.image_container)
        container = QWidget(frame)
        layout = QGridLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Generate a unique ID for this image instance
        import uuid
        unique_id = str(uuid.uuid4())
        
        label = DraggableLabel(container, file_path, self)
        label.unique_id = unique_id  # Now this will work without warnings
        scaled_pixmap = pixmap.scaled(120, 120, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        label.setPixmap(scaled_pixmap)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        remove_button = QPushButton("✖", container)
        remove_button.setFixedSize(16, 16)
        remove_button.clicked.connect(lambda: self.removeImage(frame, file_path, unique_id))
        remove_button.setStyleSheet("""
            QPushButton {
                background-color: rgb(102, 102, 102);
                color: white;
                border-radius: 8px;
                font-size: 11px;
                border: none;
            }
            QPushButton:hover {
                background-color: rgb(150, 150, 150);
            }
        """)
        layout.addWidget(label, 0, 0, Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(remove_button, 0, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(0, 0, 0, 0)
        frame_layout.addWidget(container)
        self.images.append((frame, file_path, unique_id))
        self.rearrangeImages()
        self.updatePdfButtonState()
        self.updateImageCounter()

    def reorderImages(self, source_path, target_path):
        source_index = -1
        target_index = -1
        # Update the tuple unpacking to handle all three values
        for i, (_, path, _) in enumerate(self.images):
            if path == source_path:
                source_index = i
            if path == target_path:
                target_index = i
        if source_index != -1 and target_index != -1:
            item = self.images.pop(source_index)
            self.images.insert(target_index, item)
            self.rearrangeImages()

    def removeImage(self, frame, file_path, unique_id=None):
        self.image_layout.removeWidget(frame)
        frame.deleteLater()
        
        if unique_id:
            # Remove only the specific image instance with this unique ID
            self.images = [img for img in self.images if img[2] != unique_id]
        else:
            # Backward compatibility for any code that might still call without unique_id
            self.images = [img for img in self.images if img[1] != file_path]
            
        self.rearrangeImages()
        self.updatePdfButtonState()
        self.updateImageCounter()

    def resetApp(self):
        reply = QMessageBox.question(
            self,
            'Bestätigung',
            'Möchten Sie die App wirklich zurücksetzen?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.aktennummer_input.clear()
            self.dokumentenkürzel_input.setCurrentIndex(0)
            self.dokumentenzahl_input.clear()
            for frame, _, _ in self.images:  # Updated to unpack 3 values
                self.image_layout.removeWidget(frame)
                frame.deleteLater()
            self.images.clear()
            self.empty_label.setVisible(True)
            self.updatePdfButtonState()
            self.updateImageCounter()  # Update the counter when resetting

    def updatePdfButtonState(self):
        aktennummer_filled = bool(self.aktennummer_input.text().strip())
        dokumentenzahl_filled = bool(self.dokumentenzahl_input.text().strip())
        images_present = bool(self.images)
        self.pdf_button.setEnabled(aktennummer_filled and dokumentenzahl_filled and images_present)
        # Preview does not require aktennummer to be filled
        self.preview_original_button.setEnabled(images_present)

    def createPreviewOriginal(self):
        # Validate images
        if not self.images:
            QMessageBox.warning(self, "Keine Bilder", "Bitte fügen Sie mindestens ein Bild hinzu.")
            return

        try:
            start_photo_number = int(self.start_photo_number.text().strip())
        except ValueError:
            start_photo_number = 1

        # Collect image paths
        image_paths = [p for _, p, _ in self.images]

        # Group for progress calculation to match normal flow
        def is_horizontal(fp):
            try:
                with Image.open(fp) as im:
                    im = ImageOps.exif_transpose(im)
                    w, h = im.size
                return w >= h
            except:
                return False

        grouped = []
        i = 0
        n = len(image_paths)
        while i < n:
            if not is_horizontal(image_paths[i]):
                grouped.append([image_paths[i]])
                i += 1
            else:
                if i + 1 < n and is_horizontal(image_paths[i + 1]):
                    grouped.append([image_paths[i], image_paths[i + 1]])
                    i += 2
                else:
                    grouped.append([image_paths[i]])
                    i += 1
        total_images = sum(len(g) for g in grouped)

        # Show progress dialog
        self.pdf_progress_dialog = self.showProgress(total_images, "Creating PDF Preview...")
        briefkopf_path = self.resource_path(os.path.join('resources', 'briefkopf.png'))

        # Create temporary output folder and PDF path
        preview_temp_dir = tempfile.mkdtemp(prefix="pdf_preview_")
        pdf_path = os.path.join(preview_temp_dir, "preview.pdf")

        # Start worker with preview flags and reuse the same creation routine
        self.pdf_worker = PDFCreationWorker(
            image_paths, "", "", "",
            pdf_path, briefkopf_path, preview_temp_dir, start_photo_number,
            use_original_filenames=True, save_to_disk=False,
            copy_images_to_output_dir=False, open_preview_only=True,
            zip_images=False,
            add_timestamp=self.add_timestamp_checkbox.isChecked()
        )
        self.pdf_worker.progressUpdate.connect(lambda val: self.pdf_progress_dialog.setValue(val))
        self.pdf_worker.statusUpdate.connect(lambda msg: self.pdf_progress_dialog.setLabelText(msg))
        self.pdf_progress_dialog.canceled.connect(self.pdf_worker.cancel)
        # Connect to dedicated preview finished handler
        self.pdf_worker.finished.connect(lambda path, temp_dir=preview_temp_dir: self.previewFinished(path, temp_dir))
        self.pdf_worker.errorOccurred.connect(self.pdfError)
        self.pdf_worker.start()

    def createPDF(self):
        if not self.images:
            QMessageBox.warning(self, "Keine Bilder", "Bitte fügen Sie mindestens ein Bild hinzu.")
            return
        aktennummer = self.aktennummer_input.text().strip()
        dokumentenkürzel = self.dokumentenkürzel_input.currentText().strip()
        dokumentenzahl = self.dokumentenzahl_input.text().strip()
        
        # Get the starting photo number
        try:
            start_photo_number = int(self.start_photo_number.text().strip())
        except ValueError:
            start_photo_number = 1  # Default to 1 if invalid
            
        if not aktennummer or not dokumentenzahl:
            QMessageBox.warning(self, "Fehlende Eingaben", "Bitte füllen Sie alle Eingabefelder aus.")
            return
        if self.dokumentenkürzel_input.currentIndex() == 0:
            default_folder_name = f"{aktennummer}-{dokumentenzahl}"
        else:
            default_folder_name = f"{aktennummer}-{dokumentenkürzel}-{dokumentenzahl}"
        first_image_dir = os.path.dirname(self.images[0][1]) if self.images else ""
        if not first_image_dir:
            first_image_dir = os.path.expanduser("~")
        folder_suggestion = os.path.join(first_image_dir, default_folder_name)
        old_folder_existed = False
        temp_renamed_folder = ""
        if os.path.isdir(folder_suggestion):
            old_folder_existed = True
            temp_renamed_folder = folder_suggestion + "_temp"
            try:
                os.rename(folder_suggestion, temp_renamed_folder)
            except Exception as e:
                QMessageBox.critical(self, "Fehler", f"Konnte Ordner nicht umbenennen:\n{e}")
                return
        folder_path, _ = QFileDialog.getSaveFileName(
            self,
            "Zielordner wählen (einfach Ordnername eingeben)",
            folder_suggestion,
            "Ordner (*)"
        )
        if not folder_path:
            if old_folder_existed and temp_renamed_folder and os.path.isdir(temp_renamed_folder):
                try:
                    os.rename(temp_renamed_folder, folder_suggestion)
                except:
                    pass
            return
        base_name = os.path.basename(folder_path)
        if base_name.lower().endswith(".pdf"):
            base_name = os.path.splitext(base_name)[0]
        parent_dir = os.path.dirname(folder_path)
        if not parent_dir:
            parent_dir = first_image_dir
        output_folder = os.path.join(parent_dir, base_name)
        if old_folder_existed:
            if output_folder == folder_suggestion:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("Ordner existiert bereits")
                msg_box.setText(f"Der Ordner\n\n{output_folder}\n\nexistiert bereits.\nMöchten Sie den Inhalt überschreiben?")
                msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                msg_box.setDefaultButton(QMessageBox.StandardButton.No)
                msg_box.setIcon(QMessageBox.Icon.Question)
                # Use the same icon as the main application
                icon_path = self.resource_path(os.path.join('resources', 'icon.png'))
                msg_box.setWindowIcon(QIcon(icon_path))
                reply = msg_box.exec()
                if reply == QMessageBox.StandardButton.Yes:
                    try:
                        shutil.rmtree(temp_renamed_folder)
                    except Exception as ex:
                        QMessageBox.critical(self, "Fehler", f"Konnte Ordner nicht löschen:\n{ex}")
                        try:
                            os.rename(temp_renamed_folder, folder_suggestion)
                        except:
                            pass
                        return
                else:
                    try:
                        os.rename(temp_renamed_folder, folder_suggestion)
                    except:
                        pass
                    return
            else:
                try:
                    os.rename(temp_renamed_folder, folder_suggestion)
                except:
                    pass
        if os.path.isdir(output_folder):
            reply = QMessageBox.question(
                self,
                "Ordner existiert bereits",
                f"Der Ordner\n\n{output_folder}\n\nexistiert bereits.\nMöchten Sie den Inhalt überschreiben?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
            else:
                try:
                    shutil.rmtree(output_folder)
                except Exception as ex:
                    QMessageBox.critical(self, "Fehler", f"Konnte Ordner nicht löschen:\n{ex}")
                    return
        os.makedirs(output_folder, exist_ok=True)
        pdf_path = os.path.join(output_folder, f"{base_name}.pdf")

        # Ask about timestamps for saved images if the toggle is active
        add_timestamp_to_saved_files = False
        if self.add_timestamp_checkbox.isChecked():
            reply = QMessageBox.question(
                self,
                "Zeitstempel",
                "Sollen die Bilder, die im neuen Ordner/ZIP gespeichert werden, auch mit einem Zeitstempel versehen werden?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                add_timestamp_to_saved_files = True

        # Update this line to extract file paths from the new image tuple structure
        image_paths = [p for _, p, _ in self.images]

        def is_horizontal(fp):
            try:
                with Image.open(fp) as im:
                    im = ImageOps.exif_transpose(im)
                    w, h = im.size
                return w >= h
            except:
                return False

        grouped = []
        i = 0
        n = len(image_paths)
        while i < n:
            if not is_horizontal(image_paths[i]):
                grouped.append([image_paths[i]])
                i += 1
            else:
                if i + 1 < n and is_horizontal(image_paths[i + 1]):
                    grouped.append([image_paths[i], image_paths[i + 1]])
                    i += 2
                else:
                    grouped.append([image_paths[i]])
                    i += 1
        total_images = sum(len(g) for g in grouped)
        self.pdf_progress_dialog = self.showProgress(total_images, "Creating PDF...")
        briefkopf_path = self.resource_path(os.path.join('resources', 'briefkopf.png'))
        self.pdf_worker = PDFCreationWorker(image_paths, aktennummer, dokumentenkürzel,
                                             dokumentenzahl, pdf_path, briefkopf_path, 
                                             output_folder, start_photo_number,
                                             zip_images=self.save_images_as_zip_checkbox.isChecked(),
                                             delete_originals=self.delete_originals_checkbox.isChecked(),
                                             add_timestamp=self.add_timestamp_checkbox.isChecked(),
                                             add_timestamp_to_saved_files=add_timestamp_to_saved_files)
        self.pdf_worker.progressUpdate.connect(lambda val: self.pdf_progress_dialog.setValue(val))
        self.pdf_worker.statusUpdate.connect(lambda msg: self.pdf_progress_dialog.setLabelText(msg))
        self.pdf_progress_dialog.canceled.connect(self.pdf_worker.cancel)
        self.pdf_worker.finished.connect(self.pdfFinished)
        self.pdf_worker.errorOccurred.connect(self.pdfError)
        self.pdf_worker.start()

    def pdfFinished(self, save_path):
        self.pdf_progress_dialog.close()
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("PDF erstellt")
        msg_box.setText(f"PDF erfolgreich erstellt:\n{save_path}")
        msg_box.setIcon(QMessageBox.Icon.Information)
        # Use the same icon as the main application
        icon_path = self.resource_path(os.path.join('resources', 'icon.png'))
        msg_box.setWindowIcon(QIcon(icon_path))
        msg_box.exec()
        
        if sys.platform == "win32":
            os.startfile(save_path)
        elif sys.platform == "darwin":
            subprocess.run(["open", save_path])
        else:
            subprocess.run(["xdg-open", save_path])

    def pdfError(self, error_message):
        self.pdf_progress_dialog.close()
        QMessageBox.critical(self, "Fehler", f"Beim Erstellen der PDF ist ein Fehler aufgetreten:\n{error_message}")

    def previewFinished(self, save_path, temp_dir):
        # Close progress and open in Preview only, then clean up temporary resources
        self.pdf_progress_dialog.close()

        def _open_and_cleanup(path, dir_path):
            try:
                if sys.platform == "darwin":
                    subprocess.run(["open", "-W", "-a", "Preview", path])
                elif sys.platform == "win32":
                    os.startfile(path)
                else:
                    subprocess.run(["xdg-open", path])
            finally:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception:
                    pass
                try:
                    if os.path.isdir(dir_path):
                        shutil.rmtree(dir_path)
                except Exception:
                    pass

        t = threading.Thread(target=_open_and_cleanup, args=(save_path, temp_dir), daemon=True)
        t.start()


__all__ = ["ImageUploader", "DraggableLabel"]