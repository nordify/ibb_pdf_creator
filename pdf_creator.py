import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QTranslator, QLibraryInfo

# Public API re-exports
from ui import ImageUploader
from workers import ImageImportWorker, PDFCreationWorker

__all__ = [
    "ImageUploader",
    "ImageImportWorker",
    "PDFCreationWorker",
]


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("resources/icon.png"))
    translator = QTranslator()
    translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    translator.load("qtbase_de", translations_path)
    app.installTranslator(translator)
    ex = ImageUploader()
    ex.show()
    sys.exit(app.exec())