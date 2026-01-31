import os
import shutil
import tempfile
import zipfile
from io import BytesIO
from PIL import Image, ImageOps, ImageDraw, ImageFont
from PyQt6.QtCore import QThread, pyqtSignal
from datetime import datetime
from PyQt6.QtGui import QImage
from fpdf import FPDF

os.environ["LANG"] = "de_DE.UTF-8"


class ImageImportWorker(QThread):
    imageImported = pyqtSignal(str, QImage)
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self._isCanceled = False

    def cancel(self):
        self._isCanceled = True

    def run(self):
        count = 0
        for file_path in self.file_paths:
            if self._isCanceled:
                break
            try:
                with Image.open(file_path) as img:
                    img = ImageOps.exif_transpose(img)
                    if img.mode in ("RGBA", "LA"):
                        img = img.convert("RGB")
                    buf = BytesIO()
                    img.save(buf, format='JPEG')
                    qimg = QImage.fromData(buf.getvalue())
                self.imageImported.emit(file_path, qimg)
            except Exception as err:
                print(f"Fehler beim Importieren {file_path}: {err}")
            count += 1
            self.progress.emit(count)
        self.finished.emit()


class PDFCreationWorker(QThread):
    progressUpdate = pyqtSignal(int)
    statusUpdate = pyqtSignal(str)  # For status messages like "Saving PDF..."
    finished = pyqtSignal(str)
    errorOccurred = pyqtSignal(str)

    def __init__(self, image_paths, aktennummer, dokumentenkürzel, dokumentenzahl,
                 pdf_path, briefkopf_path, output_folder, start_photo_number=1,
                 use_original_filenames=False, save_to_disk=True,
                 copy_images_to_output_dir=True, open_preview_only=False,
                 zip_images=False, delete_originals=False, add_timestamp=False,
                 add_timestamp_to_saved_files=False):
        super().__init__()
        self.image_paths = image_paths
        self.aktennummer = aktennummer
        self.dokumentenkürzel = dokumentenkürzel
        self.dokumentenzahl = dokumentenzahl
        self.save_path = pdf_path
        self.briefkopf_path = briefkopf_path
        self.output_folder = output_folder
        self.start_photo_number = start_photo_number
        self.use_original_filenames = use_original_filenames
        self.save_to_disk = save_to_disk
        self.copy_images_to_output_dir = copy_images_to_output_dir
        self.open_preview_only = open_preview_only
        self.zip_images = zip_images
        self.delete_originals = delete_originals
        self.add_timestamp = add_timestamp
        self.add_timestamp_to_saved_files = add_timestamp_to_saved_files
        self._isCanceled = False
        self._temp_processing_dir = None
        self._zipf = None
        self.zip_path = None
        self._copied_originals = []
        self._total_images = len(image_paths)

    def cancel(self):
        self._isCanceled = True

    def format_image_number(self, number):
        if self._total_images <= 9:
            return str(number)
        elif self._total_images <= 99:
            return f"{number:02d}"
        else:
            return f"{number:03d}"

    def get_exif_datetime(self, file_path):
        try:
            with Image.open(file_path) as img:
                exif_data = img._getexif()
                if exif_data:
                    for tag_id in [36867, 36868, 306]:
                        if tag_id in exif_data:
                            dt_str = exif_data[tag_id]
                            try:
                                dt = datetime.strptime(dt_str, "%Y:%m:%d %H:%M:%S")
                                return dt.strftime("%d.%m.%Y %H:%M")
                            except ValueError:
                                continue
        except Exception as e:
            print(f"Error reading EXIF datetime: {e}")
        return None

    def add_timestamp_overlay(self, img, timestamp_text):
        draw = ImageDraw.Draw(img)
        
        font_size = max(28, int(img.width * 0.04))
        
        try:
            font_paths = [
                "/System/Library/Fonts/Courier.dfont",
                "/System/Library/Fonts/Monaco.dfont",
                "/Library/Fonts/Courier New Bold.ttf",
                "/Library/Fonts/Courier New.ttf",
                "C:\\Windows\\Fonts\\courbd.ttf",
                "C:\\Windows\\Fonts\\consolab.ttf",
                "C:\\Windows\\Fonts\\consola.ttf",
                "C:\\Windows\\Fonts\\cour.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
            ]
            font = None
            for font_path in font_paths:
                try:
                    font = ImageFont.truetype(font_path, font_size)
                    break
                except (OSError, IOError):
                    continue
            if font is None:
                font = ImageFont.load_default()
        except Exception:
            font = ImageFont.load_default()
        
        bbox = draw.textbbox((0, 0), timestamp_text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        padding = int(img.width * 0.02)
        x = img.width - text_width - padding
        y = img.height - text_height - padding
        
        outline_color = (0, 0, 0)
        outline_range = 2
        for dx in range(-outline_range, outline_range + 1):
            for dy in range(-outline_range, outline_range + 1):
                if dx != 0 or dy != 0:
                    draw.text((x + dx, y + dy), timestamp_text, font=font, fill=outline_color)
        
        text_color = (255, 165, 0)
        draw.text((x, y), timestamp_text, font=font, fill=text_color)
        
        return img

    def is_horizontal(self, file_path):
        try:
            with Image.open(file_path) as img:
                img = ImageOps.exif_transpose(img)
                width, height = img.size
            return width >= height
        except Exception as e:
            print("Error in is_horizontal:", e)
            return False

    def processImage(self, file_path, image_counter):
        try:
            with Image.open(file_path) as img:
                # Apply EXIF orientation to the original image first
                img_raw = ImageOps.exif_transpose(img.copy())
                
                # Create a copy for PDF processing
                img = img_raw.copy()
                if img.mode in ("RGBA", "LA"):
                    img = img.convert("RGB")
                    img_raw = img_raw.convert("RGB")  # Also convert img_raw if needed

                width, height = img.size
                aspect_ratio = width / height

                if width >= height and 1.0 <= aspect_ratio <= 1.33:
                    new_width = width
                    new_height = int(width * 3 / 4)
                    if new_height > height:
                        new_height = height
                        new_width = int(height * 4 / 3)
                    left = (width - new_width) / 2
                    top = (height - new_height) / 2
                    right = left + new_width
                    bottom = top + new_height
                    img = img.crop((left, top, right, bottom))

                max_side = max(img.width, img.height)
                if max_side > 2000:
                    scale_factor = 2000 / max_side
                    new_width = int(img.width * scale_factor)
                    new_height = int(img.height * scale_factor)
                    img = img.resize((new_width, new_height), Image.LANCZOS)

                file_extension = os.path.splitext(file_path)[1]

                # Save the properly oriented image either to disk or zip if requested
                if self.copy_images_to_output_dir:
                    formatted_num = self.format_image_number(image_counter)
                    if self.dokumentenkürzel.startswith("("):
                        image_filename = f"{self.aktennummer}-{self.dokumentenzahl} Foto Nr. {formatted_num}{file_extension}"
                    else:
                        image_filename = f"{self.aktennummer}-{self.dokumentenkürzel}-{self.dokumentenzahl} Foto Nr. {formatted_num}{file_extension}"

                    # Logic for timestamping saved images
                    timestamp_text_saved = None
                    if self.add_timestamp and self.add_timestamp_to_saved_files:
                        timestamp_text_saved = self.get_exif_datetime(file_path)

                    if self.zip_images:
                        if self._zipf is None:
                            # Derive zip name from PDF save_path base name
                            base_name = os.path.splitext(os.path.basename(self.save_path))[0]
                            self.zip_path = os.path.join(self.output_folder, f"{base_name} Bilder.zip")
                            self._zipf = zipfile.ZipFile(self.zip_path, mode='w', compression=zipfile.ZIP_STORED)

                        if timestamp_text_saved:
                            img_to_save = img_raw.copy()
                            if img_to_save.mode in ("RGBA", "LA"):
                                img_to_save = img_to_save.convert("RGB")
                            img_to_save = self.add_timestamp_overlay(img_to_save, timestamp_text_saved)
                            buf = BytesIO()
                            img_to_save.save(buf, format='JPEG', quality=95, subsampling=0)
                            self._zipf.writestr(image_filename, buf.getvalue())
                        else:
                            self._zipf.write(file_path, image_filename)
                        
                        self._copied_originals.append(file_path)
                    else:
                        final_path = os.path.join(self.output_folder, image_filename)
                        if timestamp_text_saved:
                            img_to_save = img_raw.copy()
                            if img_to_save.mode in ("RGBA", "LA"):
                                img_to_save = img_to_save.convert("RGB")
                            img_to_save = self.add_timestamp_overlay(img_to_save, timestamp_text_saved)
                            img_to_save.save(final_path, quality=95, subsampling=0)
                        else:
                            shutil.copy2(file_path, final_path)
                        
                        self._copied_originals.append(file_path)

                if self.copy_images_to_output_dir:
                    temp_dir = os.path.join(self.output_folder, "temp")
                else:
                    if self._temp_processing_dir is None:
                        self._temp_processing_dir = tempfile.mkdtemp(prefix="pdf_temp_")
                    temp_dir = self._temp_processing_dir
                os.makedirs(temp_dir, exist_ok=True)
                
                if self.add_timestamp:
                    timestamp_text = self.get_exif_datetime(file_path)
                    if timestamp_text:
                        img = self.add_timestamp_overlay(img, timestamp_text)
                
                temp_path = os.path.join(temp_dir, f"temp_{image_counter}{file_extension}")
                img.save(temp_path, quality=85)
                
                return temp_path
        except Exception as e:
            print("Fehler bei processImage:", e)
            return file_path

    def run(self):
        try:
            # Check if there are no images to process
            if not self.image_paths or self._isCanceled:
                self.finished.emit(self.save_path)
                return
                
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.set_auto_page_break(False)
            page_width = 210
            page_height = 297
            margin_top_bottom = 10
            header_spacing = 10
            offset = 3
            spacing_between = 8
            text_line_height = 10

            with Image.open(self.briefkopf_path) as briefkopf_img:
                aspect_briefkopf = briefkopf_img.width / briefkopf_img.height


            briefkopf_width_in_pdf = page_width / 3
            briefkopf_height_in_pdf = briefkopf_width_in_pdf / aspect_briefkopf

            content_top = margin_top_bottom + briefkopf_height_in_pdf + header_spacing
            content_height = page_height - margin_top_bottom - content_top

            uniform_img_dim = (content_height - spacing_between - (2 * (offset + text_line_height))) / 1.5

            grouped = []
            i = 0
            n = len(self.image_paths)
            while i < n:
                if not self.is_horizontal(self.image_paths[i]):
                    grouped.append([self.image_paths[i]])
                    i += 1
                else:
                    if i + 1 < n and self.is_horizontal(self.image_paths[i + 1]):
                        grouped.append([self.image_paths[i], self.image_paths[i + 1]])
                        i += 2
                    else:
                        grouped.append([self.image_paths[i]])
                        i += 1


            progress_count = 0
            global_image_counter = self.start_photo_number  # Use the starting number
            
            for group in grouped:
                if self._isCanceled:
                    break

                pdf.add_page()
                x_briefkopf = (page_width - briefkopf_width_in_pdf) / 2
                y_briefkopf = margin_top_bottom
                pdf.image(self.briefkopf_path, x=x_briefkopf, y=y_briefkopf,
                          w=briefkopf_width_in_pdf, h=briefkopf_height_in_pdf)

                if len(group) == 1:
                    file_path = group[0]
                    processed_path = self.processImage(file_path, global_image_counter)
                    with Image.open(processed_path) as img:
                        orig_w, orig_h = img.size

                    new_width = uniform_img_dim
                    new_height = orig_h * (new_width / orig_w)

                    if new_height > content_height - 15:
                        new_height = content_height - 15
                        new_width = orig_w * (new_height / orig_h)

                    block_total_height = new_height + offset + text_line_height
                    y_block_top = (content_top + (content_height - block_total_height) / 2)
                    x_image = (page_width - new_width) / 2
                    y_image = y_block_top

                    pdf.image(processed_path, x=x_image, y=y_image, w=new_width, h=new_height)

                    pdf.set_font("Arial", "B", 11)

                    if self.use_original_filenames:
                        text = os.path.basename(file_path)
                    else:
                        formatted_num = self.format_image_number(global_image_counter)
                        if self.dokumentenkürzel.startswith("("):
                            text = f"{self.aktennummer}-{self.dokumentenzahl} Foto Nr. {formatted_num}"
                        else:
                            text = f"{self.aktennummer}-{self.dokumentenkürzel}-{self.dokumentenzahl} Foto Nr. {formatted_num}"
                    text_width = pdf.get_string_width(text)
                    x_text = (page_width - text_width) / 2
                    y_text = y_image + new_height + offset
                    pdf.set_xy(x_text, y_text)
                    pdf.cell(text_width, text_line_height, text, align="C")

                    global_image_counter += 1
                    progress_count += 1
                    self.progressUpdate.emit(progress_count)

                elif len(group) == 2:
                    file_path1, file_path2 = group
                    processed_path1 = self.processImage(file_path1, global_image_counter)
                    processed_path2 = self.processImage(file_path2, global_image_counter + 1)

                    with Image.open(processed_path1) as img1:
                        orig1_w, orig1_h = img1.size
                    with Image.open(processed_path2) as img2:
                        orig2_w, orig2_h = img2.size

                    pdf.set_font("Arial", "B", 11)
                    if self.use_original_filenames:
                        text1 = os.path.basename(file_path1)
                        text2 = os.path.basename(file_path2)
                    else:
                        formatted_num1 = self.format_image_number(global_image_counter)
                        formatted_num2 = self.format_image_number(global_image_counter + 1)
                        if self.dokumentenkürzel.startswith("("):
                            text1 = f"{self.aktennummer}-{self.dokumentenzahl} Foto Nr. {formatted_num1}"
                            text2 = f"{self.aktennummer}-{self.dokumentenzahl} Foto Nr. {formatted_num2}"
                        else:
                            text1 = f"{self.aktennummer}-{self.dokumentenkürzel}-{self.dokumentenzahl} Foto Nr. {formatted_num1}"
                            text2 = f"{self.aktennummer}-{self.dokumentenkürzel}-{self.dokumentenzahl} Foto Nr. {formatted_num2}"
                    text1_width = pdf.get_string_width(text1)
                    text2_width = pdf.get_string_width(text2)

                    new1_width = uniform_img_dim
                    new1_height = orig1_h * (new1_width / orig1_w)
                    new2_width = uniform_img_dim
                    new2_height = orig2_h * (new2_width / orig2_w)

                    if new1_height > (content_height - spacing_between) / 2:
                        new1_height = (content_height - spacing_between) / 2
                        new1_width = orig1_w * (new1_height / orig1_h)

                    if new2_height > (content_height - spacing_between) / 2:
                        new2_height = (content_height - spacing_between) / 2
                        new2_width = orig2_w * (new2_height / orig2_h)

                    block_total_height = (new1_height + offset + text_line_height) + spacing_between + (new2_height + offset + text_line_height)
                    y_block_top = content_top + (content_height - block_total_height) / 2

                    x1 = (page_width - new1_width) / 2
                    y1 = y_block_top
                    pdf.image(processed_path1, x=x1, y=y1, w=new1_width, h=new1_height)

                    x_text1 = (page_width - text1_width) / 2
                    y_text1 = y1 + new1_height + offset
                    pdf.set_xy(x_text1, y_text1)
                    pdf.cell(text1_width, text_line_height, text1, align="C")

                    y2 = y_text1 + text_line_height + spacing_between
                    x2 = (page_width - new2_width) / 2
                    pdf.image(processed_path2, x=x2, y=y2, w=new2_width, h=new2_height)

                    x_text2 = (page_width - text2_width) / 2
                    y_text2 = y2 + new2_height + offset
                    pdf.set_xy(x_text2, y_text2)
                    pdf.cell(text2_width, text_line_height, text2, align="C")

                    global_image_counter += 2
                    progress_count += 2
                    self.progressUpdate.emit(progress_count)

            if not self._isCanceled:
                # Emit status update before saving PDF (this can take a while with many images)
                self.statusUpdate.emit("PDF wird gespeichert...")
                
                # Always write to a path; UI will decide persistence
                pdf.output(self.save_path)

                # Clean up temporary processed images
                if self.copy_images_to_output_dir:
                    temp_dir = os.path.join(self.output_folder, "temp")
                else:
                    temp_dir = self._temp_processing_dir
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception as e:
                        print(f"Error cleaning up temp files: {e}")

                # Close zip archive if used
                if self._zipf is not None:
                    try:
                        self._zipf.close()
                    except Exception:
                        pass

                if self.delete_originals and self._copied_originals:
                    for original_path in self._copied_originals:
                        try:
                            if os.path.exists(original_path):
                                os.remove(original_path)
                        except Exception as e:
                            print(f"Error deleting original file {original_path}: {e}")

                self.finished.emit(self.save_path)

        except Exception as e:
            self.errorOccurred.emit(str(e))


__all__ = ["ImageImportWorker", "PDFCreationWorker"]

