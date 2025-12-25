import flet as ft
import os
from io import BytesIO
from datetime import datetime
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)




supported_extensions = [".txt", ".xlsx", ".pptx", ".pptm", ".ppt",".pdf",".docx"]
image_extensions = ['.png', '.jpg', '.jpeg']







def main(page: ft.Page):
    page.title = "File to PDF Converter"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER

    selected_file = {"path": None}  
    label_file_explorer = ft.Text("No file selected", color=ft.Colors.GREY)

    image_preview = ft.Image(
    src=None,
    width=250,
    height=250,
    fit=ft.ImageFit.CONTAIN,
    visible=False,
    )

    result_preview = ft.Image(
        src=None,
        width=250,
        height=250,
        fit=ft.ImageFit.CONTAIN,
        visible=False,
    )

    
    def file_picker_result(e: ft.FilePickerResultEvent):
        if not e.files:
            label_file_explorer.value = "No file selected"
            label_file_explorer.color = ft.Colors.GREY
            selected_file["path"] = None
            selected_file["files"] = []  
            
            convert_button.disabled = True
            background_remove.disabled = True
            
            image_preview.visible = False
            result_preview.visible = False
            
            image_to_pdf.disabled = True
            Merge.disabled = True
            page.update()
            
            return

        
        selected_file["files"] = [f.path for f in e.files]

        unsupported_files = []
        supported_files = []
        image_files = []

        for file in e.files:
            file_ext = os.path.splitext(file.name)[1].lower()
            if file_ext in supported_extensions:
                supported_files.append(file)
            elif file_ext in image_extensions:
                image_files.append(file)
            else:
                unsupported_files.append(file.name)

        if unsupported_files:
            label_file_explorer.value = f"⚠ Unsupported files:\n" + "\n".join(unsupported_files)
            label_file_explorer.color = ft.Colors.ORANGE
        else:
            file_count = len(e.files)
            label_file_explorer.value = f"✓ {file_count} File(s) Selected:\n" + "\n".join([f.name for f in e.files[:3]])
            if file_count > 3:
                label_file_explorer.value += f"\n... and {file_count - 3} more"
            label_file_explorer.color = ft.Colors.GREEN

        
        if supported_files:
            
            selected_file["files"] = [f.path for f in supported_files]
            convert_button.disabled = False
            Merge.disabled = False
            
            background_remove.disabled = True
            image_to_pdf.disabled = True
            
        elif image_files:
            
            selected_file["files"] = [f.path for f in image_files]
            image_preview.src = image_files[0].path
            
            image_preview.visible = True
            result_preview.visible = False
            
            convert_button.disabled = True
            Merge.disabled = True
            
            background_remove.disabled = False
            image_to_pdf.disabled = False
        else:
            convert_button.disabled = True
            background_remove.disabled = True
            image_to_pdf.disabled = True
            Merge.disabled = True

        page.update()
        print(f"Files selected: {selected_file['files']}")
    
    def merge_pdfs(e):
        from PyPDF2 import PdfReader,PdfFileWriter

        filenames = selected_file["files"]

        if not filenames or len(filenames) < 2:
            label_file_explorer.value = "Please select at least 2 PDFs to merge"
            label_file_explorer.color = ft.Colors.RED
            page.update()
            return

        try:
            pdf_writer = PdfWriter()

            for pdf_file in filenames:
                reader = PdfReader(pdf_file)
                for pdf_page in reader.pages:   
                    pdf_writer.add_page(pdf_page)

            date_str = datetime.now().strftime("%Y-%m-%d")
            output_path = rf"C:\Users\Dwayne\Downloads\MergedPDF_{date_str}.pdf"

            with open(output_path, "wb") as f:
                pdf_writer.write(f)

            label_file_explorer.value = f"✓ PDFs merged successfully:\n{output_path}"
            label_file_explorer.color = ft.Colors.GREEN
            page.update()

        except Exception as err:
            label_file_explorer.value = f"Error merging PDFs: {err}"
            label_file_explorer.color = ft.Colors.RED
            page.update()

    def background_removal(e):
        from PIL import Image
        from backgroundremover.bg import remove
        import importlib.metadata

        _real_version = importlib.metadata.version

        def safe_version(name):
            if name == "imageio":
                return "2.34.0"
            return _real_version(name)

        importlib.metadata.version = safe_version

        filenames = selected_file["files"]
        
        for filename in filenames:
            file_ext = os.path.splitext(filename)[1].lower()

        if not filename:
            label_file_explorer.value = "Please select an image first"
            label_file_explorer.color = ft.Colors.RED
            page.update()
            return

        try:
           
            with open(filename, "rb") as f:
                image_bytes = f.read()

          
            result_bytes = remove(image_bytes)

           
            image_no_bg = Image.open(BytesIO(result_bytes))

            output_path = os.path.splitext(filename)[0] + "_no_bg.png"
            image_no_bg.save(output_path)

            result_preview.src = output_path
            result_preview.visible = True


            label_file_explorer.value = f"✓ Background removed:\n{output_path}"
            label_file_explorer.color = ft.Colors.GREEN
            page.update()

            print("Background removed successfully.")

        except Exception as err:
            label_file_explorer.value = f"Error removing background: {err}"
            label_file_explorer.color = ft.Colors.RED
            page.update()

    def img_to_pdf(e):
        from PIL import Image

        filenames = selected_file["files"]   

        for filename in filenames:
                file_ext = os.path.splitext(filename)[1].lower()     
        
        if not filenames:
            label_file_explorer.value = "Please select an image first"
            label_file_explorer.color = ft.Colors.RED
            page.update()
            return
        
        try:
            image = Image.open(filename)
            
           
          
            base_name = os.path.splitext(os.path.basename(filename))[0]
            pdf_path = rf'C:\Users\Dwayne\Downloads\PDF\{base_name}.pdf'
            
           
            image.save(pdf_path, "PDF", resolution=100.0)
            
            label_file_explorer.value = f"PDF saved successfully to {pdf_path}"
            label_file_explorer.color = ft.Colors.GREEN
            page.update()
        
        except Exception as err:
            label_file_explorer.value = f"Error: {err}"
            label_file_explorer.color = ft.Colors.RED
            page.update()
            print("An error occurred:", err)

    def convert_to_pdf(e):
        filenames = selected_file["files"]

        for filename in filenames:
            file_ext = os.path.splitext(filename)[1].lower()

            if file_ext == ".xlsx":
                convert_excel_to_pdf(filename)
            elif file_ext in [".pptx", ".pptm", ".ppt"]:
                convert_powerpoint_to_pdf(filename)
            elif file_ext == ".docx":
                Convert_word_to_pdf(filename)

    def convert_excel_to_pdf(filename):
        import pythoncom
        from win32com import client

        excel = None
        workbook = None

        try:
            pythoncom.CoInitialize()

            excel = client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(filename)

            output_pdf = os.path.splitext(filename)[0] + ".pdf"

            workbook.ExportAsFixedFormat(
                Type=0,  
                Filename=output_pdf
            )

            label_file_explorer.value = f"✓ PDF created:\n{output_pdf}"
            label_file_explorer.color = ft.Colors.GREEN
            convert_button.disabled = False
            page.update()

        except Exception as err:
            label_file_explorer.value = f"Error: {err}"
            label_file_explorer.color = ft.Colors.RED
            convert_button.disabled = False
            page.update()
            print("Excel conversion error:", err)

        finally:
            try:
                if workbook:
                    workbook.Close(False)
                if excel:
                    excel.Quit()
                pythoncom.CoUninitialize()
            except:
                pass

    def convert_powerpoint_to_pdf(filename):
        import pythoncom
        from win32com import client

        try:
            pythoncom.CoInitialize()
            PowerPoint = client.Dispatch("PowerPoint.Application")
            ppt = PowerPoint.Presentations.Open(filename, WithWindow=False)
            output_pdf = os.path.splitext(filename)[0] + ".pdf"

            ppt.SaveAs(output_pdf, 32)
            ppt.Close()
            PowerPoint.Quit()
            pythoncom.CoUninitialize()

            label_file_explorer.value = f"✓ PDF created:\n{output_pdf}"
            label_file_explorer.color = ft.Colors.GREEN
            convert_button.disabled = False
            page.update()

            print(f"PowerPoint file converted to PDF: {output_pdf}")

        except Exception as e:
            label_file_explorer.value = f"Error: {e}"
            label_file_explorer.color = ft.Colors.RED
            convert_button.disabled = False
            page.update()
            print("PowerPoint conversion error:", e)

    def Convert_word_to_pdf(filename):
        import pythoncom
        from win32com import client

        try:
            pythoncom.CoInitialize()
            text = client.Dispatch("Word.Application")
            doc = text.Documents.Open(filename)
            output_pdf = os.path.splitext(filename)[0] + ".pdf"

            doc.SaveAs(output_pdf, FileFormat=17)
            doc.Close()
            text.Quit()
            pythoncom.CoUninitialize()

            label_file_explorer.value = f"✓ PDF created:\n{output_pdf}"
            label_file_explorer.color = ft.Colors.GREEN
            convert_button.disabled = False
            page.update()

            print(f"WORD file converted to PDF: {output_pdf}")
        
        except Exception as e:
            label_file_explorer.value = f"Error: {e}"
            label_file_explorer.color = ft.Colors.RED
            convert_button.disabled = False
            page.update()
            print("Word conversion error:", e)

    def edit_pdf_clicked(e):
        import fitz
        filenames = selected_file["files"]
        
        user_choice_field = ft.TextField(
            label="Text to Replace",
            hint_text="Enter the word you want to change"
        )
        new_text_field = ft.TextField(
            label="New Word",
            hint_text="Enter the new word"
        )
        
        def perform_replacement(e):
            old_text = user_choice_field.value
            new_text = new_text_field.value
            
            if not old_text or not new_text:
                label_file_explorer.value = "Please enter both old and new text"
                label_file_explorer.color = ft.Colors.RED
                page.update()
                return
            
            try:
                for filename in filenames:
                    file_ext = os.path.splitext(filename)[1].lower()
                    
                    if file_ext != '.pdf':
                        continue
                    
                    doc = fitz.open(filename)
                    
                   
                    for pdf_page in doc:
                    
                        blocks = pdf_page.get_text("dict")["blocks"]
                        
                        for block in blocks:
                           
                            if "lines" not in block:
                                continue
                            
                            for line in block["lines"]:
                                line_text = ""
                                line_rect = None
                                font_size = 12  
                                font_name = "helv"  
                                
                      
                                for span in line["spans"]:
                                    line_text += span["text"]
                                    
                                    
                                    if line_rect is None:
                                        line_rect = fitz.Rect(span["bbox"])
                                    else:
                                        line_rect |= fitz.Rect(span["bbox"])
                                    
                                    
                                    if span == line["spans"][0]:
                                        font_size = span["size"]
                                        font_name = span["font"]
                                
                         
                                if old_text in line_text:
                      
                                    new_line_text = line_text.replace(old_text, new_text)
                                    
                                  
                                    pdf_page.add_redact_annot(line_rect, fill=(1, 1, 1))
                                    pdf_page.apply_redactions()
                                    
                                    try:
                                        pdf_page.insert_text(
                                            (line_rect.x0, line_rect.y1 - 2),  
                                            new_line_text,
                                            fontsize=font_size,
                                            fontname=font_name
                                        )
                                    except:
                                        pdf_page.insert_text(
                                            (line_rect.x0, line_rect.y1 - 2),
                                            new_line_text,
                                            fontsize=font_size,
                                            fontname="helv"
                                        )
                    
          
                    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_path = filename.replace('.pdf', f'_edited_{date_str}.pdf')
                    doc.save(output_path)
                    doc.close()
                
             
                popup.open = False
                label_file_explorer.value = f"✓ PDF edited successfully!\nSaved to: {output_path}"
                label_file_explorer.color = ft.Colors.GREEN
                page.update()
                
            except Exception as err:
                popup.open = False
                label_file_explorer.value = f"Error: {err}"
                label_file_explorer.color = ft.Colors.RED
                page.update()
                pass
        
        def close_dialog(e):
            popup.open = False
            page.update()
        
        popup = ft.AlertDialog(
            title=ft.Text("Edit PDF Text"),
            content=ft.Column([
                user_choice_field,
                new_text_field
            ], tight=True, height=150),
            actions=[
                ft.TextButton("Cancel", on_click=close_dialog),
                ft.TextButton("Replace", on_click=perform_replacement)
            ]
        )
        
        page.overlay.append(popup)
        popup.open = True
        page.update()




    convert_button = ft.FilledButton(
        "Convert to PDF",
        disabled=True,
        icon=ft.Icons.PICTURE_AS_PDF,
        on_click=convert_to_pdf
    )
    
    background_remove = ft.FilledButton(
        "Remove Background",
        disabled=True,
        icon=ft.Icons.AUTO_FIX_HIGH,
        on_click=background_removal
    )

    image_to_pdf = ft.FilledButton(
        "Image --> PDF",
        disabled =  True,
        icon  = ft.Icons.PICTURE_AS_PDF_ROUNDED,
        on_click = img_to_pdf
    )

    Merge = ft.FilledButton(
        "Merge PDF",
        disabled= True,
        icon = ft.Icons.PICTURE_AS_PDF_OUTLINED,
        on_click = merge_pdfs
    )
    
    edit_pdf = ft.FilledButton(
        "Edit PDF",
        disabled= False,
        icon = ft.Icons.PICTURE_AS_PDF_OUTLINED,
        on_click = edit_pdf_clicked
    )
    

    file_picker = ft.FilePicker(on_result=file_picker_result)
    page.overlay.append(file_picker)

    page.add(
        ft.Row(
            controls=[
                ft.Container(
                    on_click=lambda e: file_picker.pick_files(allow_multiple=True),
                    content=ft.Container(
                        ft.Column(
                            controls=[
                                ft.Container(
                                    content=ft.Text("Select A File", size=40, color=ft.Colors.GREY),
                                    margin=10,
                                    padding=10,
                                    alignment=ft.alignment.center,
                                    width=400,
                                    height=400,
                                    border_radius=50,
                                    on_click=lambda e: file_picker.pick_files(allow_multiple=True),
                                    
                                ),
                                label_file_explorer,
                            ],
                            alignment=ft.MainAxisAlignment.CENTER,
                        ),
                        padding=10,
                    ),
                    border=ft.border.all(1, ft.Colors.BLUE),
                    shadow=ft.BoxShadow(blur_radius=10, color=ft.Colors.BLACK12, spread_radius=5),
                    border_radius=50, 
                    margin=10,
                    padding=10,
                    alignment=ft.alignment.center,
                    expand=True
                )
            ],
            alignment=ft.MainAxisAlignment.CENTER,
        ),
    )
    page.add(
        ft.Row(
            controls=[convert_button, background_remove, image_to_pdf, Merge,edit_pdf],
            alignment=ft.MainAxisAlignment.CENTER,
        ),
    )
    page.add(
        ft.Row(
            controls=[
                ft.Text("Supported file types: .txt, .xlsx, .pptx, .pptm, .ppt, .png, .jpeg, .jpg"),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
        ),
    )
    page.add(
        ft.Row(
            controls=[
                ft.Column(
                    [
                        ft.Text("Original", weight=ft.FontWeight.BOLD),
                        image_preview,
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Column(
                    [
                        ft.Text("Result", weight=ft.FontWeight.BOLD),
                        result_preview,
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=40,
        )
    )

ft.app(target=main)
