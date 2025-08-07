import tkinter as tk
from tkinter import messagebox, filedialog
import os
import glob
from pathlib import Path
import threading
import sys
import subprocess

# Import conversion libraries
try:
    from docx import Document
    from docx2pdf import convert as docx_to_pdf
    from pypdf import PdfReader
except ImportError as e:
    print(f"Missing required library: {e}")
    print("Please install: pip install python-docx docx2pdf pypdf")

class DocPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("docpdf")
        self.root.geometry("400x200")
        self.root.resizable(False, False)
        
        # Get the directory where the exe is located
        if getattr(sys, 'frozen', False):
            # Running as compiled exe
            self.current_dir = os.path.dirname(sys.executable)
        else:
            # Running as Python script
            self.current_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.setup_ui()
    
    def setup_ui(self):
        # Title label
        title_label = tk.Label(
            self.root, 
            text="docpdf", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=20)
        
        # Directory info
        dir_label = tk.Label(
            self.root, 
            text=f"Working Directory: {os.path.basename(self.current_dir)}", 
            font=("Arial", 10)
        )
        dir_label.pack(pady=5)
        
        # Button frame
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        
        # Convert DOC/DOCX to PDF button
        self.doc_to_pdf_btn = tk.Button(
            button_frame,
            text="Convert DOC/DOCX → PDF",
            command=self.convert_doc_to_pdf,
            width=20,
            height=2,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold")
        )
        self.doc_to_pdf_btn.pack(side=tk.LEFT, padx=10)
        
        # Convert PDF to DOCX button
        self.pdf_to_doc_btn = tk.Button(
            button_frame,
            text="Convert PDF → DOCX",
            command=self.convert_pdf_to_docx,
            width=20,
            height=2,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10, "bold")
        )
        self.pdf_to_doc_btn.pack(side=tk.LEFT, padx=10)
        
        # Status label
        self.status_label = tk.Label(
            self.root, 
            text="Ready", 
            font=("Arial", 9),
            fg="green"
        )
        self.status_label.pack(pady=10)
    
    def update_status(self, message, color="black"):
        """Update status label with message and color"""
        self.status_label.config(text=message, fg=color)
        self.root.update()
    
    def convert_doc_to_pdf(self):
        """Convert all DOC/DOCX files in current directory to PDF"""
        def conversion_thread():
            try:
                self.update_status("Converting DOC/DOCX to PDF...", "blue")
                self.doc_to_pdf_btn.config(state="disabled")
                
                # Find all DOC and DOCX files (exclude temporary files)
                doc_files = glob.glob(os.path.join(self.current_dir, "*.doc"))
                docx_files = glob.glob(os.path.join(self.current_dir, "*.docx"))
                all_files = doc_files + docx_files
                
                # Filter out temporary files (starting with ~$)
                all_files = [f for f in all_files if not os.path.basename(f).startswith('~$')]
                
                if not all_files:
                    self.update_status("No DOC/DOCX files found", "orange")
                    messagebox.showinfo("Info", "No DOC or DOCX files found in the current directory.")
                    return
                
                converted_count = 0
                for file_path in all_files:
                    try:
                        # Create output PDF path
                        pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                        
                        # Convert using docx2pdf library
                        docx_to_pdf(file_path, pdf_path)
                        converted_count += 1
                        
                        self.update_status(f"Converted {converted_count}/{len(all_files)} files...", "blue")
                        
                    except Exception as e:
                        print(f"Error converting {file_path}: {e}")
                        continue
                
                self.update_status(f"Conversion complete! {converted_count} files converted", "green")
                messagebox.showinfo("Success", f"Successfully converted {converted_count} files to PDF!")
                
            except Exception as e:
                self.update_status("Conversion failed", "red")
                messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            
            finally:
                self.doc_to_pdf_btn.config(state="normal")
        
        # Run conversion in separate thread to prevent GUI freezing
        thread = threading.Thread(target=conversion_thread)
        thread.daemon = True
        thread.start()
    

    
    def convert_pdf_to_docx(self):
        """Convert all PDF files in current directory to DOCX"""
        def conversion_thread():
            try:
                self.update_status("Converting PDF to DOCX...", "blue")
                self.pdf_to_doc_btn.config(state="disabled")
                
                # Find all PDF files
                pdf_files = glob.glob(os.path.join(self.current_dir, "*.pdf"))
                
                if not pdf_files:
                    self.update_status("No PDF files found", "orange")
                    messagebox.showinfo("Info", "No PDF files found in the current directory.")
                    return
                
                converted_count = 0
                for pdf_path in pdf_files:
                    try:
                        # Create output DOCX path
                        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
                        
                        # Extract text from PDF using pypdf
                        reader = PdfReader(pdf_path)
                        text_content = ""
                        
                        for page in reader.pages:
                            text_content += page.extract_text()
                            text_content += "\n\n"  # Add page break
                        
                        # Create DOCX document
                        docx_doc = Document()
                        
                        # Split text into paragraphs and add to document
                        paragraphs = text_content.split('\n\n')
                        for paragraph in paragraphs:
                            if paragraph.strip():
                                docx_doc.add_paragraph(paragraph.strip())
                        
                        docx_doc.save(docx_path)
                        
                        converted_count += 1
                        self.update_status(f"Converted {converted_count}/{len(pdf_files)} files...", "blue")
                        
                    except Exception as e:
                        print(f"Error converting {pdf_path}: {e}")
                        continue
                
                self.update_status(f"Conversion complete! {converted_count} files converted", "green")
                messagebox.showinfo("Success", f"Successfully converted {converted_count} files to DOCX!")
                
            except Exception as e:
                self.update_status("Conversion failed", "red")
                messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            
            finally:
                self.pdf_to_doc_btn.config(state="normal")
        
        # Run conversion in separate thread to prevent GUI freezing
        thread = threading.Thread(target=conversion_thread)
        thread.daemon = True
        thread.start()

def main():
    root = tk.Tk()
    app = DocPDF(root)
    root.mainloop()

if __name__ == "__main__":
    main()