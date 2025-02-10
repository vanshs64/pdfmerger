import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import fitz  # PyMuPDF
import os

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")
        self.root.geometry("600x500")  # Increased height to fit new buttons
        
        # Set custom icon
        icon_path = os.path.join(os.getcwd(), "favicon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        self.pdf_files = []

        # Create UI elements
        self.create_widgets()

    def create_widgets(self):
        # File Name Entry
        name_frame = ttk.Frame(self.root)
        name_frame.pack(pady=5, padx=10, fill=tk.X)

        ttk.Label(name_frame, text="Merged File Name:").pack(side=tk.LEFT, padx=5)

        # File name entry with dynamic resizing
        self.filename_entry = ttk.Entry(name_frame, width=40)
        self.filename_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.filename_entry.insert(0, "merged")  # Default name

        # Scrollbar for long filenames
        self.filename_scroll = ttk.Scrollbar(name_frame, orient="horizontal", command=self.filename_entry.xview)
        self.filename_entry.config(xscrollcommand=self.filename_scroll.set)
        self.filename_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        # Drag-and-Drop Label
        self.label = ttk.Label(self.root, text="Drag and drop PDF files here", anchor="center")
        self.label.pack(pady=10)

        # File List (Treeview)
        self.tree = ttk.Treeview(self.root, columns=("File"), show="headings", selectmode="browse")
        self.tree.heading("File", text="PDF Files")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Buttons for file management
        file_button_frame = ttk.Frame(self.root)
        file_button_frame.pack(pady=5)

        ttk.Button(file_button_frame, text="Browse Files", command=self.browse_files).grid(row=0, column=0, padx=5)
        ttk.Button(file_button_frame, text="Remove Selected", command=self.remove_selected).grid(row=0, column=1, padx=5)
        ttk.Button(file_button_frame, text="Clear All", command=self.clear_all).grid(row=0, column=2, padx=5)

        # Buttons for moving and merging files
        action_button_frame = ttk.Frame(self.root)
        action_button_frame.pack(pady=5)

        ttk.Button(action_button_frame, text="▲ Move Up", command=self.move_up).grid(row=0, column=0, padx=5)
        ttk.Button(action_button_frame, text="▼ Move Down", command=self.move_down).grid(row=0, column=1, padx=5)
        ttk.Button(action_button_frame, text="Merge PDFs", command=self.merge_pdfs).grid(row=0, column=2, padx=5)

        # Enable Drag-and-Drop
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.drop_files)

    def browse_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.tree.insert("", tk.END, values=(os.path.basename(file),))

    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            if file.lower().endswith(".pdf"):
                if file not in self.pdf_files:
                    self.pdf_files.append(file)
                    self.tree.insert("", tk.END, values=(os.path.basename(file),))

    def remove_selected(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            del self.pdf_files[index]  # Remove from list
            self.tree.delete(selected_item)  # Remove from UI

    def clear_all(self):
        self.pdf_files.clear()  # Clear the file list
        self.tree.delete(*self.tree.get_children())  # Clear UI

    def move_up(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            if index > 0:
                self.pdf_files[index], self.pdf_files[index - 1] = self.pdf_files[index - 1], self.pdf_files[index]
                self.refresh_treeview()
                self.tree.selection_set(self.tree.get_children()[index - 1])

    def move_down(self):
        selected_item = self.tree.selection()
        if selected_item:
            index = self.tree.index(selected_item)
            if index < len(self.pdf_files) - 1:
                self.pdf_files[index], self.pdf_files[index + 1] = self.pdf_files[index + 1], self.pdf_files[index]
                self.refresh_treeview()
                self.tree.selection_set(self.tree.get_children()[index + 1])

    def refresh_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for file in self.pdf_files:
            self.tree.insert("", tk.END, values=(os.path.basename(file),))

    def merge_pdfs(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDFs to merge!")
            return

        filename = self.filename_entry.get().strip()
        if not filename:
            messagebox.showerror("Error", "Please enter a file name!")
            return
        
        filename = filename if filename.endswith(".pdf") else f"{filename}.pdf"

        # Default save location: Same folder as the script
        save_path = os.path.join(os.getcwd(), filename)

        # Allow user to change save location
        save_path = filedialog.asksaveasfilename(initialfile=filename, defaultextension=".pdf",
                                                 filetypes=[("PDF files", "*.pdf")])
        if not save_path:
            return

        merged_pdf = fitz.open()
        for pdf in self.pdf_files:
            with fitz.open(pdf) as doc:
                for page in doc:
                    merged_pdf.insert_pdf(doc, from_page=page.number, to_page=page.number)

        merged_pdf.save(save_path)
        merged_pdf.close()
        messagebox.showinfo("Success", f"PDFs merged successfully as '{os.path.basename(save_path)}'!")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
