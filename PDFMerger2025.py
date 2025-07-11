import tkinter as tk
from tkinter import filedialog, messagebox, Menu, Toplevel
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import re
import os
from PyPDF2 import PdfMerger, PdfReader

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")
        self.root.geometry("700x670")  # Increased width for button visibility
        self.root.configure(bg="#f0f2f5")

        style = ttk.Style()
        style.theme_use('clam')

        style.configure("TButton",
                        padding=6,
                        relief="flat",
                        background="#4a90e2",
                        foreground="white",
                        font=("Segoe UI", 11, "bold"))
        style.map("TButton",
                  background=[('active', '#357ABD')])

        style.configure("Small.TButton",
                        padding=3,
                        relief="flat",
                        background="#4a90e2",
                        foreground="white",
                        font=("Segoe UI", 9, "bold"))
        style.map("Small.TButton",
                  background=[('active', '#357ABD')])

        style.configure("TLabel", background="#f0f2f5", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), background="#f0f2f5")

        menubar = Menu(self.root)
        helpmenu = Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.root.config(menu=menubar)

        header = ttk.Label(root, text="Select and order PDFs to merge:", style="Header.TLabel")
        header.pack(pady=(15, 5))

        frame_list = ttk.Frame(root)
        frame_list.pack(padx=20, pady=5, fill="both", expand=True)

        self.listbox = tk.Listbox(frame_list, selectmode=tk.EXTENDED, bg="white",
                                  font=("Segoe UI", 10), activestyle="none",
                                  highlightthickness=1, relief="solid", height=15)
        self.listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind("<<Drop>>", self.on_drop)

        self.listbox.bind("<<ListboxSelect>>", self.on_list_select)

        # Drag files here overlay label
        self.drag_hint = tk.Label(frame_list, text="Drag files here", fg="gray",
                                  bg="white", font=("Segoe UI", 12, "italic"))
        self.drag_hint.place(relx=0.5, rely=0.5, anchor="center")

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)

        self.btn_add = ttk.Button(btn_frame, text="Add PDF Files", command=self.add_files)
        self.btn_add.grid(row=0, column=0, padx=5, pady=5)

        self.btn_delete = ttk.Button(btn_frame, text="Delete Selected", command=self.delete_selected)
        self.btn_delete.grid(row=0, column=1, padx=5, pady=5)
        self.btn_delete.grid_remove()

        self.btn_clear = ttk.Button(btn_frame, text="Clear All", command=self.clear_all)
        self.btn_clear.grid(row=0, column=2, padx=5, pady=5)
        self.btn_clear.grid_remove()

        self.btn_up = ttk.Button(btn_frame, text="Move Up", command=self.move_up)
        self.btn_up.grid(row=0, column=3, padx=5, pady=5)
        self.btn_up.grid_remove()

        self.btn_down = ttk.Button(btn_frame, text="Move Down", command=self.move_down)
        self.btn_down.grid(row=0, column=4, padx=5, pady=5)
        self.btn_down.grid_remove()

        self.btn_merge = ttk.Button(root, text="Merge PDFs", command=self.merge_pdfs)
        self.btn_merge.pack(pady=(10, 15), ipadx=20, ipady=8)
        self.btn_merge.state(['disabled'])

        folder_frame = ttk.Frame(root)
        folder_frame.pack(padx=20, pady=5, fill="x")

        self.output_folder = os.path.expanduser("~/Desktop")
        self.output_label = ttk.Label(folder_frame, text=f"Save folder: {self.output_folder}", background="#f0f2f5", anchor="w")
        self.output_label.pack(side="left", fill="x", expand=True, padx=(5,0))

        self.change_folder_btn = ttk.Button(folder_frame, text="Change Folder", command=self.change_output_folder, style="Small.TButton")
        self.change_folder_btn.pack(side="left", padx=5)

        self.btn_open_folder = ttk.Button(root, text="Open Save Folder", command=self.open_output_folder, style="Small.TButton")
        self.btn_open_folder.pack(pady=(5, 10))

        self.created_by = ttk.Label(root, text="", font=("Segoe UI", 9, "italic"), foreground="gray")
        self.created_by.pack(pady=(0, 10))

        self.files = []
        self.last_saved_file = None

        self.drag_data = {"item": None, "index": None}
        self.listbox.bind("<ButtonPress-1>", self.start_drag)
        self.listbox.bind("<B1-Motion>", self.do_drag)
        self.listbox.bind("<ButtonRelease-1>", self.stop_drag)

    def show_about(self):
        about = Toplevel(self.root)
        about.title("About")
        about.configure(bg="#f0f2f5")
        about.geometry("460x300")
        about.resizable(False, False)
        ttk.Label(about, text="PDF Merger", font=("Segoe UI", 14, "bold"), background="#f0f2f5").pack(pady=5)
        ttk.Label(about, text="Version 1.00.4", font=("Segoe UI", 10), background="#f0f2f5").pack(pady=2)
        ttk.Label(about, text="Created by Ryan Sorrell", font=("Segoe UI", 10), background="#f0f2f5").pack(pady=2)
        ttk.Label(about, text="© 2025 - All Rights Reserved", font=("Segoe UI", 9, "italic"), background="#f0f2f5").pack(pady=2)
        ttk.Label(about, text="Unauthorized copying of this file, via any medium is strictly prohibited.",
                  font=("Segoe UI", 8), wraplength=440, justify="center", background="#f0f2f5").pack(pady=5)
        ttk.Label(about, text="Simply drag or click Add PDF Files and select your PDFs holding CTRL or Shift buttons.\n"
                             "Reorder them by dragging the files or clicking Move Up or Down buttons\n"
                             "Opening and changing the Save location are also supported.",
                  wraplength=440, justify="center", background="#f0f2f5").pack(pady=10)
        ttk.Button(about, text="Close", command=about.destroy).pack(pady=5)

    def add_files(self):
        new_files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if new_files:
            self.drag_hint.place_forget()
            self.process_files(new_files)
            self.update_buttons()

    def on_drop(self, event):
        dropped_files = self.root.tk.splitlist(event.data)
        if dropped_files:
            self.drag_hint.place_forget()
            self.process_files(dropped_files)
            self.update_buttons()

    def update_buttons(self):
        if self.files:
            self.btn_merge.state(['!disabled'])
            self.btn_clear.grid()
        else:
            self.btn_merge.state(['disabled'])
            self.btn_clear.grid_remove()
            self.hide_action_buttons()
            self.drag_hint.place(relx=0.5, rely=0.5, anchor="center")
        self.btn_delete.grid_remove()
        self.btn_up.grid_remove()
        self.btn_down.grid_remove()

    def process_files(self, files):
        invoice_files = []
        normal_files = []
        bol_files = []

        for file in files:
            if file.lower().endswith(".pdf") and all(file != f[0] for f in self.files):
                try:
                    reader = PdfReader(file)
                    pages = len(reader.pages)
                except Exception:
                    pages = 0
                name = os.path.basename(file).lower()
                if "invoice" in name:
                    invoice_files.append((file, pages))
                elif re.search(r'\bbol\b|\bbol-\d*\b', name):
                    bol_files.append((file, pages))
                else:
                    normal_files.append((file, pages))

        ordered_files = invoice_files + normal_files + bol_files
        for file, pages in ordered_files:
            self.files.append((file, pages))
            self.listbox.insert(tk.END, f"{os.path.basename(file)} ({pages} pages)")

        self.reorder_invoice_and_bol()
        self.update_total_pages()

    def reorder_invoice_and_bol(self):
        invoices = [f for f in self.files if "invoice" in os.path.basename(f[0]).lower()]
        bol = [f for f in self.files if re.search(r'\bbol\b|\bbol-\d*\b', os.path.basename(f[0]).lower())]
        others = [f for f in self.files if f not in invoices and f not in bol]

        self.files = invoices + others + bol

        self.listbox.delete(0, tk.END)
        for f in self.files:
            self.listbox.insert(tk.END, f"{os.path.basename(f[0])} ({f[1]} pages)")

    def update_total_pages(self):
        total = sum(pages for _, pages in self.files)
        if hasattr(self, 'total_pages_label'):
            self.total_pages_label.config(text=f"Total Pages: {total}")
        else:
            self.total_pages_label = ttk.Label(self.root, text=f"Total Pages: {total}", font=("Segoe UI", 10, "bold"), background="#f0f2f5")
            self.total_pages_label.pack(pady=(0, 10))

    def delete_selected(self):
        selected_indices = list(self.listbox.curselection())
        if not selected_indices:
            return
        for i in sorted(selected_indices, reverse=True):
            del self.files[i]
            self.listbox.delete(i)

        self.update_buttons()
        self.update_total_pages()

    def clear_all(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self.update_buttons()
        self.update_total_pages()

    def hide_action_buttons(self):
        self.btn_delete.grid_remove()
        self.btn_up.grid_remove()
        self.btn_down.grid_remove()

    def show_action_buttons(self):
        self.btn_delete.grid()
        self.btn_up.grid()
        self.btn_down.grid()

    def on_list_select(self, event):
        selected = self.listbox.curselection()
        if selected:
            self.show_action_buttons()
        else:
            self.hide_action_buttons()

    def move_up(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        indices = sorted(selected)
        for i in indices:
            if i == 0:
                continue
            self.files[i], self.files[i-1] = self.files[i-1], self.files[i]
        self.refresh_listbox()
        new_selection = [i-1 if i > 0 else i for i in indices]
        for i in new_selection:
            self.listbox.selection_set(i)
        self.update_total_pages()

    def move_down(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        indices = sorted(selected, reverse=True)
        max_idx = len(self.files) - 1
        for i in indices:
            if i == max_idx:
                continue
            self.files[i], self.files[i+1] = self.files[i+1], self.files[i]
        self.refresh_listbox()
        new_selection = [i+1 if i < max_idx else i for i in indices]
        for i in new_selection:
            self.listbox.selection_set(i)
        self.update_total_pages()

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for f in self.files:
            self.listbox.insert(tk.END, f"{os.path.basename(f[0])} ({f[1]} pages)")

    def change_output_folder(self):
        new_folder = filedialog.askdirectory()
        if new_folder:
            self.output_folder = new_folder
            self.output_label.config(text=f"Save folder: {self.output_folder}")

    def merge_pdfs(self):
        if not self.files or len(self.files) < 2:
            messagebox.showerror("Error", "Select at least two PDF files.")
            return

        merger = PdfMerger()
        for pdf, _ in self.files:
            merger.append(pdf)

        invoice_file = next((f for f in self.files if "invoice" in os.path.basename(f[0]).lower()), None)
        default_name = "merged.pdf"
        if invoice_file:
            digits = re.findall(r'\d+', os.path.basename(invoice_file[0]))
            if digits:
                default_name = digits[0] + ".pdf"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialdir=self.output_folder,
            initialfile=default_name,
            title="Save Merged PDF As"
        )

        if not save_path:
            merger.close()
            return

        try:
            with open(save_path, "wb") as f_out:
                merger.write(f_out)
            merger.close()
            self.last_saved_file = save_path
            self.output_folder = os.path.dirname(save_path)
            self.output_label.config(text=f"Save folder: {self.output_folder}")
            self.show_toast("Merged file saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

    def show_toast(self, message):
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.configure(bg="#444")
        toast.attributes("-topmost", True)

        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()

        width, height = 320, 80
        x = root_x + (root_w - width) // 2
        y = root_y + (root_h - height) // 2
        toast.geometry(f"{width}x{height}+{x}+{y}")

        label = tk.Label(toast, text=message, bg="#444", fg="white", font=("Segoe UI", 13, "bold"))
        label.pack(fill="both", expand=True, padx=15, pady=15)

        toast.after(3000, toast.destroy)

    def open_output_folder(self):
        if os.path.exists(self.output_folder):
            try:
                os.startfile(self.output_folder)  # Open the folder reliably on Windows
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open folder:\n{e}")
        else:
            messagebox.showerror("Error", f"Folder does not exist:\n{self.output_folder}")

    def start_drag(self, event):
        widget = event.widget
        self.drag_data["index"] = widget.nearest(event.y)
        self.drag_data["item"] = widget.get(self.drag_data["index"])
        self.listbox.itemconfig(self.drag_data["index"], bg="#cce5ff")

    def do_drag(self, event):
        widget = event.widget
        new_index = widget.nearest(event.y)
        if new_index != self.drag_data["index"]:
            self.listbox.itemconfig(self.drag_data["index"], bg="white")
            widget.delete(self.drag_data["index"])
            widget.insert(new_index, self.drag_data["item"])
            self.files.insert(new_index, self.files.pop(self.drag_data["index"]))
            self.drag_data["index"] = new_index
            self.listbox.itemconfig(new_index, bg="#cce5ff")

    def stop_drag(self, event):
        if self.drag_data["index"] is not None:
            self.listbox.itemconfig(self.drag_data["index"], bg="white")
        self.drag_data = {"item": None, "index": None}


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
