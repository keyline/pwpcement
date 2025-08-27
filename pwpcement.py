import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
import os
from pwpautomation import PWPInvoiceAutomation
from pwpupload import PWPUploadAutomation


ECO_GREEN = "#4CAF50"

def get_file(prompt, filetypes):
    return filedialog.askopenfilename(title=prompt, filetypes=filetypes)

def get_folder(prompt):
    return filedialog.askdirectory(title=prompt)

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')



def download_template(parent):
    TEMPLATE_PATH = "Template.xlsx"
    if not os.path.exists(TEMPLATE_PATH):
        messagebox.showerror("Error", f"Template file not found:\n{TEMPLATE_PATH}")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             title="Save Excel Template As")
    if file_path:
        try:
            shutil.copy(TEMPLATE_PATH, file_path)
            messagebox.showinfo("Success", f"Template saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")


def open_generate_form(parent):
    parent.withdraw()
    form = tk.Toplevel()
    form.title("Generate EPR Invoice")
    form.configure(bg=ECO_GREEN)
    center_window(form, 550, 450)

    def go_back():
        form.destroy()
        parent.deiconify()

    form.protocol("WM_DELETE_WINDOW", go_back)  # Handle window close

    tk.Button(form, text="← Back", command=go_back, bg="white", fg="black").pack(anchor="w", padx=10, pady=5)

    excel_var = tk.StringVar()
    orig_pdf_var = tk.StringVar()
    epr_pdf_var = tk.StringVar()
    output_excel_var = tk.StringVar()
    

    tk.Label(form, text="Source Excel File:", bg=ECO_GREEN, fg="white").pack(pady=5)
    tk.Entry(form, textvariable=excel_var, width=60).pack()
    tk.Button(form, text="Browse", command=lambda: excel_var.set(get_file("Source Excel File", [("Excel files", "*.xlsx *.xls")]))).pack(pady=2)

    tk.Label(form, text="Original Invoice Folder:", bg=ECO_GREEN, fg="white").pack(pady=5)
    tk.Entry(form, textvariable=orig_pdf_var, width=60).pack()
    tk.Button(form, text="Browse", command=lambda: orig_pdf_var.set(get_folder("Original Invoice Folder"))).pack(pady=2)

    tk.Label(form, text="Folder you want to generte EPR PDF's:", bg=ECO_GREEN, fg="white").pack(pady=5)
    tk.Entry(form, textvariable=epr_pdf_var, width=60).pack()
    tk.Button(form, text="Browse", command=lambda: epr_pdf_var.set(get_folder("Folder you want to generte EPR PDF's"))).pack(pady=2)
    
    tk.Label(form, text="File you want to generate execution result:", bg=ECO_GREEN, fg="white").pack(pady=5)
    tk.Entry(form, textvariable=output_excel_var, width=60).pack()
    tk.Button(form, text="Browse", command=lambda: output_excel_var.set(get_file("File you want to generate execution result", [("Excel files", "*.xlsx *.xls")]))).pack(pady=2)
    
    

    def submit():
        if not excel_var.get() or not orig_pdf_var.get() or not epr_pdf_var.get() or not output_excel_var.get():
            messagebox.showerror("Missing Input", "All fields are required.")
            return
        try:
            login_url = "https://eprplastic.cpcb.gov.in/#/plastic/home"      # Manual login page
            form_url = "https://eprplastic.cpcb.gov.in/#/epr/pwp-sales"        # Form page
            
            bot = PWPInvoiceAutomation(excel_var.get(), login_url, form_url, orig_pdf_var.get(), epr_pdf_var.get(), output_excel_var.get())
            bot.run()
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tk.Button(form, text="Initiate Automation", command=submit, bg="white", fg="black", font=("Arial", 11)).pack(pady=20)

def open_upload_form(parent):
    parent.withdraw()
    form = tk.Toplevel()
    form.title("Upload EPR Invoice")
    form.configure(bg=ECO_GREEN)
    center_window(form, 550, 350)

    def go_back():
        form.destroy()
        parent.deiconify()

    form.protocol("WM_DELETE_WINDOW", go_back)  # Handle window close

    tk.Button(form, text="← Back", command=go_back, bg="white", fg="black").pack(anchor="w", padx=10, pady=5)

    excel_var = tk.StringVar()
    epr_pdf_var = tk.StringVar() 
    result_excel_var = tk.StringVar()

    tk.Label(form, text="Select Excel File:", bg=ECO_GREEN, fg="white").pack(pady=5)
    tk.Entry(form, textvariable=excel_var, width=60).pack()
    tk.Button(form, text="Browse", command=lambda: excel_var.set(get_file("Select Excel File", [("Excel files", "*.xlsx")]))).pack(pady=2)

    tk.Label(form, text="EPR Generated PDF Folder:", bg=ECO_GREEN, fg="white").pack(pady=5)
    tk.Entry(form, textvariable=epr_pdf_var, width=60).pack()
    tk.Button(form, text="Browse", command=lambda: epr_pdf_var.set(get_folder("Select EPR Generated PDF Folder"))).pack(pady=2)
    
    tk.Label(form, text="Select Excel File where you want to see the final result:", bg=ECO_GREEN, fg="white").pack(pady=5)
    tk.Entry(form, textvariable=result_excel_var, width=60).pack()
    tk.Button(form, text="Browse", command=lambda: result_excel_var.set(get_file("Select Excel File where you want to see the final result", [("Excel files", "*.xlsx")]))).pack(pady=2)

    def submit():
        if not excel_var.get() or not epr_pdf_var.get() or not result_excel_var.get():
            messagebox.showerror("Missing Input", "All fields are required.")
            return
        try:
            login_url = "https://eprplastic.cpcb.gov.in/#/plastic/home"      
            form_url = "https://eprplastic.cpcb.gov.in/#/epr/details/sales" 
            
            uploadbot = PWPUploadAutomation(excel_var.get(), login_url, form_url, epr_pdf_var.get(), result_excel_var.get())
            uploadbot.run()            
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tk.Button(form, text="Initiate Automation", command=submit, bg="white", fg="black", font=("Arial", 11)).pack(pady=20)

def create_gui():
    root = tk.Tk()
    root.title("EcoEx :: PWP Cement [R-27082025]")
    window_width = 500
    window_height = 350
    center_window(root, window_width, window_height)
    root.configure(bg=ECO_GREEN)
    root.resizable(False, False)

    tk.Label(root, text="Choose an action", font=("Arial", 14, "bold"), bg=ECO_GREEN, fg="white").pack(pady=25)

    tk.Button(
        root, text="Download Entry Template", command=lambda: download_template(root),
        width=30, pady=8, bg="white", fg="black", font=("Arial", 11)
    ).pack(pady=15)
    
    tk.Button(
        root, text="Generate EPR Invoice", command=lambda: open_generate_form(root),
        width=30, pady=8, bg="white", fg="black", font=("Arial", 11)
    ).pack(pady=15)

    tk.Button(
        root, text="Upload EPR Invoice", command=lambda: open_upload_form(root),
        width=30, pady=8, bg="white", fg="black", font=("Arial", 11)
    ).pack(pady=15)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
