import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
import webbrowser

# Global variables
company_logo_path = ""
company_info = ""
settings_file = "DBI_Settings.txt"
loads_file = "Loads.csv"
load_number = 1
weight_entry = ""
length_entry = ""
width_entry = ""
height_entry = ""

# Global variables for GUI
canvas = None
canvas_width = 0
settings_window = None  # Declare settings_window globally
query_entry = None
query_text = ""


# Function to save settings
def save_settings():
    global company_logo_path, company_info
    settings = {
        "company_logo_path": company_logo_path,
        "company_info": company_info
    }
    with open(settings_file, "w") as file:
        for key, value in settings.items():
            file.write(f"{key}={value}\n")
    messagebox.showinfo("Settings", "Settings saved successfully!")


def load_settings():
    global company_logo_path, company_info, settings_window

    # Check if settings window is already open
    if settings_window is not None and settings_window.winfo_exists():
        settings_window.lift()
        return

    settings = {"company_logo_path": company_logo_path, "company_info": company_info}

    # Create settings window
    settings_window = tk.Toplevel()
    settings_window.title("Settings")
    settings_window.geometry("400x300")

    def select_logo(settings):
        global company_logo_path
        company_logo_path = filedialog.askopenfilename()
        settings["company_logo_path"] = company_logo_path
        messagebox.showinfo("Settings", "Company logo selected successfully!")
        display_logo_and_info()

    def input_company_info(settings):
        global company_info
        company_info = simpledialog.askstring("Settings", "Enter company information:")
        settings["company_info"] = company_info
        messagebox.showinfo("Settings", "Company information saved successfully!")
        display_logo_and_info()

    # Create select logo button
    select_logo_button = tk.Button(settings_window, text="Select Logo", command=lambda: select_logo(settings))
    select_logo_button.pack(pady=10)

    # Create input company info button
    input_info_button = tk.Button(settings_window, text="Input Company Info", command=lambda: input_company_info(settings))
    input_info_button.pack(pady=10)

    # Display logo and company info
    display_logo_and_info()


# Function to display company logo and information
def display_logo_and_info():
    global company_logo_path, company_info
    canvas.delete("all")
    if company_logo_path:
        logo_image = tk.PhotoImage(file=company_logo_path)
        canvas.create_image(canvas_width // 2, 10, image=logo_image)
        canvas.image = logo_image
    if company_info:
        canvas.create_text(canvas_width // 2, 150, text=company_info, font=("Arial", 12))


# Function to save load details
def save_load_details():
    global load_number, height_entry, width_entry, length_entry, weight_entry, product_description_entry, special_comments_entry, customer_info_entry, shipper_info_entry, driver_info_entry

    length = length_entry.get()
    width = width_entry.get()
    height = height_entry.get()
    weight = weight_entry.get()

    product_description = product_description_entry.get()
    special_comments = special_comments_entry.get("1.0", tk.END).strip()
    customer_info = customer_info_entry.get("1.0", tk.END).strip()
    shipper_info = shipper_info_entry.get("1.0", tk.END).strip()
    driver_info = driver_info_entry.get("1.0", tk.END).strip()

    fields = [
        "Load ID",
        "Length",
        "Width",
        "Height",
        "Weight",
        "Product Description",
        "Special Comments",
        "Customer Information",
        "Shipper Information",
        "Driver Information",
        "Date"
    ]

    data = {
        "Load ID": load_number,
        "Length": length,
        "Width": width,
        "Height": height,
        "Weight": weight,
        "Product Description": product_description,
        "Special Comments": special_comments,
        "Customer Information": customer_info,
        "Shipper Information": shipper_info,
        "Driver Information": driver_info,
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    try:
        with open(loads_file, "a", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=fields)
            if file.tell() == 0:
                writer.writeheader()
            writer.writerow(data)

        messagebox.showinfo("Load Details", "Load details saved successfully!")
        load_number += 1
        clear_fields()
    except Exception as e:
        messagebox.showerror("Error", f"Error saving load details: {str(e)}")


def generate_pdf():
    global loads_file

    # User input to select load ID
    load_id = simpledialog.askinteger("Generate PDF", "Enter Load ID:")
    if load_id is None:
        return

    # Generate PDF file name with load ID suffix
    pdf_file = f"DBI_LOAD_{load_id}.pdf"

    # Check if loads file exists
    if os.path.exists(loads_file):
        # Read loads data from CSV
        with open(loads_file, "r", newline="") as file:
            reader = csv.DictReader(file)
            loads_data = list(reader)

        # Find the load details for the specified load ID
        load_details = None
        for row in loads_data:
            if int(row.get("Load ID", "")) == load_id:
                load_details = row
                break

        if load_details is not None:
            # Generate PDF
            doc = SimpleDocTemplate(pdf_file, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []

            # Add logo image if available
            if company_logo_path:
                logo_image = Image(company_logo_path, width=100, height=100)
                elements.append(logo_image)

            # Add company information if available
            if company_info:
                company_info_paragraph = Paragraph(company_info, styles["Normal"])
                elements.append(company_info_paragraph)

            # Add load details from the selected load number
            load_paragraph = Paragraph(f"Load {load_id}:", styles["Heading2"])
            elements.append(load_paragraph)

            # Add product details
            elements.append(Spacer(1, 10))
            elements.append(Paragraph(f"Load ID: {load_details.get('Load ID', '')}", styles["Normal"]))
            elements.append(Paragraph(f"Length: {load_details.get('Length', '')}", styles["Normal"]))
            elements.append(Paragraph(f"Width: {load_details.get('Width', '')}", styles["Normal"]))
            elements.append(Paragraph(f"Height: {load_details.get('Height', '')}", styles["Normal"]))
            elements.append(Paragraph(f"Weight: {load_details.get('Weight', '')}", styles["Normal"]))
            elements.append(Spacer(1, 10))

            # Add special comments
            special_comments = load_details.get('Special Comments', '')
            if special_comments:
                elements.append(Paragraph("Special Comments:", styles["Heading2"]))
                elements.append(Paragraph(special_comments, styles["Normal"]))
                elements.append(Spacer(1, 10))

            # Add customer information
            customer_info = load_details.get('Customer Information', '')
            if customer_info:
                elements.append(Paragraph("Customer Information:", styles["Heading2"]))
                elements.append(Paragraph(customer_info, styles["Normal"]))
                elements.append(Spacer(1, 10))

            # Add shipping information
            shipper_info = load_details.get('Shipper Information', '')
            if shipper_info:
                elements.append(Paragraph("Shipping Information:", styles["Heading2"]))
                elements.append(Paragraph(shipper_info, styles["Normal"]))
                elements.append(Spacer(1, 10))

            # Add driver information
            driver_info = load_details.get('Driver Information', '')
            if driver_info:
                elements.append(Paragraph("Driver Information:", styles["Heading2"]))
                elements.append(Paragraph(driver_info, styles["Normal"]))
                elements.append(Spacer(1, 10))

            doc.build(elements)
            messagebox.showinfo("PDF Generation", f"PDF generated successfully: {pdf_file}")
        else:
            messagebox.showwarning("Warning", f"No load found with Load ID: {load_id}")
    else:
        messagebox.showerror("Error", "No load data found in the CSV file.")


def exit_program():
    if messagebox.askokcancel("Exit", "Do you want to exit the program?"):
        root.destroy()


# Function to clear all input fields
def clear_fields():
    length_entry.delete(0, tk.END)
    width_entry.delete(0, tk.END)
    height_entry.delete(0, tk.END)
    weight_entry.delete(0, tk.END)
    product_description_entry.delete(0, tk.END)
    special_comments_entry.delete("1.0", tk.END)
    customer_info_entry.delete("1.0", tk.END)
    shipper_info_entry.delete("1.0", tk.END)
    driver_info_entry.delete("1.0", tk.END)


# Function to query loads
def query_loads():
    global query_text
    query = query_entry.get()
    if query:
        loads_found = []
        with open(loads_file, "r", newline="") as file:
            reader = csv.DictReader(file)
            for row in reader:
                load_id = row.get("Load ID", "")
                if load_id and query.lower() in load_id.lower():
                    loads_found.append(row)
        
        if loads_found:
            messagebox.showinfo("Query Results", f"{len(loads_found)} load(s) found matching the query.")
            # Display the loads found
            for load in loads_found:
                load_id = load.get("Load ID", "")
                length = load.get("Length", "")
                width = load.get("Width", "")
                height = load.get("Height", "")
                weight = load.get("Weight", "")
                product_description = load.get("Product Description", "")
                special_comments = load.get("Special Comments", "")
                customer_info = load.get("Customer Information", "")
                shipper_info = load.get("Shipper Information", "")
                driver_info = load.get("Driver Information", "")
                date = load.get("Date", "")
                messagebox.showinfo(
                    f"Load ID: {load_id}",
                    f"Length: {length}\n"
                    f"Width: {width}\n"
                    f"Height: {height}\n"
                    f"Weight: {weight}\n"
                    f"Product Description: {product_description}\n"
                    f"Special Comments: {special_comments}\n"
                    f"Customer Information: {customer_info}\n"
                    f"Shipper Information: {shipper_info}\n"
                    f"Driver Information: {driver_info}\n"
                    f"Date: {date}"
                )
        else:
            messagebox.showinfo("Query Results", "No loads found matching the query.")
    else:
        messagebox.showwarning("Query", "Please enter a query.")


# Function to send email with the generated PDF as an attachment
def send_email():
    global query_text
    load_id = query_text.get()

    if load_id:
        pdf_file = f"DBI_LOAD_{load_id}.pdf"
        if os.path.exists(pdf_file):
            # Email functionality using webbrowser
            subject = f"Load ID: {load_id} - PDF Report"
            body = "Please find the attached PDF report for the load details."
            mailto = f"mailto:?subject={subject}&body={body}&attach={os.path.abspath(pdf_file)}"
            webbrowser.open(mailto)
       
        else:
            messagebox.showerror("Error", f"PDF file not found for Load ID: {load_id}")
    else:
        messagebox.showwarning("Email", "Please select a load ID to send as an email attachment.")


# Create the main window
root = tk.Tk()
root.title("BuilderBear Load Manager")
root.geometry("800x720")

# Create the canvas for displaying logo and company info
canvas = tk.Canvas(root, width=800, height=200)
canvas.pack()

# Create the left frame
left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=10)

# Create the right frame
right_frame = tk.Frame(root)
right_frame.pack(side=tk.RIGHT, padx=10)

# Create the load details form
load_details_label = tk.Label(left_frame, text="Load Details", font=("Arial", 14, "bold"))
load_details_label.grid(row=0, column=0, columnspan=2, pady=10)

length_label = tk.Label(left_frame, text="Length:")
length_label.grid(row=1, column=0, padx=5, pady=5)
length_entry = tk.Entry(left_frame)
length_entry.grid(row=1, column=1, padx=5, pady=5)

width_label = tk.Label(left_frame, text="Width:")
width_label.grid(row=2, column=0, padx=5, pady=5)
width_entry = tk.Entry(left_frame)
width_entry.grid(row=2, column=1, padx=5, pady=5)

height_label = tk.Label(left_frame, text="Height:")
height_label.grid(row=3, column=0, padx=5, pady=5)
height_entry = tk.Entry(left_frame)
height_entry.grid(row=3, column=1, padx=5, pady=5)

weight_label = tk.Label(left_frame, text="Weight:")
weight_label.grid(row=4, column=0, padx=5, pady=5)
weight_entry = tk.Entry(left_frame)
weight_entry.grid(row=4, column=1, padx=5, pady=5)

product_description_label = tk.Label(left_frame, text="Product Description:")
product_description_label.grid(row=5, column=0, padx=5, pady=5)
product_description_entry = tk.Entry(left_frame)
product_description_entry.grid(row=5, column=1, padx=5, pady=5)

special_comments_label = tk.Label(left_frame, text="Special Comments:")
special_comments_label.grid(row=6, column=0, padx=5, pady=5)
special_comments_entry = tk.Text(left_frame, height=3, width=20)
special_comments_entry.grid(row=6, column=1, padx=5, pady=5)

customer_info_label = tk.Label(left_frame, text="Customer Information:")
customer_info_label.grid(row=7, column=0, padx=5, pady=5)
customer_info_entry = tk.Text(left_frame, height=3, width=20)
customer_info_entry.grid(row=7, column=1, padx=5, pady=5)

shipper_info_label = tk.Label(left_frame, text="Shipper Information:")
shipper_info_label.grid(row=8, column=0, padx=5, pady=5)
shipper_info_entry = tk.Text(left_frame, height=3, width=20)
shipper_info_entry.grid(row=8, column=1, padx=5, pady=5)

driver_info_label = tk.Label(left_frame, text="Driver Information:")
driver_info_label.grid(row=9, column=0, padx=5, pady=5)
driver_info_entry = tk.Text(left_frame, height=3, width=20)
driver_info_entry.grid(row=9, column=1, padx=5, pady=5)

# Create the buttons
save_button = tk.Button(left_frame, text="Save", command=save_load_details)
save_button.grid(row=10, column=0, padx=5, pady=10)

pdf_button = tk.Button(left_frame, text="Generate PDF", command=generate_pdf)
pdf_button.grid(row=10, column=1, padx=5, pady=10)

# Create the right frame elements
query_label = tk.Label(right_frame, text="Query Loads", font=("Arial", 14, "bold"))
query_label.pack(pady=10)

query_text = tk.StringVar()
query_entry = tk.Entry(right_frame, textvariable=query_text)
query_entry.pack(pady=5)

query_button = tk.Button(right_frame, text="Query", command=query_loads)
query_button.pack(pady=5)

clear_button = tk.Button(right_frame, text="Clear", command=clear_fields)
clear_button.pack(pady=5)

email_button = tk.Button(right_frame, text="Email PDF", command=send_email)
email_button.pack(pady=5)

settings_button = tk.Button(right_frame, text="Settings", command=load_settings)
settings_button.pack(pady=5)

exit_button = tk.Button(right_frame, text="Exit", command=exit_program)
exit_button.pack(pady=5)

# Load settings if available
if os.path.exists(settings_file):
    with open(settings_file, "r") as file:
        for line in file:
            line = line.strip()
            if "=" in line:
                key, value = line.split("=", 1)
                if key == "company_logo_path":
                    company_logo_path = value
                elif key == "company_info":
                    company_info = value

# Display logo and company info
display_logo_and_info()

# Start the GUI main loop
root.mainloop()

