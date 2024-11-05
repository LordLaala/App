import customtkinter as ctk
from tkinter import filedialog, StringVar
from App import SP_Json_Converter as cs

# Initialize the main window
ctk.set_appearance_mode("Dark")  # Set to always use dark mode
ctk.set_default_color_theme("blue")
app = ctk.CTk()
app.geometry("400x300")
app.title("Data Export Tool")

# Set window icon (ensure the path is correct)
app.iconbitmap("assets\\WhatsApp-Image-2024-10-22-at-09.28.19.ico")  # Replace with your logo icon path

# Variables to store paths and format choice
input_folder_path = StringVar()
output_folder_path = StringVar()
output_format = StringVar(value="json")


# Function to select input folder
def select_input_folder():
    folder = filedialog.askdirectory()
    print(folder.replace("\\","\\\\"))
    if folder:
        input_folder_path.set(folder)


# Function to select output folder
def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_path.set(folder)



# Function to handle export
def export_data():
    input_path = input_folder_path.get()
    output_path = output_folder_path.get()
    format_choice = output_format.get()

    if not input_path or not output_path:
        ctk.messagebox.showerror("Error", "Please select both input and output folders.")
        return


    # Add your data processing and export code here
    match format_choice:
        case 'JSON':
            cs.main(input_path,output_path)
        case 'Excel':
            cs.main(input_path, output_path)
            excel = DataframeMaker(output_path)
            excel.make_json_to_df()


# Layout components
# Subtle Logo placeholder
try:
    logo_image = Image.open("path_to_logo.png")  # Replace with your logo path
    logo_image = logo_image.resize((50, 50), Image.ANTIALIAS)
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = ctk.CTkLabel(app, image=logo_photo, text="")
    logo_label.pack(pady=5)
except Exception as e:
    print("Logo not found, using text placeholder instead.")
    logo_label = ctk.CTkLabel(app, text="Logo", font=("Arial", 10), fg_color=None)
    logo_label.pack(pady=5)

# Input folder selection
input_frame = ctk.CTkFrame(app)
input_frame.pack(pady=5, padx=10, fill="x")

input_label = ctk.CTkLabel(input_frame, text="Input Folder:")
input_label.pack(side="left", padx=5)

input_entry = ctk.CTkEntry(input_frame, textvariable=input_folder_path, width=180)
input_entry.pack(side="left", padx=5)

input_button = ctk.CTkButton(input_frame, text="Browse", command=select_input_folder, width=80)
input_button.pack(side="right", padx=5)

# Output folder selection
output_frame = ctk.CTkFrame(app)
output_frame.pack(pady=5, padx=10, fill="x")

output_label = ctk.CTkLabel(output_frame, text="Output Folder:")
output_label.pack(side="left", padx=5)

output_entry = ctk.CTkEntry(output_frame, textvariable=output_folder_path, width=180)
output_entry.pack(side="left", padx=5)

output_button = ctk.CTkButton(output_frame, text="Browse", command=select_output_folder, width=80)
output_button.pack(side="right", padx=5)

# Output format selection
format_frame = ctk.CTkFrame(app)
format_frame.pack(pady=10)

json_radio = ctk.CTkRadioButton(format_frame, text="JSON", variable=output_format, value="JSON")
json_radio.pack(side="left", padx=10)

excel_radio = ctk.CTkRadioButton(format_frame, text="Excel", variable=output_format, value="Excel")
excel_radio.pack(side="right", padx=10)

# Export button
export_button = ctk.CTkButton(app, text="Export", command=export_data, width=100)
export_button.pack(pady=20)

# Run the main loop
app.mainloop()
