import tkinter as tk
from tkinter import PhotoImage, filedialog, messagebox
import pandas as pd
import json
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import os
import sys

def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def change_file_extension(filename, new_ext):
    return f"{'.'.join(filename.split('.')[:-1])}{new_ext}"

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if file_path:
        input_file.set(file_path)
        if not output_file.get():
            new_ext = '.xlsx' if file_type.get() == 'Excel' else '.csv'
            output_file.set(change_file_extension(file_path, new_ext))

def save_file_as():
    ext = '.xlsx' if file_type.get() == 'Excel' else '.csv'
    file_path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if file_path:
        output_file.set(file_path)
        file_type.set('Excel' if file_path.endswith('.xlsx') else 'CSV')

def process_survey_json_to_excel(json_data, output_path, bold_header=True, auto_fit=True):
    all_rows = []
    
    # Iterate through each survey in the JSON data
    for survey in json_data:
        survey_metadata = {
            "IdPublishSurvey": survey.get("IdPublishSurvey"),
            "IdSurveyTemplateChild": survey.get("IdSurveyTemplateChild"),
            "IdSurveyTemplateParent": survey.get("IdSurveyTemplateParent"),
            "IdSurvey": survey.get("IdSurvey"),
            "IdSurveyParent": survey.get("IdSurveyParent"),
            "IdUser": survey.get("IdUser"),
            "AcademicYear": survey.get("AcademicYear"),
            "Semester": survey.get("Semester"),
            "Role": survey.get("Role"),
            "TemplateTitle": survey.get("TemplateTitle"),
            "Status": survey.get("Status"),
        }
        
        # Process Sections
        for section in survey.get("Sections", []):
            section_metadata = {
                "SectionNumber": section.get("Number"),
                "SectionName": section.get("NameSection"),
                "SectionDescription": section.get("Description"),
            }
            
            # Process Questions
            for question in section.get("Questions", []):
                question_metadata = {
                    "QuestionNumber": question.get("Number"),
                    "QuestionText": question.get("QuestionText"),
                    "QuestionType": question.get("Type"),
                }
                
                # Process RespondentAnswers
                for answer in question.get("RespondentAnswers", []):
                    answer_data = {
                        "AnswerNumber": answer.get("Number"),
                        "AnswerValue": answer.get("Value"),
                    }
                    
                    # Combine all metadata into a single row
                    combined_row = {
                        **survey_metadata,
                        **section_metadata,
                        **question_metadata,
                        **answer_data,
                    }
                    all_rows.append(combined_row)
    
    # Convert to DataFrame
    df = pd.DataFrame(all_rows)
    
    # Write to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="SurveyData")
        worksheet = writer.sheets["SurveyData"]
        
        # Apply formatting if required
        if bold_header:
            bold_font = Font(bold=True)
            for col_num, value in enumerate(df.columns.values):
                col_letter = get_column_letter(col_num + 1)
                worksheet[f"{col_letter}1"].font = bold_font
        
        if auto_fit:
            for col_num, column in enumerate(df.columns.values):
                max_length = max(df[column].astype(str).map(len).max(), len(column))
                worksheet.column_dimensions[get_column_letter(col_num + 1)].width = max_length + 2

    print(f"Excel file saved successfully at: {output_path}")
def convert_json():
    try:
        data_list = []
        with open(input_file.get(), 'r') as file:
            data_list = json.load(file)

        output_path = output_file.get()
        all_data_rows = []
        if not is_questionnaire.get():
            df = pd.DataFrame(data_list)
            if file_type.get() == 'Excel':
                df.to_excel(output_path, index=False)
            else:
                # Save as CSV
                df.to_csv(output_path, index=False)
        else:
            process_survey_json_to_excel(data_list, output_path)

        # Show success message
        messagebox.showinfo("Success", f"File saved successfully at {output_path}")
    except Exception as e:
        # Show error message
        messagebox.showerror("Error", str(e))


def update_output_file(*args):
    if input_file.get() and output_file.get():
        new_ext = '.xlsx' if file_type.get() == 'Excel' else '.csv'
        output_file.set(change_file_extension(output_file.get(), new_ext))

def move_duck(x_pos):
    if x_pos < -logo_image.width():
        x_pos = splash_width
    logo_label.place(x=x_pos, y=(splash_height - logo_image.height()) // 2)
    splash_root.after(25, move_duck, x_pos - 5)

def show_main_window():
    splash_root.destroy()
    root = tk.Tk()
    root.title("BEBEK CONVERTER")
    iconPath = resource_path('duck.ico')
    root.iconbitmap(iconPath)
    root.geometry("400x260")
    root.resizable(False, False)

    global input_file, output_file, bold_header, auto_fit, is_questionnaire, file_type

    input_file = tk.StringVar()
    output_file = tk.StringVar()
    bold_header = tk.BooleanVar()
    auto_fit = tk.BooleanVar()
    is_questionnaire = tk.BooleanVar()
    file_type = tk.StringVar(value="Excel")
    file_type.trace_add("write", update_output_file)

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="JSON File:").grid(row=0, column=0, sticky=tk.W, pady=5)
    tk.Entry(frame, textvariable=input_file, width=40).grid(row=0, column=1, pady=5)
    tk.Button(frame, text="Browse", command=upload_file, width=10).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(frame, text="Save As:").grid(row=1, column=0, sticky=tk.W, pady=5)
    tk.Entry(frame, textvariable=output_file, width=40).grid(row=1, column=1, pady=5)
    tk.Button(frame, text="Save As", command=save_file_as, width=10).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(frame, text="File Type:").grid(row=2, column=0, sticky=tk.W, pady=5)
    tk.Radiobutton(frame, text="Excel", variable=file_type, value="Excel").grid(row=2, column=1, sticky=tk.W)
    tk.Radiobutton(frame, text="CSV", variable=file_type, value="CSV").grid(row=2, column=1)

    tk.Checkbutton(frame, text="Bold Headers", variable=bold_header).grid(row=3, column=1, sticky=tk.W, pady=5)
    tk.Checkbutton(frame, text="Auto Fit Columns", variable=auto_fit).grid(row=4, column=1, sticky=tk.W, pady=5)
    tk.Checkbutton(frame, text="Is Questionnaire", variable=is_questionnaire).grid(row=5, column=1, sticky=tk.W, pady=5)

    tk.Button(frame, text="Convert", command=convert_json, width=20).grid(row=6, column=1, pady=10)

    root.mainloop()


splash_root = tk.Tk()
splash_root.overrideredirect(True)
splash_width = 600  # Increased width to give more space for animation
splash_height = 350
screen_width = splash_root.winfo_screenwidth()
screen_height = splash_root.winfo_screenheight()
x = (screen_width // 2) - (splash_width // 2)
y = (screen_height // 2) - (splash_height // 2)
splash_root.geometry(f"{splash_width}x{splash_height}+{x}+{y}")

# Background
splash_frame = tk.Frame(splash_root, bg="#F0F0F0")
splash_frame.pack(fill="both", expand=True)

# Logo
logo_path = resource_path('duck.png')
logo_image = PhotoImage(file=logo_path)
logo_label = tk.Label(splash_frame, image=logo_image, bg="#F0F0F0")

# App name
app_name = tk.Label(splash_frame, text="BEBEK CONVERTER", font=("Arial", 24, "bold"), bg="#F0F0F0", fg="#333333")
app_name.pack(pady=(20, 0))

# Loading animation
loading_frame = tk.Frame(splash_frame, bg="#F0F0F0")
loading_frame.pack(side="bottom", pady=20)

def animate_loading(index=0):
    dots = ["   ", ".  ", ".. ", "..."]
    loading_label.config(text=f"Loading{dots[index]}")
    splash_root.after(500, animate_loading, (index + 1) % 4)

loading_label = tk.Label(loading_frame, text="Loading...", font=("Arial", 12), bg="#F0F0F0", fg="#666666")
loading_label.pack()

# Start animations
animate_loading()
move_duck(splash_width)  # Start the duck from the right edge

# Start the splash screen
splash_root.after(5000, show_main_window)  # Increased delay to 5 seconds to show more of the animation
splash_root.mainloop()