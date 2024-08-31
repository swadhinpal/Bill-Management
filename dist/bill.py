import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics  # Add this import statement
import pdfkit
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        process_excel(file_path)

def process_excel(file_path):
    df = pd.read_excel(file_path)
    pdfmetrics.registerFont(TTFont('abc', resource_path('bn2.ttf')))

    try:
        # Specify the sheet name or index (0-indexed) of the second sheet
        df = pd.read_excel(file_path, sheet_name=1)
        df1 = pd.read_excel(file_path, sheet_name=0)
    except FileNotFoundError:
        print("File not found. Please provide the correct path to the Excel file.")
        return

    # Process Excel data and generate PDF content
    generate_pdf(df, df1)

def generate_pdf(df, df1):
    def save_pdf():
        file_name = entry_file_name.get()
        if not file_name:
            return

        file_path = filedialog.askdirectory()
        if not file_path:
            return
        
        if df1.empty:
           print("DataFrame df1 is empty. Cannot generate PDF.")
           return
        
        for index, row in df1.iterrows():
            sector = row['Sector']
            amount = row['Amount']
            print(f"Sector: {sector}, Amount: {amount}")

        q_tk = df1.loc[df1['Sector'].str.strip() == 'Question-MSC  (Full)', 'Amount']

        if not q_tk.empty:
               q_tk = q_tk.values[0]
               print("Amount for 'Question-MSC(Full)':", q_tk)
        else:
               print("No amount found for 'Question-MSC(Full)'")

    
        # Generate PDF content
        example_text = ""
        for index, row in df.iterrows():
            if 'Question' in df.columns:
                t_name = row['Teacher Name']
                sem = row['Semester']
                if sem.replace(' ', '') == '1styear1stsemester':
                    sem = '১ম বর্ষ ১ম সেমিস্টার'
                elif sem.replace(' ', '') == '1styear2ndsemester' :
                    sem = '১ম বর্ষ ২য় সেমিস্টার'
                elif sem.replace(' ', '') == '2ndyear1stsemester' :
                    sem = '২য় বর্ষ ১ম সেমিস্টার'
                elif sem.replace(' ', '') == '2ndyear2ndsemester' :
                    sem = '২য় বর্ষ ২য় সেমিস্টার'
                elif sem.replace(' ', '') == '3rdyear1stsemester' :
                    sem = '৩য় বর্ষ ১ম সেমিস্টার'
                elif sem.replace(' ', '') == '3rdyear2ndsemester' :
                    sem = '৩য় বর্ষ ২য় সেমিস্টার'
                elif sem.replace(' ', '') == '4thyear1stsemester' :
                    sem = '৪র্থ বর্ষ ১ম সেমিস্টার'
                elif sem.replace(' ', '') == '4thyear2ndsemester' :
                    sem = '৪র্থ বর্ষ ২য় সেমিস্টার'     
                course = row['Course Code']
                q_tk=2250
                if 'MSC(Full)' in row['Question'].replace(' ', ''):
                   
                    q_tk = df1.loc[df1['Sector'].str.strip()== 'Question-MSC  (Full)', 'Amount'].values[0]
                    example_text= f"জনাব {t_name}<br>(ক) প্রশ্নপত্র প্রণয়ন: একটি স্নাতকোত্তর পূর্ণপত্র SE পরীক্ষার {sem} বিষয় {course} কোর্স : {q_tk} টাকা"
                    print(f"জনাব {t_name}\n(ক) প্রশ্নপত্র প্রণয়ন: একটি স্নাতকোত্তর পূর্ণপত্র SE পরীক্ষার {sem} বিষয় {course} কোর্স : {q_tk} টাকা")
                elif 'MSC(Half)' in row['Question']:
                    #q_tk =  df1.loc[df1['Sector'] == 'Question-MSC(Half)', 'Amount'].values[0]
                    print(f"জনাব {t_name}\n(ক) প্রশ্নপত্র প্রণয়ন: একটি স্নাতকোত্তর অর্ধপত্র SE পরীক্ষার {sem} বিষয় {course} কোর্স : {q_tk} টাকা")
                elif 'BSC(Half)' in row['Question']:
                    q_tk =  df1.loc[df1['Sector'] == 'Question-BSC  (Half)', 'Amount'].values[0]
                    print(f"জনাব {t_name}\n(ক) প্রশ্নপত্র প্রণয়ন: একটি স্নাতক অর্ধপত্র SE পরীক্ষার {sem} বিষয় {course} কোর্স : {q_tk} টাকা")
                elif 'BSC(Full)' in row['Question']:
                    q_tk =  df1.loc[df1['Sector'] == 'Question-BSC (Full)', 'Amount'].values[0]
                    print(f"জনাব {t_name}\n(ক) প্রশ্নপত্র প্রণয়ন: একটি স্নাতক পূর্ণপত্র SE পরীক্ষার {sem} বিষয় {course} কোর্স : {q_tk} টাকা")
            else:
                print("Column 'Question' does not exist in the DataFrame.")

            print(', '.join(df.columns))

            if 'Paper Evaluation' in df.columns:
                t_name = row['Teacher Name']
    
                
                second_semester_value = row['Semester.1']
                
                second_course_value = row['Course Code.1']
                
                final_khata = row['Number of khata(Final)']
                mid_khata = row['Number of khata (Mid)']

                if 'MSC(Full)' in row['Paper Evaluation'].replace(' ', ''): 

                    # Get the amount from df1
                    q_tk = df1.loc[df1['Sector'].str.strip() == 'Paper Evaluation-MSC  (Full)', 'Amount'].values[0] * final_khata
                    q_tk1 = df1.loc[df1['Sector'].str.strip() == 'Mid-1', 'Amount'].values[0] * mid_khata

                    example_text1 = f"(খ) উত্তরপত্র মূল্যায়ন: একটি স্নাতকোত্তর পূর্ণপত্র বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk} টাকা"
                    example_text2 = f"(গ) উত্তরপত্র মূল্যায়ন: একটি স্নাতকোত্তর পূর্ণপত্র অর্ধ-বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk1} টাকা"
                elif 'MSC(Half)' in row['Paper Evaluation'].replace(' ', ''): 

                    # Get the amount from df1
                    q_tk = df1.loc[df1['Sector'].str.strip() == 'Paper Evaluation-MSC (Half)', 'Amount'].values[0] * final_khata
                    q_tk1 = df1.loc[df1['Sector'].str.strip() == 'Mid-1', 'Amount'].values[0] * mid_khata

                    example_text1 = f"(খ) উত্তরপত্র মূল্যায়ন: একটি স্নাতকোত্তর অর্ধপত্র বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk} টাকা"
                    example_text2 = f"(গ) উত্তরপত্র মূল্যায়ন: একটি স্নাতকোত্তর অর্ধপত্র অর্ধ-বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk1} টাকা"
                elif 'BSC(Full)' in row['Paper Evaluation'].replace(' ', ''): 

                    # Get the amount from df1
                    q_tk = df1.loc[df1['Sector'].str.strip() == 'Paper Evaluation-BSC (Full)', 'Amount'].values[0] * final_khata
                    q_tk1 = df1.loc[df1['Sector'].str.strip() == 'Mid-1', 'Amount'].values[0] * mid_khata

                    example_text1 = f"(খ) উত্তরপত্র মূল্যায়ন: একটি স্নাতক পূর্ণপত্র বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk} টাকা"
                    example_text2 = f"(গ) উত্তরপত্র মূল্যায়ন: একটি স্নাতক পূর্ণপত্র অর্ধ-বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk1} টাকা"
                elif 'BSC(Half)' in row['Paper Evaluation'].replace(' ', ''): 

                    # Get the amount from df1
                    q_tk = df1.loc[df1['Sector'].str.strip() == 'Paper Evaluation-BSC (Half)', 'Amount'].values[0] * final_khata
                    q_tk1 = df1.loc[df1['Sector'].str.strip() == 'Mid-1', 'Amount'].values[0] * mid_khata

                    example_text1 = f"(খ) উত্তরপত্র মূল্যায়ন: একটি স্নাতক অর্ধপত্র বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk} টাকা"
                    example_text2 = f"(গ) উত্তরপত্র মূল্যায়ন: একটি স্নাতক অর্ধপত্র অর্ধ-বার্ষিক SE পরীক্ষার {second_semester_value} বিষয় {second_course_value} কোর্স : {q_tk1} টাকা"
                break;


        # HTML content for PDF
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <style>  
            body {{
              font-family: 'BengaliFont', sans-serif;
            }}
          </style>
        </head>
        <body>
          <p>{example_text}<br> {example_text1} <br> {example_text2}</p>
        </body>
        </html>
        """

        # Configure options for wkhtmltopdf
        options = {
            'page-size': 'Letter',
            'margin-top': '0.75in',
            'margin-right': '0.75in',
            'margin-bottom': '0.75in',
            'margin-left': '0.75in',
            'encoding': 'UTF-8'
        }

        # Path to wkhtmltopdf executable
        path_to_wkhtmltopdf=resource_path('wkhtmltopdf.exe')

        # Path to the output PDF file
        output_pdf_path = f"{file_path}\\{file_name}.pdf"

        # Generate PDF from HTML content using pdfkit
        pdfkit.from_string(html_content, output_pdf_path, options=options, configuration=pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf))
        print(f"PDF generated and saved successfully at {output_pdf_path}")
        top.destroy()

    top = tk.Toplevel()
    top.title("PDF Generation")
    
    label_file_name = tk.Label(top, text="Enter PDF File Name:")
    label_file_name.pack(pady=10)

    entry_file_name = tk.Entry(top)
    entry_file_name.pack(pady=5)

    button_save = tk.Button(top, text="Save PDF", command=save_pdf)
    button_save.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel to PDF Converter")

    browse_button = tk.Button(root, text="Select Excel File", command=browse_file)
    browse_button.pack(pady=20)

    root.mainloop()
