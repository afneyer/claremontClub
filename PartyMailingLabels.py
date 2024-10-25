import argparse
import os
from pathlib import Path
from tkinter import filedialog
import tkinter as tk
from typing import Any, IO

import pandas as pd
from fpdf import FPDF

global global_out_file
global global_download_dir

mail_name_col = 'mailing_name'
last_name_col = 'last_name'
addr1_col = 'address1'
addr2_col = 'address2'
city_col = 'city'
state_col = 'state'
zip_col = 'zip'


def get_file_from_user():
    window = tk.Tk()
    window.title("Open Club Express Export File")
    window.geometry("600x200")

    w = 600  # width for the Tk root
    h = 300  # height for the Tk root

    # get screen width and height
    ws = window.winfo_screenwidth()  # width of the screen
    hs = window.winfo_screenheight()  # height of the screen

    # calculate x and y coordinates for the Tk root window
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)

    # set the dimensions of the screen
    # and where it is placed
    window.geometry('%dx%d+%d+%d' % (w, h, x, y))

    # Create text widget and specify size.
    text = tk.Text(window, height=6, width=80, )
    text.config(font=("Arial", 14), spacing1=3, spacing2=3, spacing3=3)
    hint = ("Enter the name of the file exported from Club Express\n" +
            "- For labels of all active members go to Club Express\n" +
            "- Select Control Panel -> Club -> Admin Functions -> Data Export\n" +
            "- Select People -> Active Members (as above ... including other fields)\n" +
            "  -> Typically this will create a file called \"ActiveMembers.csv in the downloads folder\n" +
            "- Select the generated file\n")
    text.pack()
    text.insert(tk.END, hint)

    def open_file():
        global global_out_file
        global download_dir
        global_out_file = filedialog.askopenfile(initialfile="ActiveMembers.csv")
        window.destroy()

    # Create a button
    button = tk.Button(window, text="Open File", command=open_file)
    button.pack()

    # Run the application
    window.mainloop()

    return global_out_file


def create_mailing_labels():
    global global_download_dir
    input_file = get_file_from_user()
    df = pd.read_csv(input_file, sep=',')  # Use ',' as the separator
    df.columns = df.columns.str.strip()  # Strip leading and trailing spaces

    # Convert all the fields to strings
    for col in [mail_name_col, addr1_col, addr2_col, city_col, state_col, zip_col]:
        df[col] = df[col].astype(str)

    # Replace 'nan' with an empty string
    df.replace('nan', '', inplace=True)

    # Sort the DataFrame by the OWNER1 column
    df.sort_values(by=[last_name_col], inplace=True)

    # Create the mailing address
    df['mailing_address'] = df.apply(lambda arow: '\n'.join(filter(None, [arow[mail_name_col], ', '.join(
        filter(None, [arow[addr1_col], arow[addr2_col]])), ', '.join(
        [arow[city_col], arow[state_col], arow[zip_col]])])),
                                     axis=1)

    df['normalized_mailing_address'] = df.apply(lambda arow: normalized_mailing_address(arow), axis=1)

    df_unique_address = df.drop_duplicates(subset=['normalized_mailing_address'], keep='first')
    df_unique_address.reset_index(drop=True, inplace=True)

    df_dropped = df.duplicated(subset=[addr1_col, addr2_col, city_col, state_col, zip_col], keep='last')

    df = df_unique_address
    # Formatting for Office Depot 1" x 2-5/8" inch labels, 612-221, 3 rows and 10 colums per sheet

    pdf = FPDF('L', 'mm', (279.4, 215.9))  # Create a PDF in landscape orientation on letter size paper
    pdf.set_auto_page_break(False)
    # Dimensions of the text box in mm
    # box_height = 28.575
    # box_width = 114.3

    # Set the margins to 1 inch (approximately 25.4 mm) at the top and bottom and
    # 0.5 inch (approximately 12.7 mm) at the left
    pdf.set_top_margin(12.7)
    pdf.set_left_margin(5)
    pdf.set_right_margin(0.0)

    # Calculate the number of text boxes that can fit on a page
    # boxes_per_page_horizontally = int((pdf.w - 0.0 - 25.4) / box_width)
    boxes_per_page_horizontally = 3
    left_margin = 5.0
    horizontal_label_space = 3
    label_width = 66.7
    # boxes_per_page_vertically = int((pdf.h - 2 * 12.7) / box_height)
    boxes_per_page_vertically = 10
    top_margin = 12.7
    vertical_label_space = 0
    label_height = 25.4
    label_indent = 3

    pdf.add_page()
    page = 0

    for index, row in df.iterrows():
        # Check if we need to add a new page
        if index != 0 and index % (boxes_per_page_vertically * boxes_per_page_horizontally) == 0:
            pdf.add_page()
            page += 1

        # Calculate the position of the current text box
        col_index = index % boxes_per_page_horizontally
        label_space = col_index * label_width
        between_space = left_margin + col_index * horizontal_label_space
        x = label_space + between_space + label_indent

        label_on_page_index = index - page * boxes_per_page_horizontally * boxes_per_page_vertically
        row_index = label_on_page_index // boxes_per_page_horizontally
        label_space = row_index * label_height
        between_space = top_margin + row_index * vertical_label_space
        y = label_space + between_space

        font_size = 11
        pdf.set_font("Arial", size=font_size)
        points_per_inch = 72
        inch_to_mm = 25.4
        points_to_mm = inch_to_mm / points_per_inch
        line_height = 1.3 * font_size * points_to_mm
        line_space = line_height - font_size * points_to_mm

        # Count the number of lines in the address and figure out the additional margin to set inside the label
        address_text = row['mailing_address']
        line_count = address_text.count('\n') + 1
        label_margin = max(1, (label_height - line_count * line_height) / 2.0)
        y += label_margin  # would be nice if it could be set within multi_cell
        y += line_space / 2.0

        pdf.set_xy(x, y)
        pdf.multi_cell(label_width, line_height, txt=row['mailing_address'], align='L')

    pdf_label_file_name = "ActiveMemberLabels.pdf"
    pdf_label_file_path = global_download_dir + "/" + pdf_label_file_name
    pdf.output(pdf_label_file_path, 'F')


def normalized_mailing_address(row):
    nma = (normalize_address1(row[addr1_col]) + '|' +
           normalize_address2(row[addr2_col]) + '|' +
           normalize_city(row[city_col]) + '|' +
           normalize_state(row[state_col]) + '|' +
           normalize_zip(row[zip_col]))
    return nma

def normalize_address1(address1):
    address1 = clean(address1)
    address1 = address1.replace('avenue', 'ave')
    return address1


def normalize_address2(address2):
    address2 = clean(address2)
    return address2


def normalize_city(city):
    city = clean(city)
    return city


def normalize_state(state):
    state = clean(state)
    return state


def normalize_zip(zip):
    zip = clean(zip)
    zip = zip.split('-')[0]
    return zip


def clean(adr_string):
    adr_string = adr_string.strip()
    adr_string = adr_string.lower()
    adr_string = ' '.join(adr_string.split())
    return adr_string


if __name__ == "__main__":
    global global_out_file
    global_download_dir = str(Path.home() / "Downloads")

    create_mailing_labels()
