import tkinter as tk
from pathlib import Path
from tkinter import filedialog

import pandas as pd
from fpdf import FPDF

global global_download_dir

mail_name_col = 'mailing_name'
last_name_col = 'last_name'
addr1_col = 'address1'
addr2_col = 'address2'
city_col = 'city'
state_col = 'state'
zip_col = 'zip'


class MailingLabelCreator:

    def __init__(self):
        self.down_load_dir = str(Path.home() / "Downloads")
        self.pdf_label_file = "ActiveMembers.pdf"
        self.default_club_express_export_file_name = "ActiveMembersExtra.csv"
        self.club_express_export_file = None

        # Default values are initialize for 1 x 2 5/8" labels #612-221
        # all values in inches
        self.label_name = ''
        self.page_height = 11
        self.page_width = 8.5
        self.number_across = 3
        self.number_down = 10
        self.labels_per_page = self.number_across * self.number_down
        self.label_height = 1.0
        self.label_width = 2.625
        self.top_margin = 0.5
        self.side_margin = 0.137
        self.vertical_pitch = 0
        self.horizontal_pitch = 0.118

        # indentation for the label text
        self.text_indent = 0.137

    def get_file_from_user(self):
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
                "  -> Typically this will create a file called \"ActiveMembersExtra.csv in the downloads folder\n" +
                "- Select the generated file\n")
        text.pack()
        text.insert(tk.END, hint)

        # Create a button
        button = tk.Button(window, text="Open File", command=lambda: self.open_file(window))
        button.pack()

        # Run the application
        window.mainloop()

        return self.club_express_export_file

    def open_file(self, window):
        file_name = filedialog.askopenfile(initialdir=self.down_load_dir,
                                           initialfile=self.default_club_express_export_file_name)
        self.club_express_export_file = file_name
        window.destroy()

    def create_mailing_labels(self):
        input_file = self.get_file_from_user()
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

        df['normalized_mailing_address'] = df.apply(lambda arow: self.normalized_mailing_address(arow), axis=1)

        df_unique_address = df.drop_duplicates(subset=['normalized_mailing_address'], keep='first')
        df_unique_address.reset_index(drop=True, inplace=True)

        df_dropped = df.duplicated(subset=[addr1_col, addr2_col, city_col, state_col, zip_col], keep='last')

        df = df_unique_address
        # Formatting for Office Depot 1" x 2-5/8" inch labels, 612-221, 3 rows and 10 columns per sheet

        # Create a PDF in landscape orientation on letter size paper
        pdf = FPDF('L', 'in', (self.page_height, self.page_width))
        pdf.set_auto_page_break(False)

        # Set the margins for the pdf file to the margins of the label sheet
        pdf.set_top_margin(self.top_margin)
        pdf.set_left_margin(self.side_margin + self.text_indent)
        pdf.set_right_margin(0.0)

        pdf.add_page()
        page = 0

        for index, row in df.iterrows():
            # Check if we need to add a new page
            if index != 0 and index % (self.number_down * self.number_across) == 0:
                pdf.add_page()
                page += 1

            # Calculate the position of the current text box
            col_index = index % self.number_across
            label_space = col_index * self.label_width
            between_space = self.side_margin + col_index * self.horizontal_pitch
            x = label_space + between_space + self.text_indent

            label_on_page_index = index - page * self.labels_per_page
            row_index = label_on_page_index // self.number_across
            label_space = row_index * self.label_height
            between_space = self.top_margin + row_index * self.vertical_pitch
            y = label_space + between_space

            font_size = 11
            pdf.set_font("Arial", size=font_size)
            points_per_inch = 72
            points_to_inch = 1.0 / points_per_inch
            line_height = font_size * points_to_inch
            line_space = 0.4 * line_height

            # Count the number of lines in the address and figure out the additional vertical margin to set inside the label
            address_text = row['mailing_address']
            line_count = address_text.count('\n') + 1
            label_indent_vertical = ( self.label_height - (line_count * line_height + (line_count-1) * line_space) ) / 2.0
            label_indent_vertical = max(0.05, label_indent_vertical)
            y += label_indent_vertical

            pdf.set_xy(x, y)
            pdf.multi_cell(self.label_width, line_height+line_space, txt=row['mailing_address'], align='L')

        pdf_label_file_path = self.down_load_dir + "/" + self.pdf_label_file
        pdf.output(pdf_label_file_path, 'F')

    def normalized_mailing_address(self, row):
        nma = (self.normalize_address1(row[addr1_col]) + '|' +
               self.normalize_address2(row[addr2_col]) + '|' +
               self.normalize_city(row[city_col]) + '|' +
               self.normalize_state(row[state_col]) + '|' +
               self.normalize_zip(row[zip_col]))
        return nma

    def normalize_address1(self, address1):
        address1 = self.clean(address1)
        address1 = address1.replace('avenue', 'ave')
        return address1

    def normalize_address2(self, address2):
        address2 = self.clean(address2)
        return address2

    def normalize_city(self, city):
        city = self.clean(city)
        return city

    def normalize_state(self, state):
        state = self.clean(state)
        return state

    def normalize_zip(self, zip_code):
        zip_code = self.clean(zip_code)
        zip_code = zip_code.split('-')[0]
        return zip_code

    @staticmethod
    def clean(adr_string):
        adr_string = adr_string.strip()
        adr_string = adr_string.lower()
        adr_string = ' '.join(adr_string.split())
        return adr_string


if __name__ == "__main__":
    label_creator = MailingLabelCreator()
    label_creator.create_mailing_labels()
