import argparse
import pandas as pd
from fpdf import FPDF


def create_mailing_address(input_file, output_file):


    df = pd.read_csv(input_file, sep=',')  # Use ',' as the separator
    df.columns = df.columns.str.strip()  # Strip leading and trailing spaces

    mail_name_col = 'mailing_name'
    addr1_col = 'address1'
    addr2_col = 'address2'
    city_col = 'city'
    state_col = 'state'
    zip_col = 'zip'


    # Convert all the fields to strings
    for col in [mail_name_col, addr1_col, addr2_col, city_col, state_col, zip_col]:
        df[col] = df[col].astype(str)

    # Replace 'nan' with an empty string
    df.replace('nan', '', inplace=True)

    # Sort the DataFrame by the OWNER1 column
    df.sort_values(by=[mail_name_col], inplace=False)

    # Create the mailing address
    df['mailing_address'] = df.apply(lambda row: '\n'.join(filter(None, [row[mail_name_col], ', '.join(
        filter(None, [row[addr1_col], row[addr2_col]])), ', '.join([row[city_col], row[state_col], row[zip_col]])])),
                                     axis=1)

    # Formatting for Office Depot 1" x 2-5/8" inch labels, 612-221, 3 rows and 10 colums per sheet

    pdf = FPDF('L', 'mm', (279.4, 215.9))  # Create a PDF in landscape orientation on letter size paper
    pdf.set_auto_page_break(False)
    # Dimensions of the text box in mm
    # box_height = 28.575
    # box_width = 114.3

    # Set the margins to 1 inch (approximately 25.4 mm) at the top and bottom and 0.5 inch (approximately 12.7 mm) at the left
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
        x = label_space + between_space

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
        y += label_margin # would be nice if it could be set within multi_cell
        y += line_space / 2.0

        pdf.set_xy(x, y)
        pdf.multi_cell(label_width, line_height, txt=row['mailing_address'], align='L')

    pdf.output(output_file)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Create mailing addresses from a CSV file.')
    parser.add_argument('input_file', type=str, help='Input CSV file name')
    parser.add_argument('output_file', type=str, help='Output PDF file name')

    # args = parser.parse_args()

    # input_file = parser.parse_args().input_file
    # output_file = parser.parse_args().output_file
    input_file = 'ActiveMembers.csv'
    output_file = 'MailingLabels.pdf'
    create_mailing_address(input_file, output_file)