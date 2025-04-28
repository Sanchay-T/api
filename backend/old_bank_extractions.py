import re
import os
import io
import pandas as pd
import numpy as np
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
import shutil
import logging
import pdfplumber
from datetime import datetime
import regex as re
import datefinder
from calendar import monthrange
import calendar
from openpyxl.styles import Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl import Workbook, load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
# from findaddy.exceptions import ExtractionError
from .utils import get_saved_pdf_dir
TEMP_SAVED_PDF_DIR = get_saved_pdf_dir()


class CustomStatement:
    def __init__(self, bank_name, pdf_path, pdf_password, CA_ID):
        self.writer = None
        self.bank_names = bank_name
        self.pdf_paths = pdf_path
        self.pdf_passwords = pdf_password
        self.account_number = ""
        self.file_name = None
        self.CA_ID = CA_ID

    def add_top_line_to_pdf(self, input_pdf_path, output_pdf_path):
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()

        for page in input_pdf.pages:
            page_width = page.mediabox.right
            page_height = page.mediabox.top

            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setStrokeColorRGB(0, 0, 0)  # Black color
            can.setLineWidth(4)  # Thick line
            can.line(0, page_height, page_width, page_height)  # Draw line at the top
            can.save()

            packet.seek(0)
            top_line_pdf = PdfReader(packet)
            top_line_page = top_line_pdf.pages[0]
            top_line_page.mediabox.upper_right = (page_width, page_height)

            page.merge_page(top_line_page)
            output_pdf.add_page(page)

        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)

        return output_pdf_path

    def add_lines_to_pdf(self, input_pdf_path, output_pdf_path, timestamp):
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()

        for page in input_pdf.pages:
            page_width = page.mediabox.right
            page_height = page.mediabox.top

            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setStrokeColorRGB(0, 0, 0)  # Black color
            can.setLineWidth(1)  # Adjust thickness as needed

            # Draw horizontal lines at 1 cm intervals
            for y_position in range(0, int(page_height), 6):  # 1 cm = 28.35 points (assuming 1 inch = 72 points)
                can.line(0, y_position, page_width, y_position)

            can.save()

            packet.seek(0)
            line_pdf = PdfReader(packet)
            line_page = line_pdf.pages[0]
            line_page.mediabox.upper_right = (page_width, page_height)

            page.merge_page(line_page)
            output_pdf.add_page(page)

        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)

        return output_pdf_path

    def insert_all_separators(self, page, page_width, page_height):

        lines = page.extract_text().split('\n')

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFillColorRGB(0, 0, 0)
        can.setStrokeColorRGB(0, 0, 0)
        can.setLineWidth(2)
        line_height = page_height / (len(lines) + 1)

        for line_num, line in enumerate(lines):
            if line.strip():
                y_position = line_num * line_height * 0.8
                can.line(0, y_position, page_width, y_position)
        can.save()

        packet.seek(0)
        sep_pdf = PdfReader(packet)
        sep_page = sep_pdf.pages[0]
        sep_page.mediabox.upper_right = (page_width, page_height)
        page.merge_page(sep_page)

    def insert_all_separators_idbi(self, page, page_width, page_height):

        lines = page.extract_text().split('\n')

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFillColorRGB(0, 0, 0)
        can.setStrokeColorRGB(0, 0, 0)
        can.setLineWidth(2)
        line_height = page_height / (len(lines) + 1)

        for line_num, line in enumerate(lines):
            if line.strip():
                y_position = line_num * line_height
                can.line(0, y_position, page_width, y_position)
        can.save()

        packet.seek(0)
        sep_pdf = PdfReader(packet)
        sep_page = sep_pdf.pages[0]
        sep_page.mediabox.upper_right = (page_width, page_height)
        page.merge_page(sep_page)

    def separate_lines_in_pdf_idbi(self, input_pdf_path, timestamp):
        """Inserts separators between lines in the PDF."""
        CA_ID = self.CA_ID
        unlocked_pdf_filename = f"{timestamp}-{CA_ID}.pdf"
        output_pdf_path = os.path.join(TEMP_SAVED_PDF_DIR, unlocked_pdf_filename)
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()

        for page in input_pdf.pages:
            page_width = page.mediabox.upper_right[0] - page.mediabox.lower_left[0]
            page_height = page.mediabox.upper_right[1] - page.mediabox.lower_left[1]

            self.insert_all_separators_idbi(page, page_width, page_height)

            output_pdf.add_page(page)

        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)

        return output_pdf_path

    def insert_all_separators_uco(self, page, page_width, page_height):

        lines = page.extract_text().split('\n')

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFillColorRGB(0, 0, 0)
        can.setStrokeColorRGB(0, 0, 0)
        can.setLineWidth(2)
        line_height = page_height / (len(lines) + 1)

        for line_num, line in enumerate(lines):
            if line.strip():
                y_position = line_num * line_height
                can.line(0, y_position, page_width, y_position)
        can.save()

        packet.seek(0)
        sep_pdf = PdfReader(packet)
        sep_page = sep_pdf.pages[0]
        sep_page.mediabox.upper_right = (page_width, page_height)
        page.merge_page(sep_page)

    def separate_lines_in_pdf_uco(self, input_pdf_path, timestamp):
        """Inserts separators between lines in the PDF."""
        CA_ID = self.CA_ID
        unlocked_pdf_filename = f"{timestamp}-{CA_ID}.pdf"
        output_pdf_path = os.path.join(TEMP_SAVED_PDF_DIR, unlocked_pdf_filename)
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()

        for page in input_pdf.pages:
            page_width = page.mediabox.upper_right[0] - page.mediabox.lower_left[0]
            page_height = page.mediabox.upper_right[1] - page.mediabox.lower_left[1]

            self.insert_all_separators_uco(page, page_width, page_height)

            output_pdf.add_page(page)

        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)

        return output_pdf_path

    def separate_lines_in_pdf(self, input_pdf_path, timestamp):
        """Inserts separators between lines in the PDF."""
        CA_ID = self.CA_ID
        unlocked_pdf_filename = f"{timestamp}-{CA_ID}.pdf"
        output_pdf_path = os.path.join(TEMP_SAVED_PDF_DIR, unlocked_pdf_filename)
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()

        for page in input_pdf.pages:
            page_width = page.mediabox.upper_right[0] - page.mediabox.lower_left[0]
            page_height = page.mediabox.upper_right[1] - page.mediabox.lower_left[1]

            self.insert_all_separators(page, page_width, page_height)

            output_pdf.add_page(page)

        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)

        return output_pdf_path

    def insert_all_separators_scb(self, page, page_width, page_height):

        lines = page.extract_text().split('\n')

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFillColorRGB(0, 0, 0)
        can.setStrokeColorRGB(0, 0, 0)
        can.setLineWidth(2)
        line_height = page_height / (len(lines) + 1)

        for line_num, line in enumerate(lines):
            if line.strip():
                y_position = line_num * line_height
                can.line(0, y_position, page_width, y_position)
        can.save()

        packet.seek(0)
        sep_pdf = PdfReader(packet)
        sep_page = sep_pdf.pages[0]
        sep_page.mediabox.upper_right = (page_width, page_height)
        page.merge_page(sep_page)

    def separate_lines_in_pdf_scb(self, input_pdf_path, timestamp):
        """Inserts separators between lines in the PDF."""
        CA_ID = self.CA_ID
        unlocked_pdf_filename = f"{timestamp}-{CA_ID}.pdf"
        output_pdf_path = os.path.join(TEMP_SAVED_PDF_DIR, unlocked_pdf_filename)
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()

        for page in input_pdf.pages:
            page_width = page.mediabox.upper_right[0] - page.mediabox.lower_left[0]
            page_height = page.mediabox.upper_right[1] - page.mediabox.lower_left[1]

            self.insert_all_separators_scb(page, page_width, page_height)

            output_pdf.add_page(page)

        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)

        return output_pdf_path

    def insert_vertical_lines(self, page, x_positions, page_height):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFillColorRGB(0, 0, 0)  # Set fill color to black
        can.setStrokeColorRGB(0, 0, 0)  # Set stroke color to black
        can.setLineWidth(2)  # Set line width
        for x_position in x_positions:
            can.line(x_position, 0, x_position, page_height)
        can.save()
        packet.seek(0)
        sep_pdf = PdfReader(packet)
        sep_page = sep_pdf.pages[0]
        sep_page.mediabox.upper_right = (page.mediabox.upper_right[0], page_height)
        page.merge_page(sep_page)

    def separate_lines_in_vertical_pdf(self, input_pdf_path, x_positions, timestamp):
        CA_ID = self.CA_ID
        unlocked_pdf_filename = f"{timestamp}-{CA_ID}.pdf"
        output_pdf_path = os.path.join(TEMP_SAVED_PDF_DIR, unlocked_pdf_filename)
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()
        for page_num in range(len(input_pdf.pages)):
            page = input_pdf.pages[page_num]
            page_width = page.mediabox.upper_right[0] - page.mediabox.lower_left[0]
            page_height = page.mediabox.upper_right[1] - page.mediabox.lower_left[1]
            content = page.extract_text()
            lines = content.split('\n')
            self.insert_vertical_lines(page, x_positions, page_height)
            output_pdf.add_page(page)
        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)
        return output_pdf_path

    def check_date(self, df):
        df.dropna(subset=['Value Date'], inplace=True)
        if pd.to_datetime(df['Value Date'].iloc[-1], dayfirst=True) < pd.to_datetime(df['Value Date'].iloc[0],
                                                                                     dayfirst=True):
            new_df = df[::-1].reset_index(drop=True)
            print("found in reverse")
        else:
            new_df = df.copy()  # No reversal required
        return new_df

    def display_value_date_str(self, df):
        df['Value Date'] = df['Value Date'].astype(str)
        df['Value Date'] = df['Value Date'].str.split(' ').str[0].str.split('-').str[2] + '-' + \
                           df['Value Date'].str.split(' ').str[0].str.split('-').str[1] + '-' + \
                           df['Value Date'].str.split(' ').str[0].str.split('-').str[0]
        return df

    def extract_the_df(self, idf):
        balance_row_index = idf[idf.apply(lambda row: 'balance' in ' '.join(row.astype(str)).lower(), axis=1)].index

        # Check if "Balance" row exists
        if not balance_row_index.empty:
            # Get the index of the "Balance" row
            balance_row_index = balance_row_index[0]
            # Create a new DataFrame from the "Balance" row till the end
            new_df = idf.iloc[balance_row_index:]
        else:
            return idf
        return new_df

    def uncontinuous(self, df):
        df = df[~df.apply(lambda row: row.astype(str).str.contains('Balance', case=False)).any(axis=1)]
        return df

    def extract_dates_from_pdf(self, unlocked_file_path):
        with open(unlocked_file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ''
            for page in pdf_reader.pages:
                text += page.extract_text()
        return text

    def convert_to_dt_format(self, date_str):
        formats_to_try = ["%d-%m-%Y", "%d %b %Y", "%Y-%m-%d", "%d %B %Y", "%d/%m/%Y", "%d/%m/%Y",
                          "%d-%m-%Y", "%d-%b-%Y", "%d/%m/%Y", "%d-%b-%y", "%B %d, %Y", "%d-%B-%Y",
                          "%d-%m-%Y", "%d %b %Y", "%Y-%m-%d", "%d %B %Y", "%d/%m/%Y", "%d/%m/%Y",
                          "%d-%m-%Y", "%d-%b-%Y", "%d/%m/%Y", "%d-%b-%y", "%B %d, %Y", "%m/%d/%Y",
                          "%d %b %y", "%d/%m/%y"]
        for format_str in formats_to_try:
            try:
                date_obj = datetime.strptime(date_str, format_str)
                return date_obj.strftime("%d-%m-%Y")
            except ValueError:
                # pass
                raise ValueError("Invalid date format: {}".format(date_str))

    #################--------******************----------#####################
    def unlock_the_pdfs_path(self, pdf_path, pdf_password, bank_name, timestamp):
        CA_ID = self.CA_ID
        logger = logging.getLogger(self.CA_ID)
        os.makedirs(TEMP_SAVED_PDF_DIR, exist_ok=True)

        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            if pdf_reader.is_encrypted:
                pdf_reader.decrypt(pdf_password)
                try:
                    _ = pdf_reader.numPages  # Check if decryption was successful
                    pdf_writer = PyPDF2.PdfWriter()
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)
                    unlocked_pdf_filename = f"{timestamp}-{CA_ID}.pdf"
                    unlocked_pdf_path = os.path.join(TEMP_SAVED_PDF_DIR, unlocked_pdf_filename)
                    with open(unlocked_pdf_path, 'wb') as unlocked_pdf_file:
                        pdf_writer.write(unlocked_pdf_file)
                except Exception as e:
                    # raise ExtractionError("Incorrect password. Unable to unlock the PDF.")
                    raise ValueError("Incorrect password. Unable to unlock the PDF.")
            else:
                # Copy the PDF file to the "saved_pdf" folder without modification
                unlocked_pdf_filename = f"{timestamp}-{CA_ID}.pdf"
                unlocked_pdf_path = os.path.join(TEMP_SAVED_PDF_DIR, unlocked_pdf_filename)
                with open(pdf_path, 'rb') as unlocked_pdf_file:
                    with open(unlocked_pdf_path, 'wb') as output_file:
                        output_file.write(unlocked_pdf_file.read())

        reader = PdfReader(unlocked_pdf_path)
        number_of_pages = len(reader.pages)
        page = reader.pages[0]
        try:
            text = page.extract_text()[0]
        except Exception as e:
            raise ExtractionError(
                "The pdf you have uploaded is of image-only(non-text) format. Please upload a text pdf.")

        unlocked_pdf_path = self.add_top_line_to_pdf(unlocked_pdf_path, unlocked_pdf_path)

        return unlocked_pdf_path

    def idbi(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside IDBI Bank")

            def idbi_format_1(unlocked_pdf_path):
                try:

                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)

                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = df.drop(df.columns[0:2], axis=1)  # Removes column at position 1
                    df = df.rename(
                        columns={2: 'Value Date', 3: 'Description', 4: 'Cheque No', 5: 'CR/DR', 6: 'CCY',
                                 7: 'Amount (INR)',
                                 8: 'Balance'})

                    df['Credit'] = 0
                    df['Debit'] = 0
                    df.loc[df['CR/DR'] == 'Cr.', 'Credit'] = df.loc[df['CR/DR'] == 'Cr.', 'Amount (INR)']
                    df.loc[df['CR/DR'] != 'Cr.', 'Debit'] = df.loc[df['CR/DR'] != 'Cr.', 'Amount (INR)']
                    df = df.drop(['CR/DR', 'Amount (INR)'], axis=1)
                    # Replace '/' with '-' in the 'Value Date' column
                    df['Value Date'] = df['Value Date'].astype(str).str.replace('/', '-', n=2)
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'IDBI Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def idbi_format_2(unlocked_pdf_path):
                try:
                    unlocked_pdf_path = self.separate_lines_in_pdf_idbi(unlocked_pdf_path, timestamp)
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    df = self.uncontinuous(w)
                    df = df.drop([0, 3, 7, 8, 9, 10, 11, 12, 13], axis=1)
                    df = df.rename(columns={1: 'Value Date', 2: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')

                    df['new_column'] = np.nan
                    counter = 0
                    # Iterate over the dataframe rows
                    for index, row in df.iterrows():
                        if pd.notnull(row["Value Date"]):
                            counter += 1
                        df.at[index, 'new_column'] = counter
                    # Iterate over the dataframe rows
                    for index, row in df.iterrows():
                        if pd.isna(row["Value Date"]):
                            df.at[index, 'new_column'] = np.NaN
                    df['new_column'].fillna(method='ffill', inplace=True)
                    df["Description"].fillna('', inplace=True)
                    df["Description"] = df.groupby('new_column')["Description"].transform(lambda x: ' '.join(x))
                    df = df.drop_duplicates(subset='new_column').reset_index(drop=True)
                    df = df.drop(["new_column"], axis=1)

                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'IDBI Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = idbi_format_1(unlocked_pdf_path)
            if result is None:
                result = idbi_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for IDBI bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def axis(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Axis Bank")

            def axis_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.rename(
                        columns={0: 'Value Date', 1: 'Cheque No', 2: 'Description', 3: 'Debit', 4: 'Credit',
                                 5: 'Balance',
                                 6: 'Init(Br)'})
                    df = df[:-1]
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Axis Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def axis_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([1, 7], axis=1)
                    df = df.rename(
                        columns={0: 'Value Date', 2: 'Description', 4: 'Amount', 5: 'CR/DR', 6: 'Balance'})
                    df = df[:-1]
                    df['Credit'] = 0
                    df['Debit'] = 0
                    df.loc[df['CR/DR'] == 'CR', 'Credit'] = df.loc[df['CR/DR'] == 'CR', 'Amount']
                    df.loc[df['CR/DR'] != 'CR', 'Debit'] = df.loc[df['CR/DR'] != 'CR', 'Amount']
                    df = df.drop(['CR/DR', 'Amount'], axis=1)
                    df.reset_index(inplace=True)
                    print(df)
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Axis Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = axis_format_2(unlocked_pdf_path)
            if result is None:
                result = axis_format_1(unlocked_pdf_path)

            logger.info("PDF to table Extraction for Axis bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def sbi(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()

            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)

            # start custom extraction
            df = df.drop(df.columns[1:2], axis=1)  # Removes column at position 1
            df = df.rename(
                columns={0: 'Value Date', 2: 'Description', 3: 'Cheque No', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            # date
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            # Reorder the columns
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'SBI Bank'
            logger.info("PDF to table Extraction for SBI bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def idfc(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside IDFC Bank")
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()

            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)

            # start custom extraction
            df = df.drop(df.columns[1:2], axis=1)  # Removes column at position 1
            df = df.rename(
                columns={0: 'Value Date', 2: 'Description', 3: 'Cheque No', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            # date
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])
            # Reorder the columns
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'IDFC Bank'
            logger.info("PDF to table Extraction for IDFC bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def pnb(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside PNB Bank")

            def pnb_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = df.rename(
                        columns={0: 'Value Date', 1: 'Cheque No', 2: 'Debit', 3: 'Credit', 4: 'Balance',
                                 5: 'Description'})
                    df['Value Date'] = df['Value Date'].astype(str).str.replace('/', '-', n=2)
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].replace({' Cr.': '', ' Dr.': ''}, regex=True)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'PNB Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def pnb_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside PNB Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    df = df.drop(df.columns[[0, 3, 4, 8]], axis=1)
                    # start custom extraction
                    df = df.rename(
                        columns={1: 'Value Date', 2: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    # df['Value Date'] = df['Value Date'].astype(str).str.replace('/', '-', n=2)
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].replace({' Cr.': '', ' Dr.': ''}, regex=True)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'PNB Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            def pnb_format_3(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df[6] = df[2].where(df[3] == 'DR', None)
                    df[7] = df[2].where(df[3] == 'CR', None)
                    df = df.drop([1, 2, 3], axis=1)
                    df = df.rename(columns={0: 'Value Date', 5: 'Description', 6: 'Debit', 7: 'Credit', 4: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')

                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'PNB Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 3: {e}")
                    return None

            def pnb_format_4(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside PNB Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([2], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.extract('(\d+\.?\d*)')
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'PNB Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 3: {e}")
                    return None

            result = pnb_format_1(unlocked_pdf_path)
            if result is None:
                result = pnb_format_2(unlocked_pdf_path)
                if result is None:
                    result = pnb_format_3(unlocked_pdf_path)
                    if result is None:
                        result = pnb_format_4(unlocked_pdf_path)

            logger.info("PDF to table Extraction for PNB bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def yes_bank(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Yes Bank")

            def yes_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Yes Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = df.drop(df.columns[1:2], axis=1)  # Removes column at position 1
                    # df = df.iloc[1:] # Removes the first 1 rows
                    df = self.uncontinuous(df)
                    df = df.rename(
                        columns={0: 'Value Date', 2: 'Cheque No', 3: 'Description', 4: 'Debit', 5: 'Credit',
                                 6: 'Balance'})
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Yes Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def yes_format_2(unlocked_pdf_path):

                try:

                    def extract_specific_table_adjusted(pdf_path):
                        # Adjusted desired columns based on the extracted sample table headers
                        desired_columns = ["Transaction\nDate", "Value Date", "Cheque No/Reference No", "Withdrawals",
                                           "Deposits", "Running Balance"]
                        filtered_tables = []
                        with pdfplumber.open(pdf_path) as pdf:
                            total_pages = len(pdf.pages)
                            for page_num in range(total_pages):
                                tables = pdf.pages[page_num].extract_tables()
                                for table in tables:
                                    cleaned_columns = [col.strip() if col else "" for col in table[0]]
                                    if set(desired_columns).issubset(set(cleaned_columns)):
                                        df = pd.DataFrame(table[1:], columns=cleaned_columns)
                                        filtered_tables.append(df)
                        if not filtered_tables:
                            return None
                        return pd.concat(filtered_tables, ignore_index=True)

                    adjusted_specific_tables_df = extract_specific_table_adjusted(unlocked_pdf_path)
                    adjusted_specific_tables_df = adjusted_specific_tables_df.dropna(axis=1, how="all")
                    w = adjusted_specific_tables_df

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    df = df.drop(df.columns[[0, 2]], axis=1)
                    # start custom extraction
                    df = df.rename(
                        columns={"Withdrawals": 'Debit', "Deposits": 'Credit', "Running Balance": 'Balance'})

                    # df['Value Date'] = df['Value Date'].astype(str).str.replace('/', '-', n=2)
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].replace({' Cr.': '', ' Dr.': ''}, regex=True)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Yes Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            def yes_format_3(unlocked_pdf_path):

                try:
                    print("Inside Yes Bank")

                    def extract_specific_table_adjusted_v2(pdf_path):
                        desired_columns = ["Transaction\nDate", "ValueDate", "Description", "Withdrawals", "Deposits",
                                           "Balance"]
                        filtered_tables = []
                        with pdfplumber.open(pdf_path) as pdf:
                            total_pages = len(pdf.pages)
                            for page_num in range(total_pages):
                                tables = pdf.pages[page_num].extract_tables()
                                for table in tables:
                                    cleaned_columns = [col.strip() if col else "" for col in table[0]]
                                    if set(desired_columns).issubset(set(cleaned_columns)):
                                        df = pd.DataFrame(table[1:], columns=cleaned_columns)
                                        filtered_tables.append(df)
                        if not filtered_tables:
                            return None
                        return pd.concat(filtered_tables, ignore_index=True)

                    df = extract_specific_table_adjusted_v2(unlocked_pdf_path)
                    df = df.rename(columns={'ValueDate': 'Value Date', 'Withdrawals': 'Debit', 'Deposits': 'Credit'})
                    df = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Yes Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 3: {e}")
                    return None

            result = yes_format_1(unlocked_pdf_path)
            if result is None:
                result = yes_format_2(unlocked_pdf_path)
                if result is None:
                    result = yes_format_3(unlocked_pdf_path)

            logger.info("PDF to table Extraction for Yes bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def union(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Union Bank")

            def union_format_1(unlocked_pdf_path):

                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Union Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = df.drop(df.columns[[2, 3, 4]], axis=1)  # Removes column at position 3,5
                    # df = df.iloc[1:] # Removes the first 1 rows
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    # date
                    if ' ' in df['Value Date'].unique()[0]:
                        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y %H:%M:%S')
                        df['Value Date'] = df['Value Date'].dt.date
                        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%Y-%m-%d',
                                                          errors='coerce').dt.strftime(
                            '%d-%m-%Y')
                    else:
                        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y',
                                                          errors='coerce').dt.strftime(
                            '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Union Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def union_format_2(unlocked_pdf_path):

                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Union Bank")
                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = df.drop(df.columns[[0, 2]], axis=1)  # Removes column at position 3,5
                    # df = df.iloc[1:] # Removes the first 1 rows
                    df = df.rename(columns={1: 'Value Date', 3: 'Description', 4: 'Transaction', 5: 'Balance'})
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Credit'] = df['Transaction'].str.extract(r'(\d+\.\d+) \(Cr\)').astype(float)
                    df['Debit'] = df['Transaction'].str.extract(r'(\d+\.\d+) \(Dr\)').astype(float)
                    df = df.drop(columns=['Transaction'])
                    df = df.dropna(subset=['Value Date'])
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True).astype(float)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Union Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = union_format_1(unlocked_pdf_path)
            if result is None:
                result = union_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for Union bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def kotak(self, unlocked_pdf_path, timestamp):
        path = [f for f in os.listdir(TEMP_SAVED_PDF_DIR) if f.endswith('.pdf')][0]
        unlocked_pdf_path = f"{TEMP_SAVED_PDF_DIR}\{path}"

        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Kotak Bank")

            def kotak_format_1(unlocked_pdf_path):
                try:
                    directory = os.path.dirname(unlocked_pdf_path)
                    new_pdf_path = os.path.join(directory, "one.pdf")
                    shutil.copy(unlocked_pdf_path, new_pdf_path)

                    one_pdf_path = self.separate_lines_in_pdf_idbi(new_pdf_path, timestamp)

                    x_positions = [20, 86, 260, 350, 480, 580, 660]
                    unlocked_pdf_path = self.separate_lines_in_vertical_pdf(one_pdf_path, x_positions, timestamp)

                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Kotak Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)

                    w = df_total.drop_duplicates()
                    df = w.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Amount', 4: 'Balance'})
                    df = df[df['Description'] != '']

                    df.loc[df['Amount'].str.contains('Cr'), 'Credit'] = df.loc[
                        df['Amount'].str.contains('Cr'), 'Amount'].str.replace(' Cr', '')
                    df.loc[df['Amount'].str.contains('Dr'), 'Debit'] = df.loc[
                        df['Amount'].str.contains('Dr'), 'Amount'].str.replace(' Dr', '')
                    df = df.drop(['Amount'], axis=1)

                    df["new_column"] = (df["Value Date"].notna()).cumsum()
                    df.loc[df["Value Date"].isna(), "new_column"] = np.nan
                    df["new_column"].fillna(method="ffill", inplace=True)
                    df["Description"].fillna("", inplace=True)
                    df["Description"] = df.groupby("new_column")["Description"].transform(lambda x: " ".join(x))
                    df = df.drop_duplicates(subset="new_column").reset_index(drop=True)

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df.dropna(subset=['Description'])
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Kotak Bank'

                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def kotak_format_2(unlocked_pdf_path):
                try:
                    one_pdf_path = self.separate_lines_in_pdf_idbi(f"{TEMP_SAVED_PDF_DIR}/one.pdf", timestamp)

                    x_positions = [80, 200, 380, 500, 600, 690, 780]

                    unlocked_pdf_path = self.separate_lines_in_vertical_pdf(one_pdf_path, x_positions, timestamp)

                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Kotak Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)

                    w = df_total.drop_duplicates()
                    df = w.rename(columns={1: 'Value Date', 3: 'Description', 8: 'Amount', 9: 'CR/DR', 12: 'Balance'})
                    df = df[df['Description'] != '']

                    df['Credit'] = 0
                    df['Debit'] = 0
                    df.loc[df['CR/DR'] == 'CR', 'Credit'] = df.loc[df['CR/DR'] == 'CR', 'Amount']
                    df.loc[df['CR/DR'] != 'CR', 'Debit'] = df.loc[df['CR/DR'] != 'CR', 'Amount']
                    df = df.drop(['CR/DR', 'Amount'], axis=1)

                    df["new_column"] = np.nan
                    counter = 0
                    for index, row in df.iterrows():
                        if pd.notnull(row["Value Date"]):
                            counter += 1
                        df.at[index, "new_column"] = counter

                    for index, row in df.iterrows():
                        if pd.isna(row["Value Date"]):
                            df.at[index, "new_column"] = np.NaN

                    df["new_column"].fillna(method="ffill", inplace=True)
                    df["Description"].fillna("", inplace=True)
                    df["Description"] = df.groupby("new_column")["Description"].transform(lambda x: " ".join(x))
                    df = df.drop_duplicates(subset="new_column").reset_index(drop=True)

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df.dropna(subset=['Description'])
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Kotak Bank'

                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = kotak_format_1(unlocked_pdf_path)
            if result is None:
                result = kotak_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for Kotak bank completed")

            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def bob(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside BOB bank")

            def bob_format_1(unlocked_pdf_path):
                x_positions = [80, 360, 460, 590, 680]
                unlocked_pdf_path = self.separate_lines_in_vertical_pdf(unlocked_pdf_path, x_positions, timestamp)

                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Bank of Baroda")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    print(df)

                    # start custom extraction
                    df = df.drop(df.columns[[2]], axis=1)  # Removes column at position 2
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Balance'])

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Bank of Baroda'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def bob_format_2(unlocked_pdf_path):
                x_positions = [43, 88, 130, 310, 370, 430, 500, 575]
                new_pdf_path = self.separate_lines_in_vertical_pdf(unlocked_pdf_path, x_positions, timestamp)
                unlocked_pdf_path = new_pdf_path

                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Bank of Baroda")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    print(df)
                    # start custom extraction
                    df = df.drop(df.columns[[0, 2, 4, 5, 6, 9]], axis=1)  # Removes column at position 2
                    df = df.rename(columns={1: 'Value Date', 3: 'Description', 7: 'Debit', 8: 'Credit', 10: 'Balance'})
                    df = df.dropna(subset=['Balance'])

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Bank of Baroda'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = bob_format_2(unlocked_pdf_path)
            if result is None:
                result = bob_format_1(unlocked_pdf_path)
            logger.info("PDF to table Extraction for BOB bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def icici(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside ICICI Bank")

            def icici_format_1(unlocked_pdf_path):
                try:
                    df = pd.DataFrame()

                    print("Inside ICICI Bank")

                    dfs = []
                    with pdfplumber.open(unlocked_pdf_path) as pdf:
                        num_pages = len(pdf.pages)
                        print("Number of Pages in PDF:", num_pages)
                        for i in range(num_pages):
                            page = pdf.pages[i]
                            table = page.extract_tables()
                            if len(table) > 0:
                                for tab in table:
                                    df = pd.DataFrame(tab)
                                    dfs.append(df)
                    df_total = pd.concat(dfs, ignore_index=True)
                    new_df = self.extract_the_df(df_total)
                    df = self.uncontinuous(new_df)
                    # #start custom extraction
                    df = df.rename(columns={1: 'Value Date', 4: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Balance'] = df['Balance'].str.replace('-', '')
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'].notna() & (df['Description'] != '') & (df['Description'] != 'None')]
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'ICICI Bank'

                    return idf


                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def icici_format_2(unlocked_pdf_path):
                try:
                    df = pd.DataFrame()
                    dfs = []

                    with pdfplumber.open(unlocked_pdf_path) as pdf:
                        num_pages = len(pdf.pages)
                        print("Number of Pages in PDF:", num_pages)

                        for i in range(num_pages):
                            page = pdf.pages[i]
                            table = page.extract_tables()
                            if len(table) > 0:
                                for tab in table:
                                    df = pd.DataFrame(tab)
                                    dfs.append(df)
                    df = pd.concat(dfs, ignore_index=True)
                    new_df = self.extract_the_df(df)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop(columns=[2, 5, 6])
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 7: 'Balance'})
                    # df = df[df['Description'].notnull()]
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[1:]
                    df.dropna(subset=['Value Date'], inplace=True)
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'ICICI Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            def icici_format_3(unlocked_pdf_path):

                try:
                    df = pd.DataFrame()
                    dfs = []

                    with pdfplumber.open(unlocked_pdf_path) as pdf:
                        num_pages = len(pdf.pages)
                        print("Number of Pages in PDF:", num_pages)

                        for i in range(num_pages):
                            page = pdf.pages[i]
                            table = page.extract_tables()
                            if len(table) > 0:
                                for tab in table:
                                    df = pd.DataFrame(tab)
                                    dfs.append(df)
                    df = pd.concat(dfs, ignore_index=True)
                    new_df = self.extract_the_df(df)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop(columns=[0, 1, 2, 4, 5])
                    df = df.rename(columns={3: 'Value Date', 6: 'Description', 7: 'Debit', 8: 'Credit', 9: 'Balance'})
                    # df = df[df['Description'].notnull()]
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%b/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df.dropna(subset=['Value Date'], inplace=True)
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'ICICI Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 3: {e}")
                    return None

            def icici_format_4(unlocked_pdf_path):

                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside ICICI Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = self.extract_the_df(df_total)
                    df = self.uncontinuous(w)
                    # start custom extraction
                    df = df.drop([0, 2, 3], axis=1)
                    df = df.rename(columns={1: 'Value Date', 4: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b- %Y',
                                                      errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'ICICI Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 4: {e}")
                    return None

            result = icici_format_1(unlocked_pdf_path)
            if result is None:
                result = icici_format_2(unlocked_pdf_path)
                if result is None:
                    result = icici_format_3(unlocked_pdf_path)
                    if result is None:
                        result = icici_format_4(unlocked_pdf_path)
            logger.info("PDF to table Extraction for ICICI bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def indus(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Indus Bank")

            def indus_format_1(unlocked_pdf_path):

                try:
                    df = pd.DataFrame()
                    dfs = []

                    print("Inside IndusInd Bank")

                    with pdfplumber.open(unlocked_pdf_path) as pdf:
                        num_pages = len(pdf.pages)
                        print("Number of Pages in PDF:", num_pages)

                        for i in range(num_pages):
                            page = pdf.pages[i]
                            table = page.extract_tables()

                            if len(table) > 0:
                                for tab in table:
                                    df = pd.DataFrame(tab)
                                    dfs.append(df)
                    df = pd.concat(dfs, ignore_index=True)
                    new_df = self.extract_the_df(df)
                    df = self.uncontinuous(new_df)

                    if 'INR' in df[2].values:
                        # Remove the second column
                        df = df.drop(columns=[2])
                        # Reset the column numbers
                        df.columns = range(len(df.columns))
                    # start custom extraction
                    df = df.iloc[1:]  # Removes the first 1 rows
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df[df['Description'].notnull()]
                    # date

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df.dropna(subset=['Value Date']).reset_index(drop=True)
                    df = df[df['Description'] != '']
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'IndusInd Bank'
                    return idf


                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def indus_format_2(unlocked_pdf_path):

                try:
                    df = pd.DataFrame()
                    dfs = []
                    print("Inside IndusInd Bank")

                    with pdfplumber.open(unlocked_pdf_path) as pdf:
                        num_pages = len(pdf.pages)

                        for i in range(num_pages):
                            page = pdf.pages[i]
                            table = page.extract_tables()

                            if len(table) > 0:
                                for tab in table:
                                    df = pd.DataFrame(tab)
                                    dfs.append(df)
                    df = pd.concat(dfs, ignore_index=True)
                    new_df = self.extract_the_df(df)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop(columns=[0, 2, 3])
                    df = df.rename(columns={1: 'Value Date', 4: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    # df = df[df['Description'].notnull()]
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'IndusInd Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            def indus_format_3(unlocked_pdf_path):

                try:
                    unlocked_pdf_path = self.separate_lines_in_pdf(unlocked_pdf_path, timestamp)
                    print("Inside IndusInd Bank")

                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = self.extract_the_df(df_total)

                    # # start custom extraction
                    df = self.uncontinuous(w)
                    df = df.rename(
                        columns={0: 'Value Date', 1: 'TFDF', 2: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    for index, row in df.iterrows():
                        i = index
                        if df.at[i, 'TFDF'] != "":
                            df.at[i, 'new_column'] = 1
                    df.reset_index(drop=True, inplace=True)

                    for i in range(len(df) - 1):
                        if df.at[i, 'new_column'] == 1 and df.at[i + 1, 'Balance'] > 0:
                            df.at[i, 'Value Date'] = df.at[i + 1, 'Value Date']
                            df.at[i, 'Debit'] = df.at[i + 1, 'Debit']
                            df.at[i, 'Credit'] = df.at[i + 1, 'Credit']
                            df.at[i, 'Balance'] = df.at[i + 1, 'Balance']
                            df.at[i + 1, 'Value Date'] = ''
                            df.at[i + 1, 'Debit'] = ''
                            df.at[i + 1, 'Credit'] = ''
                            df.at[i + 1, 'Balance'] = ''
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df['new_column'] = np.nan
                    df.reset_index(drop=True, inplace=True)

                    counter = 0
                    for index, row in df.iterrows():
                        if pd.notnull(row["Value Date"]):
                            counter += 1
                        df.at[index, 'new_column'] = counter
                    for index, row in df.iterrows():
                        if pd.isna(row["Value Date"]):
                            df.at[index, 'new_column'] = np.NaN
                    df['new_column'].fillna(method='ffill', inplace=True)
                    df["Description"].fillna('', inplace=True)
                    df["Description"] = df.groupby('new_column')["Description"].transform(lambda x: ' '.join(x))
                    df = df.drop_duplicates(subset='new_column').reset_index(drop=True)
                    df = df.drop(['TFDF', 'new_column'], axis=1)
                    df = df.dropna(subset=['Description'])
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df = self.check_date(df)
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'IndusInd Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 3: {e}")
                    return None

            result = indus_format_1(unlocked_pdf_path)
            if result is None:
                result = indus_format_2(unlocked_pdf_path)
                if result is None:
                    result = indus_format_3(unlocked_pdf_path)
            logger.info("PDF to table Extraction for Indus bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def hdfc(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside HDFC Bank")

            def hdfc_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)

                    df_total = pd.DataFrame()
                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                        df_total = df_total.replace('', np.nan, regex=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    df = df.drop([2, 3]).reset_index(drop=True)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df.dropna(subset=["Value Date"])

                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'HDFC Bank'
                    logger.info("PDF to table Extraction for HDFC bank completed")

                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def hdfc_format_2(unlocked_pdf_path):
                try:
                    unlocked_pdf_path = self.separate_lines_in_pdf(unlocked_pdf_path, timestamp)
                    pdf = pdfplumber.open(unlocked_pdf_path)

                    df_total = pd.DataFrame()
                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                        df_total = df_total.replace('', np.nan, regex=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    df['new_column'] = np.nan
                    counter = 0
                    # Iterate over the dataframe rows
                    for index, row in df.iterrows():
                        if pd.notnull(row[0]):
                            counter += 1
                        df.at[index, 'new_column'] = counter
                    # Iterate over the dataframe rows
                    for index, row in df.iterrows():
                        if pd.isna(row[0]):
                            df.at[index, 'new_column'] = np.NaN
                    df['new_column'].fillna(method='ffill', inplace=True)
                    df[1].fillna('', inplace=True)
                    df[1] = df.groupby('new_column')[1].transform(lambda x: ' '.join(x))
                    df = df.drop_duplicates(subset='new_column').reset_index(drop=True)
                    df = df.drop([2, 3]).reset_index(drop=True)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df.dropna(subset=["Value Date"])

                    # Reorder the columns
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'HDFC Bank'
                    logger.info("PDF to table Extraction for HDFC bank completed")
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = hdfc_format_1(unlocked_pdf_path)
            if result is None:
                result = hdfc_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for HDFC bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def nkgsb(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside NKGSB Bank")

            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # start custom extraction
            df = df.iloc[:-1]  # Removes the first 1 rows
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
            # date
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
            idf['Bank'] = 'NKGSB Bank'
            logger.info("PDF to table Extraction for NKGSB bank completed")

            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def indian(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Indian Bank")

            def indian_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Indian Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    # df = df.iloc[1:] # Removes the first 1 rows
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 2: 'Debit', 3: 'Credit', 4: 'Balance'})
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df[pd.notna(df['Value Date'])]
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Indian Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def indian_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside Indian Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([1, 2, 4], axis=1)
                    df = df.rename(columns={0: 'Value Date', 3: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = df['Value Date'].astype(str).str.replace(' ', '')
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y',
                                                      errors='coerce').dt.strftime('%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Indian Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = indian_format_1(unlocked_pdf_path)
            if result is None:
                result = indian_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for Indian bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def tjsb(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside TJSB Bank")

            def tjsb_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside TJSB Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    # df = df.iloc[1:] # Removes the first 1 rows
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'TJSB Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def tjsb_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    print("Inside TJSB Bank")

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()

                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([2, 3], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)

                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'TJSB Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = tjsb_format_1(unlocked_pdf_path)
            if result is None:
                result = tjsb_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for TJSB bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def svc(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside SVC Bank")
            text = self.extract_dates_from_pdf(unlocked_pdf_path)
            date_pattern = r"\d{2} [A-Za-z]{3} \d{4}"
            dates = re.findall(date_pattern, text)
            s_date, e_date = dates[0], dates[1]
            complete_date_range = f"{s_date} to {e_date}"
            date_str = complete_date_range
            start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
            end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
            logger.info(f"Extracted Dates from pdf; Start Date: {start_date_str}, End Date: {end_date_str}")

            def extract_text_from_pdf(pdf_path, page_number=0):
                with pdfplumber.open(pdf_path) as pdf:
                    page = pdf.pages[page_number]
                    return page.extract_text()

            date_pattern = re.compile(r"\d{2}/[A-Za-z]{3}/\d{4}")
            number_pattern = re.compile(r"[\d,]+\.\d{2}")

            def further_refinement_parse_transaction_rows(extracted_text):
                rows = extracted_text.split('\n')
                parsed_data = []
                current_particulars = ""
                for row in rows:
                    date_match = date_pattern.search(row)
                    if date_match:
                        numerical_values = [match.group() for match in number_pattern.finditer(row)]
                        if numerical_values:
                            if current_particulars and parsed_data:
                                parsed_data[-1]["Particulars"] += " " + current_particulars.strip()
                                current_particulars = ""
                            particulars = row[date_match.end():row.find(numerical_values[0])].strip()
                            parsed_data.append({
                                "Date": date_match.group(),
                                "Particulars": particulars.rsplit(' ', 1)[0],
                                "Chq No": particulars.split()[-1] if len(particulars.split()) > 1 else "0.00",
                                "Debit": numerical_values[0],
                                "Credit": numerical_values[1] if len(numerical_values) > 1 else "0.00",
                                "Balance": numerical_values[2] if len(numerical_values) > 2 else "0.00"
                            })
                        else:
                            current_particulars += " " + row.strip()
                    else:
                        current_particulars += " " + row.strip()
                if current_particulars and parsed_data:
                    parsed_data[-1]["Particulars"] += " " + current_particulars.strip()
                return parsed_data

            complete_data = pd.DataFrame(columns=["Date", "Particulars", "Chq No", "Debit", "Credit", "Balance"])

            # Extracting data from all pages in the PDF and appending to the complete_data DataFrame
            with pdfplumber.open(unlocked_pdf_path) as pdf:
                for i in range(len(pdf.pages)):
                    page_text = extract_text_from_pdf(unlocked_pdf_path, page_number=i)
                    page_data = further_refinement_parse_transaction_rows(page_text)
                    complete_data = complete_data._append(pd.DataFrame(page_data), ignore_index=True)

            new_df = self.extract_the_df(complete_data)
            df = new_df.copy()
            df = df.drop("Chq No", axis=1)  # Removes column at position 1
            df = df.rename(columns={'Date': 'Value Date', 'Particulars': 'Description'})
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%b/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            # Reorder the columns
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            start_date_datetime = pd.to_datetime(start_date_str, format='%d-%m-%Y').strftime('%d-%m-%Y')
            end_date_datetime = pd.to_datetime(end_date_str, format='%d-%m-%Y').strftime('%d-%m-%Y')

            next_valid_index = idf['Balance'].iloc[1:].first_valid_index()
            if next_valid_index:
                credit_value = idf.loc[next_valid_index, 'Credit']
                debit_value = idf.loc[next_valid_index, 'Debit']
                balance_value = idf.loc[next_valid_index, 'Balance']

                # If it's a credit transaction, subtract the credit value from the balance
                if pd.notnull(credit_value) and credit_value > 0:
                    adjusted_balance = balance_value - credit_value
                # If it's a debit transaction, add the debit value to the balance
                elif pd.notnull(debit_value) and debit_value > 0:
                    adjusted_balance = balance_value + debit_value
                else:
                    adjusted_balance = balance_value
            else:
                adjusted_balance = None

            new_row = pd.DataFrame({
                'Value Date': [start_date_datetime],
                'Description': ["Opening Balance"],
                'Credit': 0,
                'Debit': 0,
                'Balance': [adjusted_balance]
            })
            last_balance = idf.iloc[-1]['Balance']

            new_last_row = pd.DataFrame({
                'Value Date': [end_date_datetime],
                'Description': ["Closing Balance"],
                'Credit': 0,
                'Debit': 0,
                'Balance': [last_balance]
            })
            idf = pd.concat([new_row, idf, new_last_row], ignore_index=True)
            idf['Bank'] = 'SVC Bank'
            logger.info("PDF to table Extraction for SVC bank completed")
            return idf


        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def deutsche(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Deutsche Bank")
            pdf = pdfplumber.open(unlocked_pdf_path)

            df_total = pd.DataFrame()
            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([2, 3], axis=1)
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            # # date
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            # df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Deutsche Bank'
            logger.info("PDF to table Extraction for Deutsche bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def indian_overseas(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside IOB")
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # start custom extraction
            df = df.drop([2, 3], axis=1)
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            df['Value Date'] = df['Value Date'].str.split(" ").str[0]
            # # date
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            # df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Indian Overseas Bank'
            logger.info("PDF to table Extraction for Indian Overseas bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def canara(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)

        try:

            def can_format_1(unlocked_pdf_path):
                try:
                    logger.info("Inside Canara Bank")

                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([0, 2, 4], axis=1)
                    df = df.rename(
                        columns={1: 'Value Date', 3: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    # # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y',
                                                      errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    # df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Canara Bank'
                    logger.info("PDF to table Extraction for Canara bank completed")
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def can_format_2(unlocked_pdf_path):
                try:
                    logger.info("Inside Canara Bank")

                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = df.drop([0, 2, 4], axis=1)
                    df = df.rename(
                        columns={1: 'Value Date', 3: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    # # date
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y',
                                                      errors='coerce').dt.strftime('%d-%m-%Y')
                    # df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Canara Bank'
                    logger.info("PDF to table Extraction for Canara bank completed")
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = can_format_1(unlocked_pdf_path)
            if result is None:
                result = can_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for Canara bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def boi(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside BOI Bank")

            def boi_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([1], axis=1)
                    df = df.rename(columns={0: 'Value Date', 2: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Bank of India'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def boi_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([0, 3], axis=1)
                    df = df.rename(columns={1: 'Value Date', 2: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Bank of India'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = boi_format_1(unlocked_pdf_path)
            if result is None:
                result = boi_format_2(unlocked_pdf_path)
            logger.info("PDF to table Extraction for BOI bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def dcb(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside DCB Bank")
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([2], axis=1)
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
            idf['Bank'] = 'DCB Bank'
            logger.info("PDF to table Extraction for DCB bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def fed(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Federal Bank")

            def fed_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # # start custom extraction
                    df = df.drop([0, 3, 4, 8], axis=1)
                    df = df.rename(columns={1: 'Value Date', 2: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Federal Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def fed_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # # start custom extraction
                    df = df.drop([1, 3, 4], axis=1)
                    df = df.rename(columns={0: 'Value Date', 2: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    # df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Federal Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            def fed_format_3(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # # start custom extraction
                    df = df.drop([0, 3, 4, 5, 9], axis=1)
                    df = df.rename(columns={1: 'Value Date', 2: 'Description', 6: 'Debit', 7: 'Credit', 8: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Federal Bank'

                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            result = fed_format_1(unlocked_pdf_path)
            if result is None:
                result = fed_format_3(unlocked_pdf_path)
                if result is None:
                    result = fed_format_2(unlocked_pdf_path)
            logger.info("PDF to table Extraction for Federal bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def cosmos(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Cosmos Bank")
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([2], axis=1)
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            # df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Cosmos Bank'
            logger.info("PDF to table Extraction for Cosmos bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def bom(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside BOM Bank")
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([1, 3, 7], axis=1)
            df = df.rename(columns={0: 'Value Date', 2: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Bank of Maharashtra'
            logger.info("PDF to table Extraction for BOM bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def tdcb(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([0, 2, 4], axis=1)
            df = df.rename(columns={1: 'Value Date', 3: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%Y-%m-%d', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            # df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df[:-1]
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
            idf['Bank'] = 'Thane District Central Bank'
            logger.info("PDF to table Extraction for Thane District Central bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def rbl(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside RBL Bank")
            x_positions = [110, 420, 505, 605, 720, 808]
            unlocked_pdf_path = self.separate_lines_in_vertical_pdf(unlocked_pdf_path, x_positions, timestamp)
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([0, 3, 4], axis=1)
            df = df.rename(columns={1: 'Value Date', 2: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'RBL Bank'
            logger.info("PDF to table Extraction for RBL bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def karnataka(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Karnataka Bank")

            x_positions = [108, 420, 545, 659]
            unlocked_pdf_path = self.separate_lines_in_vertical_pdf(unlocked_pdf_path, x_positions, timestamp)
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            # df = df.drop([0, 3, 4], axis=1)
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 2: 'Debit', 3: 'Credit', 4: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Karnataka Bank'
            logger.info("PDF to table Extraction for Karnataka bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def hsbc(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([5], axis=1)
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 2: 'Debit', 3: 'Credit', 4: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = df.reset_index(drop=True)
            for index, row in df.iterrows():
                if pd.isna(row['Value Date']) and index != 0:
                    df.at[index - 1, 'Description'] = df.at[index, 'Description']
            df = df.dropna(subset=['Value Date'])
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'HSBC Bank'

            logger.info("PDF to table Extraction for HSBC bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def bccb(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Bassein Catholic Co-op Bank")

            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([1, 2], axis=1)
            df = df.rename(columns={0: 'Value Date', 3: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df = df.iloc[1:]
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Bassein Catholic Co-op Bank'
            logger.info("PDF to table Extraction for Bassein Catholic Co-op bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def mcb(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Municipal Co-op Bank")
            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([0, 3], axis=1)
            df = df.rename(columns={1: 'Value Date', 2: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Municipal Co-op Bank'
            logger.info("PDF to table Extraction for Municipal Co-op bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def bharat(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Bharat Bank")

            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([0, 3, 7], axis=1)
            df = df.rename(columns={1: 'Value Date', 2: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Bharat Co-op Bank'
            logger.info("PDF to table Extraction for Bharat bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def scb(self, unlocked_pdf_path, timestamp, start_date):

        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside SCB Bank")

            unlocked_pdf_path = self.separate_lines_in_pdf_scb(unlocked_pdf_path, timestamp)

            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            df = w.copy()
            # # start custom extraction
            df = df.drop([0, 3], axis=1)
            df = df.rename(columns={1: 'Value Date', 2: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']

            df["Value Date"].replace('', pd.NA, inplace=True)
            df["Value Date"].fillna(method='ffill', inplace=True)
            df.loc[df['Balance'] == '', 'Value Date'] = pd.NA
            df['new_column'] = np.nan
            counter = 0
            for index, row in df.iterrows():
                if pd.notnull(row["Value Date"]):
                    counter += 1
                df.at[index, 'new_column'] = counter
            for index, row in df.iterrows():
                if pd.isna(row["Value Date"]):
                    df.at[index, 'new_column'] = np.NaN
            df['new_column'].fillna(method='ffill', inplace=True)
            df["Description"].fillna('', inplace=True)
            df["Description"] = df.groupby('new_column')["Description"].transform(lambda x: ' '.join(x))
            df = df.drop_duplicates(subset='new_column').reset_index(drop=True)
            df = df.drop(["new_column"], axis=1)

            year = start_date[-4:]
            df['Value Date'] = df['Value Date'] + f' {year}'
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%b %d %Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')

            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Balance'] = pd.to_numeric(df['Balance'].str.replace('-', ''))
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[~((df['Debit'] > 1) & (df['Credit'] > 1) | (pd.isna(df['Debit']) & pd.isna(df['Credit'])))]
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])

            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
            idf['Bank'] = 'SCB Bank'
            logger.info("PDF to table Extraction for SCB bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def uco(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside UCO Bank")

            def uco_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([2], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'UCO Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def uco_format_2(unlocked_pdf_path):
                try:
                    unlocked_pdf_path = self.add_lines_to_pdf(unlocked_pdf_path, unlocked_pdf_path)
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([2], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['new'] = 0
                    df.reset_index(drop=True, inplace=True)

                    for index, row in df.iterrows():
                        value_date_parts = str(row['Value Date']).split(" ")
                        bal_parts = str(row['Balance']).split(" ")
                        cred_parts = str(row['Credit']).split(" ")
                        deb_parts = str(row['Debit']).split(" ")
                        if len(value_date_parts) == 2:
                            new_row = row.copy()
                            df.at[index, 'Value Date'] = value_date_parts[0]
                            df.at[index, 'Balance'] = bal_parts[0]
                            new_row['Value Date'] = value_date_parts[1]
                            new_row['Balance'] = bal_parts[1]
                            new_row['new'] = 1
                            if len(cred_parts) == 2:
                                df.at[index, 'Credit'] = cred_parts[0]
                                new_row['Credit'] = cred_parts[1]
                            if len(deb_parts) == 2:
                                df.at[index, 'Debit'] = deb_parts[0]
                                new_row['Debit'] = deb_parts[1]
                            df.loc[index + 1] = new_row
                            df.reset_index(drop=True, inplace=True)

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    df.reset_index(drop=True, inplace=True)
                    for index, row in df.iterrows():
                        deb = row['Debit']
                        cred = row['Credit']
                        curr_bal = row['Balance']
                        prev_bal = df.iloc[index - 1]['Balance']
                        if deb > 1 and cred > 1:
                            if int(curr_bal) + int(deb) == int(prev_bal):
                                df.at[index, 'Credit'] = pd.NA
                            elif int(curr_bal) - int(cred) == int(prev_bal):
                                df.at[index, 'Debit'] = pd.NA
                    df.reset_index(drop=True, inplace=True)
                    for index in df[df['new'] == 1].index:
                        hey = df.at[index, 'Description'].split(" ", 1)
                        df.at[index, 'Description'] = hey[1]
                        df.at[index - 1, 'Description'] = hey[0]
                    df.reset_index(drop=True, inplace=True)

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'UCO Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = uco_format_1(unlocked_pdf_path)
            if result is None:
                result = uco_format_2(unlocked_pdf_path)
            logger.info("PDF to table Extraction for UCO bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def vasai(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Vasai Bank")

            def vasai_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([2], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Vasai Vikas Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def vasai_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([2, 3, 7], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])
                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
                    idf['Bank'] = 'Vasai Vikas Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            result = vasai_format_1(unlocked_pdf_path)
            if result is None:
                result = vasai_format_2(unlocked_pdf_path)

            logger.info("PDF to table Extraction for Vasai bank completed")
            return result

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def saraswat(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Saraswat Bank")

            def saraswat_format_1(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    df = w.copy()
                    # new_df = self.extract_the_df(w)
                    # df = self.uncontinuous(new_df)
                    # start custom extraction
                    df = df.drop([2], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Description'])
                    df = df[df['Description'] != '']
                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Saraswat Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 1: {e}")
                    return None

            def saraswat_format_2(unlocked_pdf_path):
                try:
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    new_df = self.extract_the_df(w)
                    df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = df.iloc[2:]
                    df = df.rename(columns={0: 'Transaction'})
                    df['Description'] = ''
                    df.reset_index(drop=True, inplace=True)
                    for index, row in df.iterrows():
                        i = index
                        if "Cr" in df.at[i, 'Transaction'] or "Dr" in df.at[i, 'Transaction']:
                            df.at[i, 'Description'] = df.at[i, 'Transaction']
                            df.at[i, 'Transaction'] = ''
                    adf = df["Transaction"]
                    adf.replace('', np.nan, inplace=True)
                    adf.dropna(inplace=True)
                    bdf = df["Description"]
                    bdf.replace('', np.nan, inplace=True)
                    bdf.dropna(inplace=True)
                    adf = adf.to_list()
                    df = pd.DataFrame({'Transaction': bdf, 'Description': adf})
                    df['Value Date'] = ''
                    df['Balance'] = ''
                    df['CR/DR'] = ''
                    df['Amount'] = ''
                    for index, row in df.iterrows():
                        a_list = df.at[index, 'Transaction'].split(" ")
                        df.at[index, 'Value Date'] = a_list[0]
                        df.at[index, 'CR/DR'] = a_list[1]
                        df.at[index, 'Amount'] = a_list[2]
                        df.at[index, 'Balance'] = a_list[3]
                    df.reset_index(drop=True, inplace=True)
                    df['Credit'] = 0
                    df['Debit'] = 0
                    df.loc[df['CR/DR'] == 'Cr', 'Credit'] = df.loc[df['CR/DR'] == 'Cr', 'Amount']
                    df.loc[df['CR/DR'] != 'Cr', 'Debit'] = df.loc[df['CR/DR'] != 'Cr', 'Amount']
                    df = df.drop(['CR/DR', 'Amount', 'Transaction'], axis=1)

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Saraswat Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 2: {e}")
                    return None

            def saraswat_format_3(unlocked_pdf_path):
                try:
                    unlocked_pdf_path = self.add_lines_to_pdf(unlocked_pdf_path, unlocked_pdf_path)
                    pdf = pdfplumber.open(unlocked_pdf_path)
                    df_total = pd.DataFrame()

                    for i in range(len(pdf.pages)):
                        p0 = pdf.pages[i]
                        table = p0.extract_table()
                        df_total = df_total._append(table, ignore_index=True)
                        df_total.replace({r'\n': ' '}, regex=True, inplace=True)
                    w = df_total.drop_duplicates()
                    # df = self.extract_the_df(w)
                    # df = self.uncontinuous(new_df)

                    # start custom extraction
                    df = w.drop([2, 6], axis=1)
                    df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
                    df = df.dropna(subset=['Description'])

                    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y %H:%M:%S',
                                                      errors='coerce').dt.strftime(
                        '%d-%m-%Y')
                    df = self.check_date(df)
                    df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
                    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
                    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
                    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
                    df = df[df['Description'] != '']
                    df = df.dropna(subset=['Value Date'])

                    idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

                    idf['Bank'] = 'Saraswat Bank'
                    return idf

                except Exception as e:
                    print(f"Error in format 3: {e}")
                    return None

            result = saraswat_format_1(unlocked_pdf_path)
            if result is None:
                result = saraswat_format_2(unlocked_pdf_path)
                if result is None:
                    result = saraswat_format_3(unlocked_pdf_path)
            logger.info("PDF to table Extraction for Saraswat bank completed")
            return result


        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def surat(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Surat Bank")

            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([2, 3], axis=1)
            df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%Y-%m-%d', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.split().str[0]
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Surat Co-op Bank'
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def janakal(self, unlocked_pdf_path, timestamp):
        logger = logging.getLogger(self.CA_ID)
        try:
            logger.info("Inside Janakalyan Bank")

            pdf = pdfplumber.open(unlocked_pdf_path)
            df_total = pd.DataFrame()

            for i in range(len(pdf.pages)):
                p0 = pdf.pages[i]
                table = p0.extract_table()
                df_total = df_total._append(table, ignore_index=True)
                df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            w = df_total.drop_duplicates()
            new_df = self.extract_the_df(w)
            df = self.uncontinuous(new_df)
            # # start custom extraction
            df = df.drop([0], axis=1)
            df = df.rename(columns={1: 'Value Date', 2: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
            df = df.dropna(subset=['Description'])
            df = df[df['Description'] != '']
            df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
                '%d-%m-%Y')
            df = self.check_date(df)
            df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
            df = df[df['Description'] != '']
            df = df.dropna(subset=['Value Date'])
            idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

            idf['Bank'] = 'Janakalyan Bank'
            logger.info("PDF to table Extraction for Janakalyan bank completed")
            return idf

        except Exception as e:
            logger.error(f"An error occurred while processing pdf: {str(e)}", exc_info=True)

    def custom_extraction(self, bank, pdf_path, pdf_password, timestamp):
        bank = re.sub(r'\d+', '', bank)
        unlocked_pdf_path = pdf_path
        #unlocked_pdf_path = self.unlock_the_pdfs_path(pdf_path, pdf_password, bank, timestamp)

        if bank == "AXIS":
            df = pd.DataFrame(self.axis(unlocked_pdf_path, timestamp))

        elif bank == "IDBI":
            df = pd.DataFrame(self.idbi(unlocked_pdf_path, timestamp))

        elif bank == "SBI":
            df = pd.DataFrame(self.sbi(unlocked_pdf_path, timestamp))

        elif bank == "JANAKAL":
            df = pd.DataFrame(self.janakal(unlocked_pdf_path, timestamp))

        elif bank == "IDFC":
            df = pd.DataFrame(self.idfc(unlocked_pdf_path, timestamp))

        elif bank == "PNB":
            df = pd.DataFrame(self.pnb(unlocked_pdf_path, timestamp))

        elif bank == "YES":
            df = pd.DataFrame(self.yes_bank(unlocked_pdf_path, timestamp))

        elif bank == "KOTAK":
            df = pd.DataFrame(self.kotak(unlocked_pdf_path, timestamp))

        elif bank == "UNION":
            df = pd.DataFrame(self.union(unlocked_pdf_path, timestamp))

        elif bank == "ICICI":
            df = pd.DataFrame(self.icici(unlocked_pdf_path, timestamp))

        elif bank == "BOB":
            df = pd.DataFrame(self.bob(unlocked_pdf_path, timestamp))

        elif bank == "INDUSIND":
            df = pd.DataFrame(self.indus(unlocked_pdf_path, timestamp))

        elif bank == "INDIAN":
            df = pd.DataFrame(self.indian(unlocked_pdf_path, timestamp))

        elif bank == "TJSB":
            df = pd.DataFrame(self.tjsb(unlocked_pdf_path, timestamp))

        elif bank == "NKGSB":
            df = pd.DataFrame(self.nkgsb(unlocked_pdf_path, timestamp))

        elif bank == "HDFC":
            df = pd.DataFrame(self.hdfc(unlocked_pdf_path, timestamp))

        elif bank == "SVC":
            df = pd.DataFrame(self.svc(unlocked_pdf_path, timestamp))

        elif bank == "DEUTSCHE":
            df = pd.DataFrame(self.deutsche(unlocked_pdf_path, timestamp))

        elif bank == "IOB":
            df = pd.DataFrame(self.indian_overseas(unlocked_pdf_path, timestamp))

        elif bank == "CANARA":
            df = pd.DataFrame(self.canara(unlocked_pdf_path, timestamp))

        elif bank == "BOI":
            df = pd.DataFrame(self.boi(unlocked_pdf_path, timestamp))

        elif bank == "BOM":
            df = pd.DataFrame(self.bom(unlocked_pdf_path, timestamp))

        elif bank == "COSMOS":
            df = pd.DataFrame(self.cosmos(unlocked_pdf_path, timestamp))

        elif bank == "DCB":
            df = pd.DataFrame(self.dcb(unlocked_pdf_path, timestamp))

        elif bank == "FED":
            df = pd.DataFrame(self.fed(unlocked_pdf_path, timestamp))

        elif bank == "TDCB":
            df = pd.DataFrame(self.tdcb(unlocked_pdf_path, timestamp))

        elif bank == "RBL":
            df = pd.DataFrame(self.rbl(unlocked_pdf_path, timestamp))

        elif bank == "KARNATAKA":
            df = pd.DataFrame(self.karnataka(unlocked_pdf_path, timestamp))

        elif bank == "HSBC":
            df = pd.DataFrame(self.hsbc(unlocked_pdf_path, timestamp))

        elif bank == "BCCB":
            df = pd.DataFrame(self.bccb(unlocked_pdf_path, timestamp))

        elif bank == "MCB":
            df = pd.DataFrame(self.mcb(unlocked_pdf_path, timestamp))

        elif bank == "BHARAT":
            df = pd.DataFrame(self.bharat(unlocked_pdf_path, timestamp))

        elif bank == "SCB":
            df = pd.DataFrame(self.scb(unlocked_pdf_path, timestamp, "01-01-01"))

        elif bank == "UCO":
            df = pd.DataFrame(self.uco(unlocked_pdf_path, timestamp))

        elif bank == "VASAI":
            df = pd.DataFrame(self.vasai(unlocked_pdf_path, timestamp))

        elif bank == "SARASWAT":
            df = pd.DataFrame(self.saraswat(unlocked_pdf_path, timestamp))

        elif bank == "SURAT":
            df = pd.DataFrame(self.surat(unlocked_pdf_path, timestamp))

        else:
            raise ExtractionError("Bank Does not Exist")

        return df
