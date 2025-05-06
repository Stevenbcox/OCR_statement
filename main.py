import os
import re
import pyodbc
import shelve
import fitz
import datetime
import pandas as pd


def get_connection_string():
    with shelve.open('P:/Users/Steven Cox/sql_creds/credentials') as db:
        server = db['server']
        database = db['database']
        username = db['username']
        password = db['password']
    print('✅ Credentials loaded')

    return f"DRIVER=ODBC Driver 18 for SQL Server; SERVER={server}; DATABASE={database}; ENCRYPT=no; UID={username}; PWD={password}"


def extract_text_from_pdf(pdf_path):
    text = ''
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text += page.get_text()
        doc.close()
    except Exception as e:
        print(f"❌ Error extracting text from {pdf_path}: {e}")
    return text


def find_refno(pdf_filename):
    match = re.search(r'\d{9}', pdf_filename)
    return match.group(0) if match else None


def find_parameters(text):
    new_balance_match = re.search(r'(New Balance:|New balance)\b.*?(\$[\d,]+\.\d{2})', text, re.DOTALL | re.IGNORECASE)
    closing_date_match = re.search(r'Statement Closing Date[^\n]*\n(.+)', text, re.IGNORECASE)
    due_date_match = re.search(r'(Payment Due Date:|Payment due date)[^\n]*\n(.+)', text, re.IGNORECASE)

    new_balance = new_balance_match.group(2) if new_balance_match else None
    closing_date = closing_date_match.group(1).strip() if closing_date_match else None
    due_date = due_date_match.group(2).strip() if due_date_match else None

    return new_balance, closing_date, due_date


def execute_sql_query(search_refno, connection_string):
    try:
        connection = pyodbc.connect(connection_string)
        cursor = connection.cursor()
        query = """
            SELECT m.FILENO, m.FORW_FILENO, m.FORW_REFNO, d.NAME, m.CHARGE_OFF, m.ORIG_CLAIM, m.CHARGE_OFF_DATE
            FROM MASTER m INNER JOIN DEBTOR d ON m.FILENO = d.FILENO
            WHERE m.FORW_REFNO = ?
        """
        cursor.execute(query, (search_refno,))
        data = cursor.fetchall()
    except Exception as e:
        print(f"❌ Error executing SQL query: {e}")
        data = []
    finally:
        connection.close()
    return data


def write_to_excel(data, output_path, new_balance, closing_date, due_date):
    try:
        if os.path.exists(output_path):
            df = pd.read_excel(output_path, dtype=str)  # Read with all columns as strings
        else:
            df = pd.DataFrame(columns=[
                'MASTER.FILENO', 'MASTER.FORW_FILENO', 'MASTER.FORW_REFNO', 'DEBTOR.NAME',
                'MASTER.CHARGE_OFF', 'MASTER.ORIG_CLAIM', 'MASTER.CHARGE_OFF_DATE',
                'ACCOUNT NUMBER', 'NEW BALANCE', 'STATEMENT CLOSING DATE', 'DUE DATE'
            ])

        for row in data:
            # Convert FILENO to a string and prepend an apostrophe to force Excel to treat it as text
            fileno = f"'{row[0]}"  # Ensures Excel sees it as a text field
            filled_row = [fileno] + list(row[1:]) + [row[1], new_balance, closing_date, due_date]

            new_row_df = pd.DataFrame([filled_row], columns=df.columns)
            df = pd.concat([df, new_row_df], ignore_index=True)

        # Save the DataFrame to Excel without altering text formatting
        df.to_excel(output_path, index=False, engine='openpyxl')

        print(f"✅ Excel file updated: {output_path}")

    except Exception as e:
        print(f"❌ Error writing to Excel: {e}")


def main(input_path, output_folder):
    if not os.path.exists(input_path):
        print(f"❌ Input path does not exist: {input_path}")
        return

    # If the input is a single file, process it directly
    if os.path.isfile(input_path):
        if not input_path.lower().endswith('.pdf'):
            print(f"❌ Input file is not a PDF: {input_path}")
            return

        print(f"📄 Processing single PDF: {input_path}")
        process_pdf(input_path, output_folder)

    # If the input is a folder, process all PDFs inside
    elif os.path.isdir(input_path):
        print(f"📂 Input folder found: {input_path}")
        pdf_files_found = False

        for dirpath, dirnames, filenames in os.walk(input_path):
            for filename in filenames:
                if filename.lower().endswith('.pdf'):
                    pdf_files_found = True
                    pdf_path = os.path.join(dirpath, filename)
                    print(f"📄 Processing PDF: {pdf_path}")
                    process_pdf(pdf_path, output_folder)

        if not pdf_files_found:
            print(f"❌ No PDF files found in {input_path}")

    try:
        os.startfile(output_folder)
        print(f"📂 Output folder opened: {output_folder}")
    except Exception as e:
        print(f"❌ Failed to open output folder: {e}")


def process_pdf(pdf_path, output_folder):
    try:
        pdf_text = extract_text_from_pdf(pdf_path)
        search_refno = find_refno(os.path.basename(pdf_path))
        print(f"🔍 Query REFNO: {search_refno}")

        new_balance, closing_date, due_date = find_parameters(pdf_text)
        print(f"📑 Extracted parameters: New Balance={new_balance}, Closing Date={closing_date}, Due Date={due_date}")

        if new_balance and closing_date and due_date:
            connection_string = get_connection_string()
            query_result = execute_sql_query(search_refno, connection_string)
            current_date = datetime.datetime.now().strftime('%m-%d-%Y')
            output_excel = os.path.join(output_folder, f"ocr_statement_output_{current_date}.xlsx")
            write_to_excel(query_result, output_excel, new_balance, closing_date, due_date)

            if os.path.exists(output_excel):
                print(f"✅ Excel file saved: {output_excel}")
            else:
                print(f"❌ Excel file not found in output folder!")

    except Exception as e:
        print(f"❌ Error processing {pdf_path}: {e}")


if __name__ == "__main__":
    input_path = r''
    output_folder = r''

    main(input_path, output_folder)


def run_from_gui(input_folder, output_folder):
    """ Function to be called from the GUI to process the selected folders """
    if input_folder and output_folder:
        main(input_folder, output_folder)
