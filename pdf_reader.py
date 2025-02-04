# Importing required modules
import pdftotext
import xlsxwriter

# Retrieve text from PDF
def get_text_from_pdf(filename):
    # Open pdf file for reading operation
    with open(filename, "rb") as file_handle:
        # Convert pdf to text
        pdf = pdftotext.PDF(file_handle)
        # Get pdf 1st page (assuming there is only one)
        return pdf[0]

def text_parser(text):
    # Build list with each text line
    comp_list = text.split('\n')
    # Remove empty elements from list
    comp_list = list(filter(None, comp_list))
    # Get invoice number and date headers indexes
    inv_nb_hdr_idx = comp_list.index("INVOICE #")
    date_hdr_idx = comp_list.index("DATE")
    # Get start and end index to get data
    idx = comp_list.index("ITEMS")
    end_idx = comp_list.index("NOTES:")
    # Initialize empty list of items lists
    items_list = []
    # First iteration of the loop woll be the header titles
    is_header = True
    while idx < end_idx:
        if is_header == True:
            inv = comp_list[inv_nb_hdr_idx]
            date = comp_list[date_hdr_idx]
            is_header = False
        else:
            inv = comp_list[inv_nb_hdr_idx + 2]
            date = comp_list[date_hdr_idx + 2]
        name   = comp_list[idx]
        desc   = comp_list[idx + 1]
        qty    = comp_list[idx + 2]
        price  = comp_list[idx + 3]
        tax    = comp_list[idx + 4]
        amount = comp_list[idx + 5]
        items_list.append([inv, date, name, desc, qty, price, tax, amount])
        idx += 6
    return items_list

def export_data(excel_filename, items_list):
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet = workbook.add_worksheet()
    row = 0
    for list in items_list:
        column = 0
        for elem in list:
            worksheet.write(row, column, elem)
            column += 1
        row += 1

# Defining main function
def main():
    # Retrieve the text from the 1st PDF page
    text = get_text_from_pdf('Invoice.pdf')
    # Parsing the data
    items = text_parser(text)
    # Export data to excel
    export_data('InvoiceData.xlsx', items)

if __name__=="__main__":
    main()

