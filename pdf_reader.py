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

def text_parser(text, filename):
    # Build list with each text line
    comp_list = text.split('\n')
    # Remove empty elements from list
    comp_list = list(filter(None, comp_list))
    # Get invoice number and date
    inv_nb = comp_list[comp_list.index("INVOICE #") + 2]
    date = comp_list[comp_list.index("DATE") + 2]
    print("inv. number " + inv_nb + " date " + date)
    # Get start and end index to get items data
    idx = comp_list.index("ITEMS")
    end_idx = comp_list.index("NOTES:")
    row = 0
    wb, ws = open_excel(filename)
    while idx < end_idx:
        name   = comp_list[idx]
        desc   = comp_list[idx + 1]
        qty    = comp_list[idx + 2]
        price  = comp_list[idx + 3]
        tax    = comp_list[idx + 4]
        amount = comp_list[idx + 5]
        item_list = []
        item_list.extend([name, desc, qty, price, tax, amount])
        write_to_excel(ws, item_list, row)
        print("name " + name + " desc " + desc + " qty " + qty + " price " + price + " tax " + tax + " amount " + amount)
        idx += 6
        row += 1
    close_excel(wb)

def open_excel(filename):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    return workbook, worksheet

def close_excel(workbook):
    workbook.close()

def write_to_excel(worksheet, item_list, row):
    column = 0
    for elem in item_list:
        worksheet.write(row, column, elem)
        column += 1

# Defining main function
def main():
    # Retrieve the text from the 1st PDF page
    text = get_text_from_pdf('Invoice.pdf')
    # Parsing the data
    text_parser(text, 'InvoiceData.xlsx')

if __name__=="__main__":
    main()

