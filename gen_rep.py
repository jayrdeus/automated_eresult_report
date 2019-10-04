import xlsxwriter as writer
import db
import datetime as dt
class Report:
    def __init__(self,data,filename):
        self.data = data
        self.filename = filename
    def generate_report(self):
        print(f'filname : {self.filename}')
        row = 5
        col = 0
        # create Work book
        workbook = writer.Workbook(f'reports/{self.filename}.xlsx')
        worksheet = workbook.add_worksheet()

        # transfer data from workbook
        column_header = 0
        row_header = 4

        # Title
        worksheet.set_column('B:G', 12)
        worksheet.set_row(3, 30)
        worksheet.set_row(6, 30)
        worksheet.set_row(7, 30)
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'size': 20,
            'fg_color': 'gray'})

        worksheet.merge_range('A2:I3',f'e-Result Report',merge_format)

        # Headers
        headers = ['Patient Name','Procedure','Mobile Number','Date Registered','Result date','Variance','Date SMS Sent','Date Downloaded','Downloaded']
        format_header = workbook.add_format({
            'align' : 'center',
            'border': 1,
            'bold' : 1,
            'size': 14,
            'fg_color': 'yellow'
        })
        for header in headers:
            worksheet.write(row_header,column_header,header,format_header)
            column_header += 1
        cell_format = workbook.add_format({
            'border':1,
            'align':'center'
        })
        # Body
        for number in self.data:
            worksheet.write(row,col,number[0],cell_format) # Patient Name
            worksheet.write(row,col + 1,number[1],cell_format) # Procedure
            worksheet.write(row,col + 2,number[2],cell_format) # Mobile number          
            worksheet.write(row,col + 3,number[3],cell_format) # Date Registered
            worksheet.write(row,col + 4,number[4],cell_format) # Result Date
            worksheet.write(row,col + 5,number[5],cell_format) # Variance
            worksheet.write(row,col + 6,number[6],cell_format) # Date SMS Sent
            worksheet.write(row,col + 7,number[7],cell_format) # Date Downloaded
            worksheet.write(row,col + 8,number[8],cell_format) # Downloaded
            row +=1

        
        worksheet.set_column("A:A",40)
        worksheet.set_column("B:B",56)
        worksheet.set_column("C:C",19)
        worksheet.set_column("D:D",22)
        worksheet.set_column("E:E",18)
        worksheet.set_column("F:F",12)
        worksheet.set_column("G:G",21)
        worksheet.set_column("H:H",21)
        worksheet.set_column("I:I",17)
        

        workbook.close()
        print('Successfully Created')
