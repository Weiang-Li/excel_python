import win32com.client
from win32com.client import constants
import fnmatch

ExcelApp = win32com.client.Dispatch("Excel.Application")
ExcelApp.Visible = True

column_index = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J',
                11: 'K', 12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T',
                21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z', 27: 'AA', 28: 'AB', 29: 'AC', 30: 'AD'
                }


class Vba:
    '''
    Note:
        if constants not working:
            import win32com
            print(win32com.__gen_path__)
            go to the folder, and delete the folder with long id
    '''

    def __init__(self):
        pass

    def open_file(self, filename_with_path):
        ExcelApp.Workbooks.Open(filename_with_path)

    def create_new_workbook(self):

        ExcelWorkbook = ExcelApp.Workbooks.Add()

        return ExcelWorkbook.Name

    def create_newsheet(self, file_name, name_sheet=None):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        Excelsheet = ExcelWorkbook.Worksheets.Add()
        if name_sheet:
            ExcelWorkbook.ActiveSheet.Name = name_sheet
        return Excelsheet.Name

    def rename_sheet(self, file_name, current_sheetname, rename_sheetname_to):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelWorkbook.Worksheets(current_sheetname).Name = rename_sheetname_to

    def insert_column(self, file_name, sheetnumber_or_sheetname, column_index, insert_how_many_column):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        for i in range(1, insert_how_many_column + 1):
            ExcelSheet.Columns(column_index).Insert()

    def insert_row(self, file_name, sheetnumber_or_sheetname, row_index, insert_how_many_row):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        for i in range(1, insert_how_many_row + 1):
            ExcelSheet.Rows(row_index).Insert()

    def get_last_row(self, file_name, sheetnumber_or_sheetname):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        lastrow = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).UsedRange.Rows.Count
        return lastrow

    def row_AutoFit(self, file_name, sheetnumber_or_sheetname, row_number=None):
        Excelworkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = Excelworkbook.Worksheets(sheetnumber_or_sheetname)
        if row_number:
            ExcelSheet.Rows(row_number).AutoFit()
        else:
            ExcelSheet.Rows.AutoFit()

    def set_rowwidth(self, file_name, sheetnumber_or_sheetname, row_number, height):
        Excelworkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = Excelworkbook.Worksheets(sheetnumber_or_sheetname)
        ExcelSheet.Rows(row_number).Rowheight = height

    def get_last_column(self, file_name, sheetnumber_or_sheetname):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        lastcolumn = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).UsedRange.Columns.Count
        return lastcolumn

    def set_columnwidth(self, file_name, sheetnumber_or_sheetname, column_number, width):
        Excelworkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = Excelworkbook.Worksheets(sheetnumber_or_sheetname)
        ExcelSheet.Columns(column_number).Columnwidth = width

    def column_AutoFit(self, file_name, sheetnumber_or_sheetname, column_number=None):
        Excelworkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = Excelworkbook.Worksheets(sheetnumber_or_sheetname)
        if column_number:
            ExcelSheet.Columns(column_number).AutoFit()
        else:
            ExcelSheet.Columns.AutoFit()

    def create_FormulaR1C1_for_range(self, file_name, sheetnumber_or_sheetname,
                                     range_of_formula_you_want_to_put, FormulaR1C1):

        Excelworkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = Excelworkbook.Worksheets(sheetnumber_or_sheetname)
        excelrange = ExcelSheet.Range(range_of_formula_you_want_to_put)
        excelrange.FormulaR1C1 = FormulaR1C1

    def insert_column_and_get_name_initial_and_autofill_to_lastrow(self, file_name, sheetnumber_or_sheetname,
                                                                   searchcriteria_in_header, name_new_column):

        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        lastcolumn = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).UsedRange.Columns.Count
        lastrow = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).UsedRange.Rows.Count
        for i in range(1, lastcolumn + 1):
            if fnmatch.fnmatch(ExcelSheet.Cells(1, i).Value, '*' + searchcriteria_in_header + '*'):
                new_column = ExcelSheet.Cells(1, i).Offset(2, 2).Column
                ExcelSheet.Columns(i + 1).Insert()
                ExcelSheet.Cells(1, i).Offset(1, 2).Value = name_new_column
                ExcelSheet.Cells(1, i).Offset(2,
                                              2).FormulaR1C1 = '=concatenate(left(R[0]C[-2],1),".",left(R[0]C[-1],1))'

                try:
                    ExcelSheet.Cells(1, i).Offset(2, 2).AutoFill(ExcelSheet.Range(
                        column_index.get(i + 1) + '2:' + column_index.get(i + 1) + "" + str(lastrow) + ""),
                        constants.xlFillDefault)
                    print('true')
                except:
                    ExcelSheet.Range(column_index.get(i + 1) + '2:' + column_index.get(i + 1) + "" + str(
                        lastrow) + "").Value = '=concatenate(left(R[0]C[-2],1),".",left(R[0]C[-1],1))'
                    print('second option')
                break

    def autofill(self, file_name, sheetnumber_or_sheetname, cell_reference):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        lastrow = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).UsedRange.Rows.Count
        if len(cell_reference) > 4:
            first_cell_column = str(cell_reference)[0]
            first_cell_row = str(cell_reference)[1]
            second_cell_column = str(cell_reference)[3]
            second_cell_row = str(cell_reference)[4]
            for i in range(int(first_cell_row), lastrow, 1):
                try:
                    ExcelSheet.Range(cell_reference).AutoFill(ExcelSheet.Range(
                        first_cell_column + first_cell_row + ':' + second_cell_column + str(lastrow)),
                        constants.xlFillDefault)
                except:
                    ExcelSheet.Range(first_cell_column + first_cell_row + ":" + second_cell_column + str(
                        lastrow)).Value = ExcelSheet.Range(cell_reference).Value
        else:
            cell_column = str(cell_reference)[0]
            cell_row = str(cell_reference)[1]
            ExcelSheet.Range(cell_reference).Select()
            for i in range(int(cell_row), lastrow, 1):
                try:
                    ExcelSheet.Range(cell_reference).AutoFill(ExcelSheet.Range(
                        cell_column + cell_row + ':' + cell_column + str(lastrow)),
                        constants.xlFillDefault)
                except:
                    ExcelSheet.Range(
                        cell_column + cell_row + ":" + cell_column + str(lastrow)).Value = ExcelSheet.Range(
                        cell_reference).Value

    def filter(self, file_name, sheetnumber_or_sheetname, filter_column, filter_criteria):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        lastcolumn = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).UsedRange.Columns.Count
        ExcelSheet.Columns('A:' + column_index.get(lastcolumn)).AutoFilter(Field=filter_column,
                                                                           Criteria1=filter_criteria)

    def enter_value(self, file_name, sheetnumber_or_sheetname, Range, value):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        range_selected = ExcelSheet.Range(Range)
        range_selected.Value = value

    def loop_each_cell_in_range(self, file_name, sheetnumber_or_sheetname, Range, search_criteria):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        range_selected = ExcelSheet.Range(Range)
        for eachcell in range_selected:
            if eachcell.Value == search_criteria:
                print('True')
            else:
                print('False')

    def set_Bold(self, file_name, sheetnumber_or_sheetname, Range, Bold_True_or_False=None):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        range_selected = ExcelSheet.Range(Range)
        range_selected.Font.Bold = Bold_True_or_False

    def set_font(self, file_name, sheetnumber_or_sheetname, Range, Font=None):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        range_selected = ExcelSheet.Range(Range)
        range_selected.Font.Name = Font

    def set_font_color(self, file_name, sheetnumber_or_sheetname, Range, font_color_index=None):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        range_selected = ExcelSheet.Range(Range)
        try:
            range_selected.Font.ColorIndex = font_color_index
        except:
            print('Error, Colorindex does not exists')
            exit()

    def set_background_color(self, file_name, sheetnumber_or_sheetname, Range, background_color_index=None):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        range_selected = ExcelSheet.Range(Range)
        try:
            range_selected.Interior.Color = background_color_index
        except:
            print('Error, background color index does not exit')
            exit()

    def set_bottom_border(self, file_name, sheetnumber_or_sheetname, Range):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelSheet = ExcelWorkbook.Worksheets(sheetnumber_or_sheetname)
        ExcelSheet.Range(Range).Borders(constants.xlEdgeBottom).LineStyle = constants.xlContinuous

    def copy_pastespecial(self, file_name, from_sheetnumber_or_sheetname, from_Range, to_sheetnumber_or_sheetname,
                          to_Range):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelWorkbook.Worksheets(from_sheetnumber_or_sheetname).Range(from_Range).Copy()
        ExcelWorkbook.Worksheets(to_sheetnumber_or_sheetname).Range(to_Range).PasteSpecial()

    def merge_cells(self, file_name, sheetnumber_or_sheetname, Range=None):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        try:
            if Range:
                ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).Range(Range).Merge()
            else:
                ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).Cells.Merge()
        except:
            exit()

    def unmerge_cells(self, file_name, sheetnumber_or_sheetname, Range=None):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        if Range:
            ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).Range(Range).UnMerge()
        else:
            ExcelWorkbook.Worksheets(sheetnumber_or_sheetname).Cells.UnMerge()

    def if_columnheader_is_searchcriteria_copy_to_next_sheet(self, file_name, from_sheetnumber_or_sheetname,
                                                             to_sheetnumber_or_sheetname,
                                                             searchcriteria_please_insert_a_list_using_square_bracket):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        lastcolumn = ExcelWorkbook.Worksheets(from_sheetnumber_or_sheetname).UsedRange.Columns.Count
        From_ExcelSheet = ExcelWorkbook.Worksheets(from_sheetnumber_or_sheetname)
        To_ExcelSheet = ExcelWorkbook.Worksheets(to_sheetnumber_or_sheetname)
        searchlist = list(searchcriteria_please_insert_a_list_using_square_bracket)
        for item in searchlist:
            lastcolumn2 = ExcelWorkbook.Worksheets(to_sheetnumber_or_sheetname).UsedRange.Columns.Count
            for i in range(1, lastcolumn):
                if From_ExcelSheet.Cells(1, i).Value == item:
                    From_ExcelSheet.Columns(i).Copy()
                    if To_ExcelSheet.Range("A1").Value is None:
                        To_ExcelSheet.Columns(1).PasteSpecial()
                    else:
                        To_ExcelSheet.Columns(lastcolumn2 + 1).PasteSpecial()

    def saveas_workbook(self, file_name, saveas_filename_use_Double_forward_slash_to_separate_path):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelWorkbook.SaveAs(saveas_filename_use_Double_forward_slash_to_separate_path)

    def save_workbook(self, file_name):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelWorkbook.Save()

    def exit_excel(self, file_name):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        ExcelWorkbook.Application.Quit()

    def winnie_only_find_dupliate(self, file_name, from_sheetnumber_or_sheetname, to_sheetnumber_or_sheetname):
        ExcelWorkbook = ExcelApp.Workbooks(file_name)
        from_ExcelSheet = ExcelWorkbook.Worksheets(from_sheetnumber_or_sheetname)
        to_ExcelSheet = ExcelWorkbook.Worksheets(to_sheetnumber_or_sheetname)
        lastrow = ExcelWorkbook.Worksheets(from_sheetnumber_or_sheetname).UsedRange.Rows.Count
        lastrow2 = ExcelWorkbook.Worksheets(to_sheetnumber_or_sheetname).UsedRange.Rows.Count
        lst_all = []
        lst_exact = []
        lst_broad = []
        for i in range(2, lastrow + 1):
            new_value = str(from_ExcelSheet.Cells(i, 1).Value).replace('+', '').replace('[', '').replace(']', '')
            lst_all.append([new_value, from_ExcelSheet.Cells(i, 4).Value])

        for item in lst_all:
            if item[1] == 'Exact':
                lst_exact.append(item)

        for item in lst_all:
            if item[1] == 'Broad':
                lst_broad.append(item)
        print(lst_exact)
        print(lst_broad)

        if len(lst_exact) > len(lst_broad):
            loop_count = len(lst_broad)
        else:
            loop_count = len(lst_exact)
        print(loop_count)

        for i in range(0, loop_count):
            # print(lst_exact[i][0])
            if lst_exact[i][0] in [item[0] for item in lst_broad]:
                print(lst_exact[i], 'True')


vb = Vba()

filename = 'all.XLSX'
# vb.open_file(path)
# new_workbook = vb.create_new_workbook()
# vb.saveas_workbook(new_workbook,'edison2')
# vb.loop_each_cell_in_range(filename,1,'A1:B1',2)
# vb.create_FormulaR1C1_for_range(filename,'Sheet1','D10:D20','=sum(R[0]C[+1]:R[0]C[+2])')
# vb.set_background_color(filename,'Sheet1','E1:F6',True)
# vb.copy_pastespecial(filename,'Sheet1','E1:F2','Sheet2','D20')
# vb.unmerge_cells(filename,'Sheet2')
# vb.if_columnheader_is_searchcriteria_copy_to_next_sheet(filename,'Sheet1','Sheet2',['quantity','Color','price','green'])
# vb.rename_sheet(filename,'Sheet2','new')
# vb.enter_value('Book3','Sheet1','A1','good')
# vb.insert_column_and_get_name_initial_and_autofill_to_lastrow('all.XLSX','Sheet1','last name','initial')
# vb.insert_column(filename,'Sheet1',2,1)
# vb.insert_row(filename,'Sheet1',2,1)
# vb.save_workbook(filename)
# vb.status_reference(filename,'Sheet1')
# vb.column_AutoFit(filename,'Sheet1')
# vb.row_AutoFit(filename,'Sheet1',5)
# vb.merge_cells(filename,'Sheet1','H1:I1')
# vb.set_Bold(filename,'Sheet1','H1:J1',True)
# vb.set_background_color(filename,'Sheet1','G1',250)
# vb.set_font(filename,'Sheet1','G2','Arial')
# vb.set_font_color(filename,'Sheet1','G3', 255)
# vb.saveas_workbook(filename,'all bruh')
# vb.save_workbook(filename)
# vb.exit_excel(filename)
# vb.create_newsheet(filename,'bad')
# vb.rename_sheet(filename,'bad','good')
# vb.filter(filename,'Sheet1',4,'red')
# vb.if_columnheader_is_searchcriteria_copy_to_next_sheet(filename,'Sheet1','saint james',['Color','price','quantity'])
# vb.status_reference(filename,'saint james')
# vb.set_rowwidth(filename,'Sheet1',2,100)
# vb.autofill(filename,'Sheet1','G3')
# vb.autofill(filename,'Sheet1','A6:D6')

