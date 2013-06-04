import sys
import glob


reload(sys)
sys.setdefaultencoding('iso-8859-1')

try:
    from xlrd import open_workbook
except ImportError, error:
    print "`from xlrd import open_workbook` failed.  you might want to `virtualenv parse_xls` and pip install xlrd"


scriptPythonVersion = (2,7,2)
hostPythonVersion = sys.version_info
if hostPythonVersion < scriptPythonVersion:
    sys.exit('please use a virtualenv or install at least python {0} to use this script'.format(scriptPythonVersion))


def recursive_find(dir=None, file_extension=None, filter_string='~'):
    """
    populate an array with files in a 'dir' matching 'file_extension'
    """
    if dir is not None and file_extension is not None:
        #if recursive file lookups are necessary, use fnmatch and os.walk
        file_array = glob.glob("{0}/*.{1}".format(dir,file_extension))
        #glob expands tildes for paths, so we should not be seeing tildes at this point, esp. as ~ is common
        #in microsoft land for temporary file backups
        filtered_file_array = [ filtered_file_array for filtered_file_array in file_array if filter_string not in filtered_file_array]
        return filtered_file_array
    else:
        print "dir or file not specified"


def column_from_xls(file=None, sheet=1, column_name=None, column_name_in_row=None, start_at_row=None, end_at_row=None):
    """
    returns column values from an xls file.  column_name_in_row is useful in finding the column name in case
    the spreadsheet author puts names in rows instead of letting the spreadsheet be a spreadsheet
    """
    def get_indice_by_name(sheet, column_name, column_name_in_row):
        """
        find the column indice by name for passing row contents into a list
        """
        for col in range(sheet.ncols):
            if sheet.col_values(col)[column_name_in_row] == column_name:
                return col

    if file is not None:
        file = open_workbook(file)
        current_sheet = file.sheet_by_index(sheet)
        row_number = get_indice_by_name(current_sheet, column_name, column_name_in_row)
        column_values = current_sheet.col_values(row_number)
        return_set = column_values[start_at_row:end_at_row]
        return return_set

def main():
    for i in recursive_find('/Users/tristanfisher/Downloads/530-order copy', 'xlsx'):
        return_val = column_from_xls(i, column_name='SN (Scan from OEM box)', column_name_in_row=2, start_at_row=3)
        for i in return_val:
            print i


if __name__ == '__main__':
    main()