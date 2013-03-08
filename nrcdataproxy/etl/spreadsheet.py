from datetime import datetime, time
from optparse import OptionParser
from nrcdataproxy.client import NRCDataClient
import mimetypes
import os
import openpyxl
import sys
import xlrd

class SpreadsheetExtractor():
    mimetype = None
    mapped_names = {'material_inv0lved_cr': 'material_involved_cr'}

    def __init__(self, filename):
        self.filename = filename

    def __iter__(self):
        return self

    def mapped_name(self, name):
        return self.mapped_names.get(name, name)

    def next(self):
        raise StopIteration

    def extract_data(self, repository):
        for record in self:
            repository.save(record)
    
class XlsExtractor(SpreadsheetExtractor):
    mimetype = 'application/vnd.ms-excel'

    def __init__(self, *args):
        SpreadsheetExtractor.__init__(self, *args)
        self.workbook = xlrd.open_workbook(self.filename)
        self.metadata = self.load_metadata()

    def load_metadata(self):
        metadata = {'layout': [],
                    'positions': {}}
        metadata['current_row'] = 1
        
        for name in self.workbook.sheet_names():
            sheet = self.workbook.sheet_by_name(name)
            columns = []
            for x in range(0, sheet.ncols):
                columns.append(sheet.cell(0, x).value)

            if columns[0] == u' ':
                columns[0] = 'SEQNOS'

            sheet_keys = {}
            for row, cell in enumerate(sheet.col(0)):
                sheet_keys[cell.value] = row

            metadata['layout'].append((name, columns))
            metadata['positions'][name] = sheet_keys
        return metadata

    def next(self):
        data = {}

        name, columns = self.metadata['layout'][0]
        sheet = self.workbook.sheet_by_name(name)
        rowx = self.metadata['current_row']

        if rowx == -1:
            raise StopIteration
        
        for colx, col_name in enumerate(columns):
            cell = sheet.cell(rowx, colx)
            value = cell.value
            if cell.ctype == xlrd.XL_CELL_DATE:
                value = xlrd.xldate_as_tuple(cell.value, 0)
                value = datetime(*value).isoformat()
            elif cell.ctype == xlrd.XL_CELL_TEXT:
                value = value.strip()
            data[col_name.lower()] = value

        try:
            sheet.cell(rowx+1, 0).value
            self.metadata['current_row'] = rowx + 1
        except IndexError:
            self.metadata['current_row'] = -1

        for name, columns in self.metadata['layout'][1:]:
            detail_data = self.incident_details(name, columns, data['seqnos'])

            if detail_data.has_key('SEQNOS'):
                del detail_data['SEQNOS']

            if name == 'INCIDENT_COMMONS':
                for k, v in detail_data.items():
                    data[k.lower()] = v
            else:
                lname = self.mapped_name(name.lower())
                data[lname] = {}
                for k, v in detail_data.items():
                    data[lname][k.lower()] = v

        return data
                
    def incident_details(self, sheet_name, columns, seqnos):
        sheet = self.workbook.sheet_by_name(sheet_name)
        data = {}
        
        if self.metadata['positions'][sheet_name].has_key(seqnos):
            sheet_row = self.metadata['positions'][sheet_name][seqnos]
            for colx, col_name in enumerate(columns):
                cell = sheet.cell(sheet_row, colx)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE and cell.value != 0.0:
                    try:
                        value = xlrd.xldate_as_tuple(cell.value, 0)
                        value = datetime(*value).isoformat()
                    except ValueError:
                        print "%s.%s couldn't convert %s" % (sheet_name,
                                                             str(seqnos),
                                                             str(cell.value))
                elif cell.ctype == xlrd.XL_CELL_TEXT:
                    value = value.strip()
                        
                data[col_name.lower()] = value
            
        return data

class XlsxExtractor(SpreadsheetExtractor):
    mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    def __init__(self, *args):
        SpreadsheetExtractor.__init__(self, *args)
        self.workbook = openpyxl.reader.excel.load_workbook(self.filename)
        self.metadata = self.load_metadata()

    def load_metadata(self):
        metadata = {'layout': [],
                    'positions': {}}
        metadata['current_row'] = 1

        for sheet in self.workbook.worksheets:
            name = sheet.title
            columns = []
            for x in range(0, len(sheet.column_dimensions.keys())):
                # CY11 has a null value in the last column of INCIDENT_DETAILS.
                # We won't take any column that has a null header
                if sheet.cell(row=0, column=x).value is not None:
                    columns.append(sheet.cell(row=0, column=x).value)

            if len(columns):
                if columns[0] == u' ':
                    columns[0] = u'SEQNOS'

                sheet_keys = {}
                for row, cell in enumerate(sheet.columns[0]):
                    sheet_keys[cell.value] = row

                metadata['layout'].append((name, columns))
                metadata['positions'][name] = sheet_keys

        return metadata
                
    def next(self):
        data = {}

        name, columns = self.metadata['layout'][0]
        sheet = self.workbook.get_sheet_by_name(name)
        rowx = self.metadata['current_row']

        if rowx == -1:
            raise StopIteration

        for colx, col_name in enumerate(columns):
            cell = sheet.cell(row=rowx, column=colx)
            value = cell.value
            if isinstance(value, datetime):
                value = value.isoformat()
            elif isinstance(value, time):
                value = None
            elif isinstance(value, unicode):
                value = value.strip()
            data[col_name.lower()] = value

        if sheet.cell(row=rowx+1, column=0).value != None:
            self.metadata['current_row'] = rowx + 1
        else:
            self.metadata['current_row'] = -1

        for name, columns in self.metadata['layout'][1:]:
            detail_data = self.incident_details(name, columns, data['seqnos'])

            if detail_data.has_key('SEQNOS'):
                del detail_data['SEQNOS']

            if name == 'INCIDENT_COMMONS':
                for k, v in detail_data.items():
                    data[k.lower()] = v
            else:
                data[self.mapped_name(name.lower())] = detail_data
        return data

    def incident_details(self, sheet_name, columns, seqnos):
        sheet = self.workbook.get_sheet_by_name(sheet_name)
        data = {}

        if self.metadata['positions'][sheet_name].has_key(seqnos):
            sheet_row = self.metadata['positions'][sheet_name][seqnos]
            for coly, col_name in enumerate(columns):
                cell = sheet.cell(row=sheet_row, column=coly)
                value = cell.value
                if isinstance(value, datetime):
                    value = value.isoformat()
                elif isinstance(value, time):
                    value = None
                elif isinstance(value, unicode):
                    value = value.strip()
                data[col_name.lower()] = value
        return data
        
extractors = [
    XlsExtractor,
    XlsxExtractor
    ]

def extractor_command():
    """
    This command extracts data from National Response Center (NRC) incident archive
    spreadsheets into programmer friendly JSON documents.
    """

    parser = OptionParser(usage="usage: %%prog OutputRepositoryClass [options]\n%s" % extractor_command.__doc__)
    parser.add_option("-f", "--input-file", dest="input_file")
    parser.add_option("-i", "--input-directory", dest="input_directory",
                      help="Parse all files in the provided directory")
    parser.add_option("-u", "--nrc-data-url", dest="data_url",
                      help="NRCDataProxy URL")


    (options, args) = parser.parse_args()

    if options.data_url is None:
        print "Please provide an NRC Data URL"
        parser.print_help()
        sys.exit(-1)

    data_storage = NRCDataClient(options.data_url)

    if options.input_file is None and options.input_directory is None:
        print "You must supply an input file or input directory"
        parser.print_help()
        sys.exit(-1)

    file_paths = []

    if options.input_file:
        if not os.path.exists(options.input_file):
            print "The input file does not exist"
            sys.exit(-1)
        else:
            file_paths.append(os.path.abspath(options.input_file))

    if options.input_directory:
        input_dir = options.input_directory
        if not os.path.exists(input_dir):
            print "The input directory does not exist"
            sys.exit(-1)
        else:
            input_dir = os.path.abspath(input_dir)
            for name in os.listdir(input_dir):
                file_paths.append(os.path.sep.join((input_dir, name)))

    for file_path in file_paths:
        extractor = None
        for extract_class in extractors:
            if extract_class.mimetype == mimetypes.guess_type(file_path)[0]:
                extractor = extract_class(file_path)
                
        if extractor is None:
            print "No extractor available for %s" % file_path
        else:
            extractor.extract_data(data_storage)
            
    print "Extracted data from NRC spreadsheets"

    
