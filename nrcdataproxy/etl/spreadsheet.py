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

    location_keys = ['lat_', 'long_',]
    mapped_names = {'material_inv0lved_cr': 'material_involved_cr'}
    negatives = ['N', 'NO']
    positives = ['Y', 'YES']
    multi_entry_sheets = ['MATERIAL_INVOLVED', 
                          'MATERIAL_INV0LVED_CR', 
                          'TRAINS_DETAIL',
                          'DERAILED_UNITS',
    ]
    
    def __init__(self, filename):
        self.filename = filename

    def __iter__(self):
        return self

    def mapped_name(self, name):
        return self.mapped_names.get(name, name)

    def next(self):
        raise StopIteration

    def seperate_location_data(self, record):
        location_data = {}
        for k, v in record.items():
            if k.startswith("location_"):
                location_data[k.replace("location_", "")] = v
                del record[k]
            for lk in self.location_keys:
                if k.startswith(lk):
                    location_data[k] = v
                    del record[k]
        record['location'] = location_data
        return record
        
    def scrub_data(self, record):
        for k, v in record.items():
            if isinstance(v, dict):
                if len(v.keys()):
                    record[k] = self.scrub_data(v)
                else:
                    del record[k]
            if isinstance(v, list):
                if not len(v):
                    del record[k]
                else:
                    for vprime in v:
                        self.scrub_data(vprime)
            if isinstance(v, basestring) and v == "" or v is None:
                del record[k]
            if v in self.negatives:
                record[k] = False
            if v in self.positives:
                record[k] = True
        return record

    def extract_data(self, repository):
        for record in self:
            repository.save(self.seperate_location_data(self.scrub_data(record)))
    
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
                rows = sheet_keys.get(cell.value, [])
                rows.append(row)
                sheet_keys[cell.value] = rows

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

            if len(detail_data) > 1 and name not in self.multi_entry_sheets:
                msg = "%s: multiple entries in a single entry sheet: %s" % (data['seqnos'],
                                                                            name)
                raise Exception(msg)

            for dd in detail_data:
                if dd.has_key('SEQNOS'):
                    del dd['SEQNOS']

            if name == 'INCIDENT_COMMONS' and len(detail_data):
                for k, v in detail_data[0].items():
                    data[k.lower()] = v
            else:
                lname = self.mapped_name(name.lower())
                if name in self.multi_entry_sheets:
                    data[lname] = []
                    for dd in detail_data:
                        data[lname].append(dd)
                elif len(detail_data):
                    data[lname] = detail_data[0]
        return data
                
    def incident_details(self, sheet_name, columns, seqnos):
        sheet = self.workbook.sheet_by_name(sheet_name)
        all_data = []

        if self.metadata['positions'][sheet_name].has_key(seqnos):
            for sheet_row in self.metadata['positions'][sheet_name][seqnos]:
                data = {}
                for colx, col_name in enumerate(columns):
                    cell = sheet.cell(sheet_row, colx)
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE and cell.value != 0.0:
                        try:
                            value = xlrd.xldate_as_tuple(cell.value, 0)
                            value = datetime(*value).isoformat()
                        except ValueError:
                            print("%s.%s couldn't convert %s" % (sheet_name,
                                                                 str(seqnos),
                                                                 str(cell.value)))
                    elif cell.ctype == xlrd.XL_CELL_TEXT:
                        value = value.strip()

                    data[col_name.lower()] = value
                del data['seqnos']
                all_data.append(data)
        return all_data

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
                    rows = sheet_keys.get(cell.value, [])
                    rows.append(row)
                    sheet_keys[cell.value] = rows
                    
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

            if len(detail_data) > 1 and name not in self.multi_entry_sheets:
                msg = "%s: multiple entries in a single entry sheet: %s" % (data['seqnos'],
                                                                            name)
                raise Exception(msg)

            for dd in detail_data:
                if dd.has_key('SEQNOS'):
                    del dd['SEQNOS']

            if name == 'INCIDENT_COMMONS' and len(detail_data):
                for k, v in detail_data[0].items():
                    data[k.lower()] = v
            else:
                lname = self.mapped_name(name.lower())
                if name in self.multi_entry_sheets:
                    data[lname] = []
                    for dd in detail_data:
                        data[lname].append(dd)
                elif len(detail_data):
                    data[lname] = detail_data[0]

        return data

    def incident_details(self, sheet_name, columns, seqnos):
        sheet = self.workbook.get_sheet_by_name(sheet_name)
        all_data = []

        if self.metadata['positions'][sheet_name].has_key(seqnos):
            for sheet_row in self.metadata['positions'][sheet_name][seqnos]:
                data = {}
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
                del data['seqnos']
                all_data.append(data)
        return all_data
        
extractors = [
    XlsExtractor,
    XlsxExtractor
    ]

def extractor_command():
    """
    This command extracts data from National Response Center (NRC) incident archive
    spreadsheets into programmer friendly JSON documents.
    """

    parser = OptionParser(usage="usage: %%prog [options]\n%s" % extractor_command.__doc__)
    parser.add_option("-f", "--input-file", dest="input_file")
    parser.add_option("-i", "--input-directory", dest="input_directory",
                      help="Parse all files in the provided directory")
    parser.add_option("-u", "--nrc-data-url", dest="data_url",
                      help="NRCDataProxy URL")


    (options, args) = parser.parse_args()

    if options.data_url is None:
        print("Please provide an NRC Data URL")
        parser.print_help()
        sys.exit(-1)

    data_storage = NRCDataClient(options.data_url)

    if options.input_file is None and options.input_directory is None:
        print("You must supply an input file or input directory")
        parser.print_help()
        sys.exit(-1)

    file_paths = []

    if options.input_file:
        if not os.path.exists(options.input_file):
            print("The input file does not exist")
            sys.exit(-1)
        else:
            file_paths.append(os.path.abspath(options.input_file))

    if options.input_directory:
        input_dir = options.input_directory
        if not os.path.exists(input_dir):
            print("The input directory does not exist")
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
            print("No extractor available for %s" % file_path)
        else:
            extractor.extract_data(data_storage)
            
    print("Extracted data from NRC spreadsheets")

    
