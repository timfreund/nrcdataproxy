from nrcdataproxy.etl import spreadsheet
import json
import os
import unittest

json_decoder = json.JSONDecoder()

def get_extracted_record(seqnos):
    record_path = os.path.sep.join([os.path.dirname(__file__),
                                   'nrc-%d.json' % seqnos])
    record_file = open(record_path, 'r')
    record = json_decoder.decode(record_file.read())
    record_file.close()
    return record

class SpreadsheetExtractorTest(unittest.TestCase):
    def setUp(self):
        self.extractor = spreadsheet.SpreadsheetExtractor(None)

    def test_scrub_data_no_nones(self):
        original_record = get_extracted_record(1034226)
        record = self.extractor.scrub_data(original_record)
        
        for k, v in record.items():
            self.assertIsNotNone(v, "%s is None, was not removed in scrub_data" % k)
        print record
