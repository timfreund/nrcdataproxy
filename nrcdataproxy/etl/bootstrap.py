from optparse import OptionParser
from urllib.request import urlopen

import hashlib
import os
import platform
import sys

archive_sources = {
    'authoritative': 'http://www.nrc.uscg.mil/FOIAFiles',
    'mirror': 'https://resources.codemuxer.com/data-pipelines/nrc-data/',
}

archives = {
    # '1982': 'http://www.nrc.uscg.mil/download/nrc_82.exe',
    # '1983': 'http://www.nrc.uscg.mil/download/nrc_83.exe',
    # '1984': 'http://www.nrc.uscg.mil/download/nrc_84.exe',
    # '1985': 'http://www.nrc.uscg.mil/download/nrc_85.exe',
    # '1986': 'http://www.nrc.uscg.mil/download/nrc_86.exe',
    # '1987': 'http://www.nrc.uscg.mil/download/nrc_87.exe',
    # '1988': 'http://www.nrc.uscg.mil/download/nrc_88.exe',
    # '1989': 'http://www.nrc.uscg.mil/download/nrc_89.exe',
    '1990': ('CY90.xlsx', '7f348c04fc213c6863920a05e3d7a2686d03443e9c1fb9904b6eb816525d7644'),
    '1991': ('CY91.xlsx', '7f14c922c46e726ff903fba86233b0dd26c6959a40250638007409dc68a89de7'),
    '1992': ('CY92.xlsx', '4c4325f4ebfbc230a15fb00b813409bb0ac107fe8bf1b4e328952efd080e79e1'),
    '1993': ('CY93.xlsx', 'ab5c63a2e1c2938b26b1443c0f1dbd8494fd8f9a4ae11d1b7ab2b71ccc405dbd'),
    '1994': ('CY94.xlsx', '9fb59a1f8b1de91bf1b19e3ff2d0911b70847bc066041ac2c199475e897efa49'),
    '1995': ('CY95.xlsx', '7747578b77aa1da6a617322660010ed5b3766ecb86bdcda966765d171c9eed1e'),
    '1996': ('CY96.xlsx', 'ac7a2735bfe80df5a2b54f7c71ebaecc69e9f347992566ad4a352e06505b88fd'),
    '1997': ('CY97.xlsx', '4a62f801f039d47d1d53983d8354788438b8a311dbeb4052eded2cc731e28caf'),
    '1998': ('CY98.xlsx', 'f5d31d16236645b77c16b7850266d7546735bc20c86e72064ed08adf2a83d929'),
    '1999': ('CY99.xlsx', 'ddd2acc3c63f7da70d959b292e845dffeff8984d74f557bbab54aafc753f60df'),
    '2000': ('CY00.xlsx', '2294092c7b2fbf500b95c89869d5128a422967672d554a4df9802a82b2db5d35'),
    '2001': ('CY01.xlsx', 'a0932452b76c54c99d4ad4d06cc907470abb233b4174ff5574d4dd0071f053b4'),
    '2002': ('CY02.xlsx', '69738f2f12ee0bee48042b9a760313746f3f775a97a4c9ef304937af11b99e4f'),
    '2003': ('CY03.xlsx', 'bdeb4045a78bcad341a60e31f29833eaa025ad1feca962605ebe59420fca803e'),
    '2004': ('CY04.xlsx', 'c32005d5c9971e7eb81344e17861e3ff5f96666664452ec744a8be9fc7f36507'),
    '2005': ('CY05.xlsx', '2d932903a1955d52793a3b6692861559b4eed2a6daf75d3587c8025c0b2a4184'),
    '2006': ('CY06.xlsx', '5ca1c67f045e5fc832a96d8a1dca83bf0de0a914edfb66718690fe3d35cb80e2'),
    '2007': ('CY07.xlsx', '6b68ffd59ed8a0a8ee6e76f5deda2dcac64c649f4f766fe24de17ccb56c98e8b'),
    '2008': ('CY08.xlsx', 'ead423b5470bc1c61803cc3f722a97604dd5bef244597e10f728f0e5a584ab9b'),
    '2009': ('CY09.xlsx', '1191e146b09209c4e8c5a1a63f65d2b1c04e2155f8ccb85f8a1a2d663683e5ae'),
    '2010': ('CY10.xlsx', 'd8fe9e4f8769e4d883f31e228a3856768cf44475789287872e067918de4d508e'),
    '2011': ('CY11.xlsx', '3b276967627a025490636c1d6bc9c085a2c7362fd6362274beb909c38afef333'),
    '2012': ('CY12.xlsx', '4f8531112d46f3fe1401eb6c037a0c2d6f17853d5d975fe3cf599b1a82b10d85'),
    '2013': ('CY13.xlsx', '190260eeaf29d10f2f0f038faed75333d0d46bb217d004fc79a5c155cd8f02cf'),
    '2014': ('CY14.xlsx', '5255c2dcf68f923b01ba62c4cebf27065101c12a623525071a851cc8628fd9c8'),
    '2015': ('CY15.xlsx', 'c1e11c410e0ba9f5ca2acf26e50dce42337e518ad904af05c0d4133596a97beb'),
    '2016': ('CY16.xlsx', '35fbdba3489692ba8c21f0405a64be808997b5b8ebb4f24128f72a65920e197e'),
    '2017': ('CY17.xlsx', '189e2d12756f641106dba51d09fbf9f05911f16aa097e3c992d85f67128fc522'),
    '2018': ('CY18.xlsx', '9e86c8ee58bb2655b76148b8e63c1796273b8a6bc68b1e2c0efb6b6c49b50d31'),
    '2019': ('Current.xlsx', '7afd874bf947398471ae859ef460fc8be639a1d2d9ef78f57036499970f378c0'),
    }


def download_year(output_dir, year):
    # TODO: let users switch between the mirror and authoritative sources
    filename, sha256 = archives[year]
    destination_path = os.path.sep.join((output_dir,
                                         'archives',
                                         filename))
    if os.path.exists(destination_path):
        with open(destination_path, 'rb') as data_file:
            hash = hashlib.sha256()
            hash.update(data_file.read())
            if sha256 == hash.hexdigest():
                print("%s exists and checksum matches, skipping" % year)
                return
            else:
                print("%s checksum doesn't match, overwriting" % year)

    print("Downloading %s" % year)
    archive = urlopen("%s/%s" % (archive_sources['mirror'], filename))
    destination_file = open(destination_path, "wb")
    destination_file.write(archive.read())
    destination_file.close()

def spreadsheet_downloader():
    """
    This command downloads yearly archives from the National 
    Response Center (NRC) incident archive.  You'll end up with
    self-extracting EXE's.  The following years are available:
    1990 through 2010
    """

    parser = OptionParser(usage="usage: %%prog [year01 year02 ...]\n%s" % spreadsheet_downloader.__doc__)
    parser.add_option("-a", "--all-years", action="store_true",
                      dest="all_years", default=False)
    parser.add_option("-o", "--output-directory", dest="output_directory"
                      )

    
    (options, args) = parser.parse_args()
    
    if options.output_directory is None:
        print("You must provide an output directory")
        parser.print_help()
        sys.exit(-1)
    elif not os.path.exists(options.output_directory):
        os.makedirs(options.output_directory)

        if not os.path.exists(options.output_directory):
            print("The provided output directory does not exist, and directory creation failed.")
            parser.print_help()
            sys.exit(-1)
        
    for sub_dir in ['archives', 'source']:
        sub_dir_path = os.path.sep.join([options.output_directory, sub_dir])
        if not os.path.exists(sub_dir_path):
            os.makedirs(sub_dir_path)

    if options.all_years == False and len(args) == 0:
        print("You must supply one or more years or explictly request all data with the -a flag.\n")
        parser.print_help()
        sys.exit(-1)

    if options.all_years:
        print("Downloading all years")
        for year in archives.keys():
            download_year(options.output_directory, year)
    else:
        for year in args:
            download_year(options.output_directory, year)




