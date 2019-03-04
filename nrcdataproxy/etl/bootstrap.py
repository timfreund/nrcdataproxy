from optparse import OptionParser
from urllib.request import urlopen

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
    '1990': 'CY90.xlsx',
    '1991': 'CY91.xlsx',
    '1992': 'CY92.xlsx',
    '1993': 'CY93.xlsx',
    '1994': 'CY94.xlsx',
    '1995': 'CY95.xlsx',
    '1996': 'CY96.xlsx',
    '1997': 'CY97.xlsx',
    '1998': 'CY98.xlsx',
    '1999': 'CY99.xlsx',
    '2000': 'CY00.xlsx',
    '2001': 'CY01.xlsx',
    '2002': 'CY02.xlsx',
    '2003': 'CY03.xlsx',
    '2004': 'CY04.xlsx',
    '2005': 'CY05.xlsx',
    '2006': 'CY06.xlsx',
    '2007': 'CY07.xlsx',
    '2008': 'CY08.xlsx',
    '2009': 'CY09.xlsx',
    '2010': 'CY10.xlsx',
    '2011': 'CY11.xlsx',
    '2012': 'CY12.xlsx',
    '2013': 'CY13.xlsx',
    '2014': 'CY14.xlsx',
    '2015': 'CY15.xlsx',
    '2016': 'CY16.xlsx',
    '2017': 'CY17.xlsx',
    '2018': 'CY18.xlsx',
    '2019': 'Current.xlsx',
    }


def download_year(output_dir, year):
    # TODO: see if the file exists before we download
    # TODO: let users switch between the mirror and authoritative sources
    print("Downloading %s" % year)
    archive = urlopen("%s/%s" % (archive_sources['mirror'], archives[year]))
    destination_path = os.path.sep.join((output_dir,
                                         'archives',
                                        archives[year]))
    if os.path.exists(destination_path):
        print("%s already exists, skipping" % year)
    else:
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




