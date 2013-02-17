from optparse import OptionParser
import os
import platform
import sys
import urllib2

archives = {
    '1982': 'http://www.nrc.uscg.mil/download/nrc_82.exe',
    '1983': 'http://www.nrc.uscg.mil/download/nrc_83.exe',
    '1984': 'http://www.nrc.uscg.mil/download/nrc_84.exe',
    '1985': 'http://www.nrc.uscg.mil/download/nrc_85.exe',
    '1986': 'http://www.nrc.uscg.mil/download/nrc_86.exe',
    '1987': 'http://www.nrc.uscg.mil/download/nrc_87.exe',
    '1988': 'http://www.nrc.uscg.mil/download/nrc_88.exe',
    '1989': 'http://www.nrc.uscg.mil/download/nrc_89.exe',
    '1990': 'http://www.nrc.uscg.mil/download/CY90.exe',
    '1991': 'http://www.nrc.uscg.mil/download/CY91.exe',
    '1992': 'http://www.nrc.uscg.mil/download/CY92.exe',
    '1993': 'http://www.nrc.uscg.mil/download/CY93.exe',
    '1994': 'http://www.nrc.uscg.mil/download/CY94.exe',
    '1995': 'http://www.nrc.uscg.mil/download/CY95.exe',
    '1996': 'http://www.nrc.uscg.mil/download/CY96.exe',
    '1997': 'http://www.nrc.uscg.mil/download/CY97.exe',
    '1998': 'http://www.nrc.uscg.mil/download/CY98.exe',
    '1999': 'http://www.nrc.uscg.mil/download/CY99.exe',
    '2000': 'http://www.nrc.uscg.mil/download/CY00.exe',
    '2001': 'http://www.nrc.uscg.mil/download/CY01.exe',
    '2002': 'http://www.nrc.uscg.mil/download/CY02.exe',
    '2003': 'http://www.nrc.uscg.mil/download/CY03.exe',
    '2004': 'http://www.nrc.uscg.mil/download/CY04.exe',
    '2005': 'http://www.nrc.uscg.mil/download/CY05.exe',
    '2006': 'http://www.nrc.uscg.mil/download/CY06.exe',
    '2007': 'http://www.nrc.uscg.mil/download/CY07.exe',
    '2008': 'http://www.nrc.uscg.mil/download/CY08.exe',
    '2009': 'http://www.nrc.uscg.mil/download/CY09.exe',
    '2010': 'http://www.nrc.uscg.mil/download/CY10.exe',
    '2011': 'http://www.nrc.uscg.mil/download/CY11.exe',
    '2012': 'http://www.nrc.uscg.mil/download/CY12.exe',
    }


def download_year(output_dir, year):
    print "Downloading %s" % year
    archive = urllib2.urlopen(archives[year])
    destination_path = os.path.sep.join((output_dir,
                                         'archives',
                                        archives[year].split('/')[-1]))
    if os.path.exists(destination_path):
        print "%s already exists, skipping" % year
    else:
        destination_file = open(destination_path, "w")
        destination_file.write(archive.read())
        destination_file.close()

    if platform.system() in ['Linux', 'Darwin']:
        os.system('unzip -n -d %s %s' % (os.path.sep.join([output_dir, 'source']),
                                         destination_path))
    else:
        # not sure if this will work, need to get on a windows 
        # machine to test it.
        import pdb; pdb.set_trace()
        os.system(destination_path)

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
        print "You must provide an output directory"
        parser.print_help()
        sys.exit(-1)
    elif not os.path.exists(options.output_directory):
        os.makedirs(options.output_directory)

        if not os.path.exists(options.output_directory):
            print "The provided output directory does not exist, and directory creation failed."
            parser.print_help()
            sys.exit(-1)
        
    for sub_dir in ['archives', 'source']:
        sub_dir_path = os.path.sep.join([options.output_directory, sub_dir])
        if not os.path.exists(sub_dir_path):
            os.makedirs(sub_dir_path)

    if options.all_years == False and len(args) == 0:
        print "You must supply one or more years or explictly request all data with the -a flag.\n"
        parser.print_help()
        sys.exit(-1)

    if options.all_years:
        print "Downloading all years"
        for year in archives.keys():
            download_year(options.output_directory, year)
    else:
        for year in args:
            download_year(options.output_directory, year)




