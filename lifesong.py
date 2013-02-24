# lifesong.py - Python script for converting Microsoft Word documents
# to PDF format in bulk.  Currently only known to work on Windows
# with MS Office installed.  Also requires "comtypes" Python module.
#
# Basic usage: python lifesong.py -i "<indir>" -o "<outdir>" [--replace]
#
# Copyright (C) 2013  Johan Meiring
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
# Based on the accepted StackOverflow answer at
# http://stackoverflow.com/questions/6011115/doc-to-pdf-using-python

import sys
import os
import comtypes.client
import getopt
import ftplib
from ftplib import FTP

def usage():
    print """
Usage: python lifesong.py [OPTIONS]
    -h, --help      Displays usage details.
    -i, --indir     Input directory.
    -o, --outdir    Output directory.
    -r, --replace   Replace existing files (default=False).
    -H, --host      FTP hostname/IP address.
    -u, --username  FTP login username.
    -p, --password  FTP login password.
    -d, --directory Remote upload directory.
    --passive       Use PASSIVE transfer mode for FTP.
    """
    return

def main(argv):
    if len(argv) == 0:
        usage()
        sys.exit()

    try:
        opts, args = getopt.getopt(argv, "hi:o:rH:u:p:d:", ["help", "indir=", \
            "outdir=", "replace", "host=", "username=", "password=", \
            "directory=", "passive"])
    except getopt.GetoptError:
        usage()
        sys.exit(2)

    indir = ''
    outdir = ''
    replace = False
    ftp_passive = False
    ftp_host = ''
    ftp_username = ''
    ftp_password = ''
    ftp_directory = ''

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()
            sys.exit()
        elif opt in ("-i", "--indir"):
            indir = arg.strip()
        elif opt in ("-o", "--outdir"):
            outdir = arg.strip()
        elif opt in ("-r", "--replace"):
            replace = True
        elif opt in ("-H", "--host"):
            ftp_host = arg.strip()
        elif opt in ("-u", "--username"):
            ftp_username = arg.strip()
        elif opt in ("-p", "--password"):
            ftp_password = arg.strip()
        elif opt in ("-d", "--directory"):
            ftp_directory = arg.strip()
        elif opt in ("--passive"):
            ftp_passive = True

    if not os.path.isdir(indir):
        print 'Invalid input directory "%s"' % indir
        sys.exit(2)

    if not os.path.isdir(outdir):
        print 'Invalid output directory "%s"' % outdir
        sys.exit(2)

    wdDoNotSaveChanges = 0
    wdFormatPDF = 17

    if len(os.listdir(indir)) > 0:
        word = comtypes.client.CreateObject('Word.Application')

        ftp_client = None
        if ftp_host != '':
            try:
                ftp_client = FTP(ftp_host, ftp_username, ftp_password)
                ftp_client.set_pasv(ftp_passive)
                ftp_client.cwd(ftp_directory)
            except ftplib.all_errors as strerror:
                print "Something horrible has gone wrong..."
                print strerror
                ftp_client = None

        for i in os.listdir(indir):
            if i.endswith(".doc") or i.endswith(".docx"):
                print "%s - " % i,
                try:
                    outfile = "%s/%s.pdf" % (outdir, os.path.splitext(i)[0])
                    if os.path.exists(outfile) and not replace:
                        print "File exists... skipping..."
                        continue
                    doc = word.Documents.Open("%s/%s" % (indir, i))
                    doc.SaveAs(outfile, FileFormat=wdFormatPDF)
                    doc.Close(wdDoNotSaveChanges)

                    if ftp_client:
                        ftp_client.storbinary("STOR " + \
                            os.path.basename(outfile), \
                            open(outfile, "rb", 8192))

                    print "Done"
                except comtypes.COMError:
                    print "Whoops, file seems to be corrupt."
        word.Quit()
        if ftp_client:
            ftp_client.quit()

if __name__ == "__main__":
    main(sys.argv[1:])
    sys.exit()
