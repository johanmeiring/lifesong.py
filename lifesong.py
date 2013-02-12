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

def usage():
    print """
Usage: python lifesong.py [OPTIONS]
    -h, --help      Displays usage details.
    -i, --indir     Input directory.
    -o, --outdir    Output directory.
    -r, --replace   Replace existing files (default=False)
    """
    return

def main(argv):
    if len(argv) == 0:
        usage()
        sys.exit()

    try:
        opts, args = getopt.getopt(argv, "hi:o:r", ["help", "indir=", "outdir=", \
            "replace"])
    except getopt.GetoptError:
        usage()
        sys.exit(2)

    indir = ''
    outdir = ''
    replace = False

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

    if not os.path.isdir(indir):
        print 'Invalid input directory "' + indir + '"'
        sys.exit(2)

    if not os.path.isdir(outdir):
        print 'Invalid output directory "' + output + '"'
        sys.exit(2)

    wdDoNotSaveChanges = 0
    wdFormatPDF = 17

    if len(os.listdir(indir)) > 0:
        word = comtypes.client.CreateObject('Word.Application')
        for i in os.listdir(indir):
            if i.endswith(".doc") or i.endswith(".docx"):
                print i + " - ",
                try:
                    doc = word.Documents.Open(indir + "/" + i)
                    outfile = outdir + "/" + os.path.splitext(i)[0] + '.pdf'
                    if os.path.exists(outfile) and not replace:
                        print "File exists... skipping..."
                        continue
                    doc.SaveAs(outfile, FileFormat=wdFormatPDF)
                    doc.Close(wdDoNotSaveChanges)
                    print "Done"
                except comtypes.COMError:
                    print "Whoops, file seems to be corrupt."
        word.Quit()

if __name__ == "__main__":
    main(sys.argv[1:])
    sys.exit()
