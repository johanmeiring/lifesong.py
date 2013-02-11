# Based on the accepted StackOverflow answer at http://stackoverflow.com/questions/6011115/doc-to-pdf-using-python
import sys
import os
import comtypes.client
import getopt

def usage():
  print """
Fill stuff in here...
  """
  return

def main(argv):
  if len(argv) == 0:
    usage()
    sys.exit()

  try:
    opts, args = getopt.getopt(argv, "i:o:r", ["indir=", "outdir=", "replace"])
  except getopt.GetoptError:
    usage()
    sys.exit(2)

  indir = ''
  outdir = ''
  replace = False

  for opt, arg in opts:
    if opt in ("-i", "--indir"):
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

  wdFormatPDF = 17

  # in_file = os.path.abspath(sys.argv[1])
  # out_file = os.path.abspath(sys.argv[2])

  if len(os.listdir(indir)) > 0:
    word = comtypes.client.CreateObject('Word.Application')
    for i in os.listdir(indir):
      if i.endswith(".doc") or i.endswith(".docx"):
        print i
        try:
          doc = word.Documents.Open(indir + "/" + i)

          thesplit = os.path.splitext(i)
          print thesplit
          doc.SaveAs(outdir + "/" + thesplit[0] + '.pdf', FileFormat=wdFormatPDF)
          doc.Close(0)
        except comtypes.COMError:
          print "Whoops"
    word.Quit()

if __name__ == "__main__":
  main(sys.argv[1:])
  sys.exit()
