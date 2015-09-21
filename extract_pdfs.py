#!/usr/bin/env python

import olefile
import zipfile

import sys

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('USAGE: extract_pdfs.py PATH')
        print('Will extract PDF files from PATH (docx file) '
              'into current directory.')
        sys.exit(1)
    z = zipfile.ZipFile(sys.argv[1])
    ole_objects = [olefile.OleFileIO(z.read(i))
                   for i in z.filelist if 'oleObject' in i.filename]

    print('Found %d OLE files. Extracting...' % len(ole_objects))
    for i, ole in enumerate(ole_objects):
        contents = ole.openstream('CONTENTS')
        filename = 'extracted_%d.pdf' % i
        with open(filename, 'w') as f:
            print(filename)
            f.write(contents.read())
        print('Done.')
