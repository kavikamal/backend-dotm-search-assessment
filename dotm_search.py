#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!

import sys
import glob
import os
import zipfile

def main(): 
    if len(sys.argv) < 2:
        print 'Please enter search text'
        sys.exit(1)
    search_text = sys.argv[1] 
    if len(sys.argv)>3 and sys.argv[2]=="--dir":
        directory = sys.argv[3]
    else:    
        directory = os.getcwd()  
    searchFile(search_text,directory)

def searchFile(search_text,directory):
    files = glob.glob(os.path.join(directory, '*.dotm'))
    no_files_matched=0
    print "Searching directory {} for text {} ...".format(directory,search_text)
    for file in files:
        zf = zipfile.ZipFile(file, 'r')
        with zf.open('word/document.xml', 'r') as file:
            for line in file:
                if search_text in line:
                    no_files_matched += 1
                    index = line.find(search_text)
                    print "Match found in file {}/{}".format(directory, file)        
                    print "..." + line[index-40:index+40] + "..."
    print "Total dotm files searched: {}".format(len(files))
    print "Total dotm files matched: {}".format(no_files_matched) 

if __name__ == '__main__':
  main()      