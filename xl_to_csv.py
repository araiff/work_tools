import xlrd
import csv
import os
import datetime

def open_and_convert(path, filename, sheetname):
    """
    Opens .xls file at path\filename and converts data from sheetname to a .csv file. Row 2 contains 
    header and rows 17+ contain data.
    """
    wb = xlrd.open_workbook(path+'\\'+filename)
    new_filename = filename.replace('.xls','.csv')
    sh = wb.sheet_by_name(sheetname)
    try: 
        os.makedirs(path)
    except OSError:
        if not os.path.isdir(path):
            raise
    if not os.path.exists('C:\\csvexport\\'+sheetname):
        os.makedirs('C:\\csvexport\\'+sheetname)
    your_csv_file = open('C:\\csvexport\\'+sheetname+'\\'+new_filename, 'wb')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
    for rownum in xrange(sh.nrows):
        if rownum < 17 and not rownum == 2: 
            continue
        if rownum == 2:
            wr.writerow(sh.row_values(rownum))
        else:
			write_list = sh.row_values(rownum)
			write_list[0] = datetime.datetime(*xlrd.xldate_as_tuple(write_list[0],wb.datemode))
			wr.writerow(write_list)
    your_csv_file.close()

def select_file(rootdir):
	total_files = count_files(rootdir)
	count = 0
	print 'Convert the %i files listed above? Type y for yes or n for no.' % total_files
	if raw_input('> ') == 'n':
	    raise SystemExit
	for subdir, dirs, files in os.walk(rootdir):
	    for file in files:
		    if file.endswith('.xls') and file.startswith('RX') and subdir.endswith('0'):
				print '%i/%i complete: Converting %s...' % (count, total_files, file)
				count += 1
				open_and_convert(subdir, file, 'R'+subdir[-3:])
		    
def count_files(rootdir):
	count = 0
	for subdir, dirs, files in os.walk(rootdir):
		print '-------------------------'
		print '%s\\' % subdir
		for file in files:
		    if file.endswith('.xls') and file.startswith('RX') and subdir.endswith('0'):
				print file
				count = count + 1
	return count
