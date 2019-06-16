#!/usr/bin/env python
import yaml
import sys
from xlrd import open_workbook

#Read input configuration file
stream = file('input.yaml', 'r')
dict = yaml.safe_load(stream)


def main():
	if len(sys.argv) <= 1:
		print "Usage: python <programfile> <excel_file_name>"
		return
	#read the excel workbook
	wb = open_workbook(sys.argv[1])
	#sheet_number = len(wb.sheets())
	num_of_sheets = len(dict['sheets'])
	index =0
	while index < num_of_sheets:

		#Read value from input configuration file for a particular page
		page_num =  dict['sheets'][index]['number']
		col_num = dict['sheets'][index]['col']
		row_num = dict['sheets'][index]['row']
		file_name = dict['sheets'][index]['filename']

		#Convert column to integer number
		col_num_int = ord(col_num) - (ord('a'))

		#Get the particular sheet
		sheet = wb.sheet_by_index(page_num-1)

		#Open file to write (always overwrite mode)
		f = open(file_name, "w")
		print "Writing " + file_name

		i = row_num-1;

		#Read all rows
		while i < sheet.nrows:
			val =  sheet.cell(i, col_num_int).value
			#Write to file
			f.write("%s\n" % str(val))
			i+=1

		#Close file
		f.close()
		#move to next page
		index += 1

	print "Processed the provided excel file"




if __name__ == "__main__":
	main()
