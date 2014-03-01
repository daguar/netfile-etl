import xlrd
import os
import glob

xlsx_files_in_current_dir = glob.glob("*.xlsx")
for excel_filename in xlsx_files_in_current_dir:
  print "Opening " + excel_filename + "..."
  workbook = xlrd.open_workbook("./" + excel_filename)
  filename_without_extension = excel_filename.split(".")[0]
  print filename_without_extension
  sheet_names = workbook.sheet_names()
  for sheet_name in sheet_names:
    print "Writing " + sheet_name + ".csv..."
    os.system("in2csv " + filename_without_extension + " --sheet " + "\"" + sheet_name + "\"" + " > " + "csv_files/" + sheet_name + ".csv")
