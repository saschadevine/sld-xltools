#!/usr/bin/env/python3
#
# eval2csv.py
# Evaluate all cell values (including formulas) in a specified sheet
# of an XLSX source and write out to CSV
#
# SLD 2025

# License
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
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#

# Algorithm
#
# 1) Accept 5x command line arguments (clargs):
#       i)   XLSX source file
#       ii)  Source sheet from XLSX input file
#       iii) Source range of cells from source sheet of XLSX source file
#       iv)  CSV dest file
#       v)   Truncate on empty first element (boolean)
# 2) Verify XLSX source and CSV dest files exist
# 3) Verify source sheet exists in XLSX source file 
# 4) Evaluate values of each cell of each row in the source, until reach
#    row where first element is null
#

# Imports
import sys, os, openpyxl, csv
from xlcalculator import ModelCompiler, Model, Evaluator

#from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

# Prepare a usage string, to display in case of an error
usage_text = "USAGE : python3 eval2csv.py SOURCE.XLSX SrcSheetName SrcUbound:SrcLbound DEST.CSV TruncateOnFirst\n"

# Get clargs and parse
# Expect exactly five, in exactly this order:
#
# python3 eval2csv.py SOURCE.XLSX SrcSheetName SrcUbound:SrcLbound DEST.CSV
#
# Sanity checks:
# 1) Ensure have 5x clargs
# 2) Check that SOURCE.XLSX exists

# Begin by collecting clargs
clargs = sys.argv

# Aside from the script itself (counted as a clarg), do we have exactly 5x (i.e., len(clargs)=6)?
# If not have exactly 5x clargs outside of script name, quit and announce why
if len(clargs) < 6:
    quit("\nERROR : Insufficient arguments, expect exactly five\n"+usage_text)

# If we've made it this far, have enough clargs
# Before continuing, define variables with values from clargs
source_filename = clargs[1]
source_sheet_name = clargs[2]
source_range = clargs[3]
dest_filename = clargs[4]
truncate_on_first = clargs[5].lower()

# Now, test if source XLSX file exists (quit with announcement if not)
if not os.path.exists(source_filename): quit("\nERROR : source XLSX file "+source_filename+" does not exist\n"+usage_text)

# Now, test if truncation flag is boolean
if not truncate_on_first in ["true","false"]: quit("\nERROR : truncate-on-empty-first-value flag is not boolean\n"+usage_text)

# If we've made it this far, source file exists
# Ready to copy cells from source to dest
# Begin by opening source (XLSX)
source_book = openpyxl.load_workbook(source_filename, data_only=True)

# Test if appropriate sheets exist in source & dest XLSX files (quit with announcement if not)
# If yes, set source & dest sheets
try: source_sheet = source_book[source_sheet_name]
except: quit("\nERROR : sheet '"+source_sheet_name+"' not found in source XLSX file "+source_filename+"\n"+usage_text)

# Generate coordinates of all cells to evaluate in source
source_coordinates = []
for row in source_sheet[source_range.split(":")[0]:source_range.split(":")[1]]:
    this_row = []
    for this_cell in row:
        this_row.append(this_cell.coordinate)
    source_coordinates.append(this_row)
source_book.close

# Prepare to evaluate source cells
compiler = ModelCompiler()
new_model = compiler.read_and_parse_archive(source_filename)
evaluator = Evaluator(new_model)

# Next, prepare CSV output
output_stream = open(dest_filename, "w+")
writer = csv.writer(output_stream)

# Pass through each row of source coordinates
for row in source_coordinates:
    output_row = []

    # Evaluate values of cells in this row & collect for output
    for this_coordinate in row:
        this_cell_index = source_sheet_name+"!"+this_coordinate
        output_row.append(evaluator.evaluate(this_cell_index).value)
    
    # If not truncating on empty first, just output the row
    if truncate_on_first == "false": writer.writerow(output_row)
    else:
        
        # If truncating on empty first, check that first element of row is not None before outputting
        if output_row[0] != None: writer.writerow(output_row)

# Clean up
output_stream.close