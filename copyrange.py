#!/usr/bin/env/python3
#
# copyrange.py
# Copy a range of cells from a specified sheet of an XLSX source file
# and insert into a same-sized range of cells in a specified sheet of 
# an XLSX destination file
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
# 1) Accept 6x command line arguments (clargs):
#       i)   XLSX source file
#       ii)  Source sheet from XLSX input file
#       iii) Source range of cells from source sheet of XLSX source file
#       iv)  XLSX dest file
#       v)   Dest sheet for XLSX dest file
#       vi)  Dest range of cells for dest sheet of XLSX dest file
# 2) Verify XLSX source and dest files exist
# 3) Verify source sheet exists in XLSX source file 
# 4) Verify dest sheet exists in XLSX dest file
# 5) Copy values from source range of source sheet in XLSX source, into
#    dest range of dest sheet in XLSX dest
#

# Imports
import sys, os, openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

# Prepare a usage string, to display in case of an error
usage_text = "USAGE : python3 copyrange.py SOURCE.XLSX SrcSheetName SrcUbound:SrcLbound DEST.XLSX DestSheetName DestUbound:DestLbound\n"

# Get clargs and parse
# Expect exactly six, in exactly this order:
#
# python3 copyrange.py SOURCE.XLSX SourceSheetName SUboundCell:SLboundCell DEST.XLSX DestSheetName DUboundCell:DLboundCell
#
# Sanity checks:
# 1) Ensure have 6x clargs
# 2) Check that SOURCE.XLSX exists
# 3) Check that DEST.XLSX exists

# Begin by collecting clargs
clargs = sys.argv

# Aside from the script itself (counted as a clarg), do we have exactly 6x (i.e., len(clargs)=7)?
# If not have exactly 6x clargs outside of script name, quit and announce why
if len(clargs) < 7:
    quit("\nERROR : Insufficient arguments, expect exactly six\n"+usage_text)

# If we've made it this far, have enough clargs
# Before continuing, define variables with values from clargs
source_filename = clargs[1]
source_sheet_name = clargs[2]
source_range = clargs[3]
dest_filename = clargs[4]
dest_sheet_name = clargs[5]
dest_range = clargs[6]

# Now, test if source and dest XLSX files exists (quit with announcement if not)
if not os.path.exists(source_filename): quit("\nERROR : source XLSX file "+source_filename+" does not exist\n"+usage_text)
if not os.path.exists(dest_filename): quit("\nERROR : dest XLSX file "+dest_filename+" does not exist\n"+usage_text)

# If we've made it this far, source & dest files exist
# Ready to copy cells from source to dest
# Begin by opening source & dest
source_book = openpyxl.load_workbook(source_filename, data_only=True)
dest_book = openpyxl.load_workbook(dest_filename)

# Test if appropriate sheets exist in source & dest XLSX files (quit with announcement if not)
# If yes, set source & dest sheets
try: source_sheet = source_book[source_sheet_name]
except: quit("\nERROR : sheet '"+source_sheet_name+"' not found in source XLSX file "+source_filename+"\n"+usage_text)
try: dest_sheet = dest_book[dest_sheet_name]
except: quit("\nERROR : sheet '"+dest_sheet_name+"' not found in source XLSX file "+dest_filename+"\n"+usage_text)

# Determine source and dest coordinates
source_coordinates = []
dest_coordinates = []
for this_cell in source_sheet[source_range.split(":")[0]:source_range.split(":")[1]]: source_coordinates.append(this_cell[0].coordinate)
for this_cell in dest_sheet[dest_range.split(":")[0]:dest_range.split(":")[1]]: dest_coordinates.append(this_cell[0].coordinate)

# Check if ranges are of the same size (quit with announcement if not)
if not len(source_coordinates) == len(dest_coordinates): quit("\nERROR : source & dest ranges differ in size\n"+usage_text)

# Set values of dest sheet
for index in range(0, len(source_coordinates)):
    dest_sheet[dest_coordinates[index]].value = source_sheet[source_coordinates[index]].value

# Clean up
dest_book.save(dest_filename)