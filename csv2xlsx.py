#!/usr/bin/env/python3
#
# csv2xlsx.py
# Copy values from a CSV source file and insert into the single sheet
# of a new XLSX destination file
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

# Algorithm
#
# 1) Accept 2x command line arguments (clargs):
#       i)   CSV source file, assumed to exist
#       ii)  XLSX dest file, to be newly created
#

# Imports
import sys, os, csv
import pandas as pd

# Prepare a usage string, to display in case of an error
usage_text = "USAGE : python3 csv2xlsx.py SOURCE.CSV DEST.XLSX\n"

# Get clargs and parse
# Expect exactly two, in exactly this order:
#
# python3 csv2xlsx.py SOURCE.CSV DEST.XLSX
#
# Sanity checks:
# 1) Ensure have 2x clargs
# 2) Check that SOURCE.CSV exists

# Begin by collecting clargs
clargs = sys.argv

# Aside from the script itself (counted as a clarg), do we have exactly 2x (i.e., len(clargs)=3)?
# If not have exactly 2x clargs outside of script name, quit and announce why
if len(clargs) < 3:
    quit("\nERROR : Insufficient arguments, expect exactly two\n"+usage_text)

# If we've made it this far, have enough clargs
# Before continuing, define variables with values from clargs
source_filename = clargs[1]
dest_filename = clargs[2]

# Now, test if source CSV file exists (quit with announcement if not)
if not os.path.exists(source_filename): quit("\nERROR : source CSV file "+source_filename+" does not exist\n"+usage_text)

# If we've made it this far, source file exists
# Ready to copy cells from source and write to dest
# Perform algorithm
read_file = pd.read_csv(source_filename)
read_file.to_excel(dest_filename, index=None, header=True)