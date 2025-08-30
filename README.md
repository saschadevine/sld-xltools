# sld-xltools
# SLD 2025

## License
```
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
```

## Purpose
A suite of Python scripts using ```openpyxl``` for manipulating XLSX and CSV files:

- ```copyrange.py``` copies a specified range of source cells, from a specified sheet of an source XLSX file, into a specified, same-size range of destination cells, in a specified sheet of a destination XLSX file
- ```eval2csv.py``` evaluates all formulae in a range of cells from an input XLSX file, and outputs as a flat CSV file

## Example Applications
- From multiple input XLSX files associated with a single experiment, insert important statistics into an analysis template, evaluate everything in the report sheet of the template, and output as a flat CSV for distribution to teammates

## Requirements
```copyrange.py``` requires:
- Python 3.10.12 or better
- Python library ```openpyxl```

```eval2csv.py``` requires:
- Python 3.10.12 or better
- Python library ```openpyxl```
- Python library ```xlcalculator```

## Usage
### ```copyrange.py```

#### Syntax
```
python3 copyrange.py SOURCE.XLSX SrcSheetName SrcUbound:SrcLbound DEST.XLSX DestSheetName DestUbound:DestLbound
```

#### Assumptions

- ```SOURCE.XLSX``` is a valid XLSX file containing a sheet with source data
- ```SrcSheetName``` is a valid sheet in ```SOURCE.XLSX```
- ```SrcUbound:SrcLbound``` is a range of cells in ```SrcSheetName``` and contains the source data (takes the form ```A1:Z99```)
- ```DEST.XLSX``` is a valid XLSX file containing a sheet to be used as data destination
- ```DestSheetName``` is a valid sheet in ```DEST.XLSX```
- ```DestUbound:DestLbound``` is a range of cells in ```DestSheetName``` and is where source data will be inserted (takes the form ```A1:Z99```)

### ```eval2csv.py```

#### Syntax
```
python3 eval2csv.py SOURCE.XLSX SrcSheetName SrcUbound:SrcLbound DEST.CSV TruncateOnFirst
```

#### Assumptions

- ```SOURCE.XLSX``` is a valid XLSX file containing a sheet with source data
- ```SrcSheetName``` is a valid sheet in ```SOURCE.XLSX```
- ```SrcUbound:SrcLbound``` is a range of cells in ```SrcSheetName``` and contains the source data (takes the form ```A1:Z99```)
- ```DEST.CSV``` is a valid CSV filename to serve as destination for data (does not need to exist; existing files will be overwritten)
- ```TruncateOnFirst``` is a boolean flag:
  - If ```TRUE```, only rows with values in column A are evaluated and written to ```DEST.CSV```. Rows containing no value in column A are treated as containing no data (even if formulae are present) and are not written to ```DEST.CSV```
  - If ```FALSE```, all rows will be evaulated written to ```DEST.CSV```, even for rows where column A contains no value
