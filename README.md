# XLWT Examples from PythonExcels

This repository contains examples of Python scripts that generate Excel spreadsheets from other raw data. The scripts have been tested using Python versions 3.7.3 and 2.7.11.

## xlwt_hospdata.py

This script reads some raw text containing account information for hospital, parses the account number and account name information, then writes the result to an excel spreadsheet using the [XLWT][python-excel] module. A complete description of this script can be found at [pythonexcels.com](https://pythonexcels.com/python/2009/09/19/another-xlwt-example).

## xlwt_bostonhousing.py

This script opens a URL to a tab delimited file containing location and pricing
data on the Boston housing market. The data for this example is saved in this
repository as boston_corrected.txt and comes from research done by David
Harrison and Daniel L. Rubinfeld in "Hedonic Housing Prices and the Demand for
Clean Air", published in the Journal of Environmental Economics and Management,
Volume 5, (1978). The data is read, parsed and written to a spreadsheet using
the XLWT module. a reads some raw text containing account information for
hospital, parses the account number and account name information, then writes
the result to an excel spreadsheet using the [XLWT][python-excel] module. A
complete description of this script can be found at http://pythonexcels.github.io/2009_09_19_Another_XLWT_Example.html .

## boston_corrected.txt

This is a copy of the data read by the xlwt_bostonhousing.py script.

## hospdata.txt

This is a copy of the data file read by the xlwt_hospdata.py script.

[pythonexcels]: http://www.pythonexcels.com
[python-excel]: http://www.python-excel.org
