# Filling RTF files by data from a CSV table

## General description

This repository contains one script `fill_rtf_templates.py` that allows filling RTF files using data from a CSV table.

To run the script place in this folder one .csv file in utf-8 encoding  
(typically the file is received from Google Sheets using Files -> Export -> Export as CSV)

Also the script searches in its folder one RTF file which name starts with underscore `_` -- it is the template.

The first line in the CSV file is the header -- it contains the names of the placeholders in the template that
will be substituted by values from the csv table.

Note that during the substitution the values from CSV file are converted from utf-8 encoding to RTF.

## Details on conversion utf-8 to RTF

The script `fill_rtf_templates.py` uses the files `alphabet_for_csv.txt` and `alphabet_for_rtf_ch2.rtf` to match
the corresponding symbols from utf-8 with the RTF encodings
(the file `alphabet_for_rtf_ch2.rtf` was edited manually in a binary files editor to have a form suitable for parsing).

Note that the script is intended for work with the Russian language, 
so I filled the alphabet files for cyrillic letters and several special chars.

Also note that the placeholderds should be written in the same language as the values in CSV files  
(it was a simple and rude approach to resolve the issue with RTF context)
