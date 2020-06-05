# Camtasia_and_Audiate_tools
Tools for use with TechSmith's Camtasia and Audiate programs. These tools
are intended to be examples of how one can work with SRT and Camtasia project files.

Programs can be called with the option "-v" or "--verbose" you get lots of output - showing in detail the operations of the program.

## srt2spreadsheet.py

Purpose: To convert a SRT file to a XLSX spreadsheet

Input: 
```
./srt2spreadsheet.py filename
```

Output: outputs an XLSX spreadsheet with the contents of the SRT file

Example (edited to show only some of the output):
```
./srt2spreadsheet.py ethical-research-edited.srt
```

## spreadsheet2srt.py 

Purpose: To convert a XLSX spreadsheet ro a SRT file

Input: 
```
./spreadsheet2srt.py filename
```

Output: converts the spreadsheet to an SRT file

Example (edited to show only some of the output):
```
./spreadsheet2srt.py captions-ethical-research-edited.srt.xlsx
```

## parse_camproj_file.py

Purpose: To get data from a Camtasia project file and save it into a JSON formatted file

Input: 
```
./parse_camproj_file.py filename
```

Output: parses the Camtasia project file and outputs some of the information about keyframes in a JSON file

Example (edited to show only some of the output):
```
./parse_camproj_file.py  II2202-ethical-research.camproj 
```
