# Excel-Extractor

Extracts from a given .xls or .xlsx files rows and colums.

The purpose of this software is to create csv files in order to save copy/pasts jobs.

Similar project: <https://pypi.org/project/excelextract/> This wasn't fexible enough for our needs.


## Usage

You need an excel file.

Create a Json config file e.g. `config.json`.


### Example

```json
{
  "source": "tests/pattern.xlsx",
  "sheet": "Sheet1",
  "headers": [
    { "static": "Year" },
    { "range": "A1:C1" },
    { "static": "Middle" },
    { "range": "E1:H1" },
    { "fixed": "F1:F1" }
  ],
  "data": [
    { "static": "2005" },
    { "static": "2005" },
    { "fixed": "C1:C1" },
    { "range": "H14:M19" },
    { "static": "2005" }
  ]
}
```

**source**: mandatory source file
**sheet**: mandatory sheet name in the excel file
**headers**: mandatory header section
  - **static**: a static text
  - **fixed**: content of a single cell in the excel file
  - **range**: content of a single row in the excel file used a sheader
  - you can combine as many entries here
**data**
  - **static**: a static text, the data will be repeated for all rows
  - **fixed**: content of a single cell in the excel file, the data will be repeated for all rows
  - **range**: content of a block in the excel file. The content will be copied and gets new x/y pos
  - you can combine as many entries here


Constraints:

- The sum of the entries in the data section must be equal to the header. E.g. if you have 4 static entries in the header, you can have one fixed entry in the data and a range with 3 colums.
- If you use multiple ranges in a block, the number of rows must be equal within all blocks. However the excel file might have no data.
