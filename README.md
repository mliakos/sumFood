# sumFood

This simple, single-purpose Python script sums CSVs (comma separated values) in an excel file and groups them by their relative keys (FOO: 5, BAR: 10, FOOBAR: 12, etc.). 


Only files, sheets and columns that meet the following requirements are considered in the sum process:

- Files with '.xls' or '.xlsx' extension.
- Files with a title that contain the keyword 'ΤΡΟΦΙΜΑ'.
- Sheets with a title that contain the keyword 'ΠΑΚΕΤΟ'.
- Columns D and/or E. 

## Example

In every sheet, each row contains a cell with food in the following format:

5 bread, 1L olive oil, 2 packs crackers, etc...

The script loops through every row in every relevant sheet and sums the food resulting in the food key with its aggregated value (i.e. BREAD: 72, OLIVE OIL: 32L, CRACKERS: 15 PACKS).

Consider this example sheet as part of a multi-sheet workbook:

![Screenshot_1](https://user-images.githubusercontent.com/34876695/67626602-108c0b00-f856-11e9-8656-fd363ccf417b.png)

This is the result after summing all sheets in the workbook:

![Screenshot_2](https://user-images.githubusercontent.com/34876695/67626647-ceaf9480-f856-11e9-9d52-0fcd704fd14c.png)


## Dependencies

- openpyxl


```bash
pip install openpyxl
```

## Usage

- Open a command prompt in the directory containing both the script and the provided excel file.
- Run the script using the Python interpreter.
- Typing yes in the first question will sum food in column D (food). Typing no will omit this column.
- Typing yes in the second question will sum food in column E (extra food). Typing no will omit this column.

## License
[MIT](https://choosealicense.com/licenses/mit/)
