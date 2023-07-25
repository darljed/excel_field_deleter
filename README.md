# excel_field_deleter
The excel_field_deleter is a Python code that removes unnecessary fields from an Excel sheet. When dealing with a large sheet containing more information than required, using this script is helpful. It allows you to select and retain only the fields you need, making the data more manageable and loading the sheet faster.

## Requirements 

- Python version 3.8 and above is required

- Install the required dependencies.

`pip install -r requirements.txt`

- create a file `fields_to_keep.txt` and list all the fields that you want to keep. 
for example: 
```
Field 1
Field 2
Field 4
```

## How to Use

The script requires one argument: 

```Argument 1 = path to target excel file```

Example: 

```
python excel_field_deleter.py sample_file.xlsx
python excel_field_deleter.py "My Sample File.xlsx"
```

## Result

The output will be an new excel file with filename `output-{current date and time}.xlsx` and with fields deleted except for the fields listed in `fields_to_keep.txt`

