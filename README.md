# Working with CSV and XLSX files in Python 

<a href="http://creativecommons.org/licenses/by-nc/4.0/" rel="license"><img style="border-width: 0;" src="https://i.creativecommons.org/l/by-nc/4.0/88x31.png" alt="Creative Commons License" /></a>
This tutorial is licensed under a <a href="http://creativecommons.org/licenses/by-nc/4.0/" rel="license">Creative Commons Attribution-NonCommercial 4.0 International License</a>.

## Lab Goals

This lab covers various methods for reading structured data into Python using `csv` and `openpyxl`. It also covers other types of delimiters and escape characters. The lab also overs writing data from Python to a structured data file.

By the end of this lab, students will be able to:
- Describe the structure and components of a `.csv` file
- Read a `.csv` file into Python using the `csv` module and a `for` loop
- Understand how to approach other types of delimited files
- Understand how to work with escape characters when loading structured data
- Understand how Python dictionaries work as a type of structured data
- Write data from Python to a `.csv` file
- Understand the basic of working with `.xlsx` files in Python

## Acknowledgements

Information and exercises in this lab are adapted from Al Sweigart's [*Automate the Boring Stuff With Python*](https://nostarch.com/automatestuff2) (No Starch Press, 2020).
- Chapter 13, "Working With Excel Spreadsheets" (302-328)
- Chapter 16 "Working With CSV Files and JSON Data" (371-388)

# Table of Contents
- [Data](#data)
- [`.csv` data in Python](#csv-data-in-python)
  * [What is a `.csv` file?](#what-is-a-csv-file)
  * [Reading a `.csv` file into Python](#reading-a-csv-file-into-Python)
    * [Reading `.csv` data using a `for` loop](#reading-csv-data-using-a-for-loop)
  * [Other delimiters](#other-delimiters)
  * [Escape characters](#escape-characters)
- [Reading in `.csv` files using dictionaries](#reading-in-csv-files-using-dictionaries)
- [Writing to a `.csv` file](#writing-to-a-csv-file)
  * [Writing from a dictionary to a `.csv` file](#writing-from-a-dictionary-to-a-csv-file)
- [Working with `.xlsx` files](#working-with-xlsx-files)
- [Project Prompts](#project-prompts)
  * [Project #1](#project-1)
  * [Project #2](#project-2)
  * [Project #3](#project-3)
- [Lab notebook questions](#lab-notebook-questions)

# Data

You'll need four data files for this lab.
- `example.csv`
- `example.txt`
- `example.xlsx`
- `exampleWithHeader.csv`

They can all be downloaded from this GitHub repository as individual files or a zip folder.

You can also access them [via Google Drive](https://drive.google.com/drive/folders/1Sp_N34753ONJRU2AFKcocQ2DhCEhyL-m?usp=sharing) (ND users only).

[Click here](https://raw.githubusercontent.com/kwaldenphd/tabular-data-python/main/tabular_data_python.ipynb) and select the "Save As" option to download this lab as a Jupyter Notebook.

# `.csv` data in Python

## What is a `.csv` file? 

1. CSV stands for "comma-separated values." CSV files are tabular data structures (i.e. a spreadsheet), stored in a plain-text format.

2. Python includes a built-in `csv` module that allows us to read in data from a `.csv` file (or other type of delimited plain-text file), as well as write data to a `.csv` file.

3. In `.csv` files, each line represents a row in the spreadsheet, and commas separate cells in each row (thus the file format name comma-separated values).

4. At first glance, `.csv` files look similar to proprietary spreadsheet program file formats, such as files saved in Microsoft Excel or Apple Numbers. 

5. However, file formats like `.xls` or `.xlsx` (Microsoft Excel) or `.numbers` (Apple Numbers) are NOT plain-text formats. Try to open these file types in a text editor and you'll quickly see the additional content and markup added by the spreadsheet program. More on this later.

6. A few characteristics that distinguish `.csv` files (or other plain-text structured data formats) from proprietary spreadsheet file types:
- Columns in a `.csv` file don't have a value type. Everything is a string.
- Values in a `.csv` file don't have font or color formatting
- `.csv` files only contain single worksheets
- `.csv` files don't store formatting information like cell width/height
- `.csv` files don't recognize merged cells or other kinds of special formatting (frozen or hidden rows/columns, embedded images, etc.)

7. Given these limitations, especially compared to the way we often interact with spreadsheet programs like Microsoft Excel or Google Sheets, what's the advantage of working with `.csv` files?

8. One key advantage of `.csv` files is their simplicity. You can load or open a `.csv` file in a text editor and be able to quickly see the values in the file. 

9. When working with data in a programming environment, `.csv` files as a plain-text format simplify the process of loading structured data.

<blockquote>Q1: Open the <code>example.xlsx</code> file in a text editor. Describe what you see.</blockquote>

<blockquote>Q2: How does your answer to Q1 compare to what you see when you open the <code>example.csv</code> file in a text editor?</blockquote>

<blockquote>Q3: Open the <code>example.xlsx</code> file in a spreadsheet program. Save the file as a <code>.csv</code> format. What happens? Or what happens when you open the newly-created <code>.csv</code> file in a spreadsheet program or text editor?</blockquote>

## Reading a `.csv` file into Python

10. To read data from a `.csv` file into Python, we will use the `csv` module.

<blockquote>Check out <a href = "https://docs.python.org/3/library/csv.html#module-csv">Python's documentation</a> to learn more about the <code>csv</code> module.</blockquote>

11. The `csv` module allows us to create a `reader` object that iterates over lines in a `.csv` file.

12. What does this workflow look like? 
```python
# import csv module
import csv

# open csv file
exampleFile = open('data/example.csv')
# NOTE: If you are on Windows, you may have to change 'data/example.csv' to 'data\example.csv'

# create reader object from lines of data in example.csv file using csv.reader function
exampleReader = csv.reader(exampleFile)

# create list with rows of data from example.csv file
exampleData = list(exampleReader)

# output list rows
exampleData
```

13. You'll notice that the `exampleData` output is a list of lists, or a list that contains sub-lists. 

14. Each row of data from the original `example.csv` file is a sub-list (with field values separated by commas) within the `exampleData` list.

15. Now we can access the value at a particular row and column using the expression `exampleData[row][col]`, where `row` is the index position of one of the lists in `exampleData`, and `col` is the index position of the item located in that list.

16. For example, `exampleData[0][0]` would give us the first string from the first list. `exampleData[0][1]` would give us the second string from the first list.

<blockquote>Q4: Create a list of sublists and use the index positions to access specific values.</blockquote>

<blockquote>Q5: Read in small CSV and access specific values. Include code + comments.</blockquote>

## Reading `.csv` data using a `for` loop

17. The method we just used to read data from a `.csv` file into Python loads the entire file into memory at once.

18. If we use this method on a large `.csv` file, Python is going to try to load the entire file into memory at once. This does not bode well for Python or your computer's performance.

19. We can use a `reader` object as part of a `for` loop to iterate through the lines in a `.csv` file and load the file line-by-line.

<blockquote>Remember <code>for</code> loops iterate through each item in a series or list of items and performs the content of the loop on each item.</blockquote>

20. What does this workflow look like?
```python
# import csv module
import csv

# open .csv file
exampleFile = open('data/example.csv')
# NOTE: If you are on Windows, you may have to change 'data/example.csv' to 'data\example.csv'

# create reader object from .csv file
exampleReader = csv.reader(exampleFile)

# iterate through each row in .csv file and print out row content and number
for row in exampleReader:
  print('Row #' + str(exampleReader.line_num) + ' ' + str(row))
```

21. This program imports the `csv` module, makes a `reader` object from the `example.csv` file, and loops through each of the rows in the `reader` object.

22. Each row is a list of values, and each value represents a cell.

23. The `print()` function prints the current row number and that row's contents. 

24. The `reader` object includes a `line_num` variable, which contains the number of the current line.

25. NOTE: The `reader` object can only be looped over once. If you need to re-read the same `.csv` file, you'll use `csv.reader` to create a new `reader` object.

<blockquote>Q6: Read in a small CSV file using a <code>for</code> loop and access specific values. Include code + comments.</blockquote>

## Other delimiters

26. But what happens if you need to load in structured data that uses another delimiter, not a comma? 

27. Remember when we opened a `.csv` file in a plain-text editor, the value fields are separated by a comma.

28. But commas are not the only possible delimiter. Tabs, spaces, pipes, or other characters can be used to separate or delimit fields in a dataset.

29. The `csv` module includes a range of formatting parameters, known as a `Dialect` class. 

30. The `Dialect` class includes a range of methods you can use to specify alternate delimiters and (as we'll discover shortly), handle situations like special characters, line breaks, etc.

31. The `delimiter` attribute in the `Dialect` class lets us specify what delimiter is being used in the data we want to load.
```Python
# import csv module
import csv

# load tab-separated value file
tsv_file = open('data/example.txt')
# NOTE: If you are on Windows, you may have to change 'data/example.txt' to 'data\example.txt'

# create a reader object and specify the new delimiter
read_tsv = csv.reader(tsv_file, delimiter="\t")

# use a for loop to read in the data
for row in read_tsv:
  print(row)
```

<blockquote>Q7: Modify code to load in example.txt file.</blockquote>

## Escape characters

32. But what happens if the values in your dataset include the same character that's being used as a delimiter?

33. For example, let's say you have address data in the following structure:

Name | Age | Address
--- | --- | ---
Jerry | 10 | 2776 McDowell Street, Nashville, Tennessee
Tom | 20 | 3171 Jessie Street, Westerville, Ohio
Mike | 30 | 1818 Sherman Street, Hope, Kansas

34. In this example, we want to keep `Address` as an intact field and not separate based on the commas located within the address.

35. In order to do this, we need to specify how Python parses fields that include the delimiter character.

36. The `quotechar` attribute in the `Dialect` class specifies what character will be used to enclose fields that should be treated as distinct entities and not be split into columns or fields based on the presence of the delimiter character within the field.

37. The default for `quotechar` is `"` (double quotation marks).

38. So what does that mean? We put double quotation marks around the field that includes the delimiter character.

39. Modified data structure:

Name | Age | Address
--- | --- | ---
Jerry | 10 | "2776 McDowell Street, Nashville, Tennessee"
Tom | 20 | "3171 Jessie Street, Westerville, Ohio"
Mike | 30 | "1818 Sherman Street, Hope, Kansas"

40. Then we can read the modified data into Python.

41. But what happens if we have quotation marks within a field that needs to be treated as a distinct entity?

42. For example, the following data structure would run into problems when read into Python.

Id | User | Comment
--- | --- | ---
1 | Bob | "John said "Hello World""
2 | Tom | ""The Magician""
3 | Harry | ""walk around the corner" she explained to the child"
4 | Louis | "He said, "stop pulling the dog's tail""

43. See our problem? The `Comment` field is enclosed with double quotation marks but also includes quotation marks in the field. 

44. We need Python to understand the enclosing double quotation marks serve a different purpose than the double quotation marks contained within the `Comment` field.

45. We can use a blackslash `\` character to escape the embedded double quotes.

46. Modified data structure:

Id | User | Comment
--- | --- | ---
1 | Bob | "John said \"Hello World\""
2 | Tom | "\"The Magician\""
3 | Harry | "\"walk around the corner\" she explained to the child"
4 | Louis | "He said, \"stop pulling the dog's tail\""

47. Since the default for `quotechar` is `"`, we need to modify that default to reflect the new data structure.

Create a file called `data.csv` to match the table above before running the code below.
```Python
# import csv module
import csv

# read csv using quote quotechar
with open('data.csv', 'rt') as f:
  csv_reader = csv.reader(f, skipinitialspace=True, quotechar='\\')
  
  for line in csv_reader:
    print(line)
```

<blockquote>Q8: Define escape characters in your own words. Describe a situation in which escape characters would be needed, and how you would address that challenge using Python syntax.</blockquote>
  
# Reading in `.csv` files using dictionaries

48. For `.csv` files that contain header rows, we might want to connect the header row values with subsequent row values.

49. We can do this by reading the `.csv` file as a dictionary, rather than a list containing row sub-lists.

50. Remember dictionaries have key-value pairs, where we can access a value by using its key name.

51. For tabular data, we can think of the key as the field name contained in the header row and the value as column or field values.

52. We read a `.csv` file to a dictionary using a `DictReader` object (versus the `csv.reader` object).
```Python
# import csv module
import csv

# open csv file
exampleFile = open('data/exampleWithHeader.csv')
# NOTE: If you are on Windows, you may have to change '/' to '\'

# reading exampleFile to a dictionary
exampleDictReader = csv.DictReader(exampleFile)

# set keys for key-value pairs
for row in exampleDictReader:
  print(row['Timestamp'], row['Fruit'], row['Quantity'])
```

53. Within the `for` loop, the `DictReader` object sets `row` to a dictionary object with keys derived from the headers in the first row.

54. The `DictReader` object means we don't have to separate the header information from the rest of the data contained in the file, because the `DictReader` object does this for us.

<blockquote>Q9: Load a small CSV file that includes header info. Include code + comments.</blockquote>

55. But what can we do if we want to read to a dictionary a `.csv` file that doesn't incldue a header row?

56. We can pass a second argument to the `DictReader()` function to manually set header names.
```Python
# import csv module
import csv

# open csv file
exampleFile = open('data/example.csv')
# NOTE: modify the forward slash to a backslash if you are on Windows.

# reading exampleFile to a dictionary with added field names
exampleDictReader = csv.DictReader(exampleFile, ['time', 'name', 'amount'])

# set keys for key-value pairs
for row in exampleDictReader:
  print(row['time'], row['name'], row['amount'])
```

<blockquote>Q10: Load a small CSV that does include headers and manually set the headers. Include code + comments.</blockquote>

# Writing to a `.csv` file

57. We'll do more with writing to a `.csv` file later in the semester.

58. But for now, we can create a `writer` object using the `csv.writer()` function to write data to a `.csv` file.
```Python
# import csv module
import csv

# create and open output.csv file in write mode
outputFile = open('output.csv', 'w', newline='')

# create writer object
outputWriter = csv.writer(outputFile)

# write first row
outputWriter.writerow(['spam', 'eggs', 'bacon', 'ham'])

# write another row
outputWriter.writerow(['Hello, world!', 'eggs', 'bacon', 'ham'])

# write a third row
outputWriter.writerow([1, 2, 3.141592, 4])

# close the output file
outputFile.close()
```

59. The `writerow()` method takes a list argument and writes that to a new row in the `writer` object, that is added to the `.csv` file.

<blockquote>Q11: Create your own small data structure and write it to CSV file. Describe what you expect to see and what actually happens when you look at the output. Include code + comments.</blockquote>

## Writing from a dictionary to a `.csv` file

60. We can use the `DictWriter` object to write data in a dictionary to a `.csv` file.
```Python
# import csv module
import csv

# create and open output.csv file in write mode
outputFile = open('output.csv', 'w', newline='')

# create writer object
outputDictWriter = csv.DictWriter(outputFile, fieldnames=['Name', 'Pet', 'Phone'])

# create header row
outputDictWriter.writeheader()

# write first row
outputDictWriter.writerow({'Name': 'Alice', 'Pet': 'cat', 'Phone': '555-1234'})

# write another row
outputDictWriter.writerow({'Name': 'Bob', 'Phone': '555-9999'})

# write a third row
outputDictWriter.writerow({'Phone': '555-5555', 'Name': 'Carol', 'Pet': 'dog'})

# close the output file
outputFile.close()
```

61. Note that the order of the key-value pairs in the dictionaries created manually using `writerow()` doesn't matter.

62. Python writes the dictionaries to the `.csv` file using the order of the keys given to `DictWriter()`.

63. Missing keys will be empty in the newly-created `.csv` file.

<blockquote>Q12: Create a small dictionary and write it to a CSV file. Include code + comments.</blockquote>

# Working with `.xlsx` files

64. As noted earlier, file formats like `.xls` or `.xlsx` (Microsoft Excel) or `.numbers` (Apple Numbers) are NOT plain-text formats. 

65. Try to open these file types in a text editor and you'll quickly see the additional content and markup added by the spreadsheet program. More on this later.

66. A few characteristics that distinguish Excel workbooks from `.csv` files (or other plain-text structured data formats):
- Columns in an Excel workbook do have a value type, either determined by Excel or set manually.
- Values in an Excel workbook can include font or color formatting
- Excel workbooks can contain multiple sheets
- Excel workbooks store formatting information like cell width/height
- Excel files can include merged cells or other kinds of special formatting (frozen or hidden rows/columns, embedded images, formulas, etc.)

67. In most cases, the most reliable option is to save each sheet of an Excel workbook to an individual `.csv` file which you can then read into Python.

68. But if you need to load an Excel workbook directly into Python, the `openpyxl` module can let us read and modify Excel files.

<blockquote><a href="https://openpyxl.readthedocs.io/en/stable/">Click here</a> to learn more about the openpyxl module.</blockquote>

69. First thing we need to do is install the `openpyxl` module.
```Python
pip install openpyxl
```

70. We can load the `example.xlsx` workbook into Python using `openpyxl`.
```Python
# import openpyxl
import openpyxl

# load workbook
wb = openpyxl.load_workbook('data/example.xlsx')
#NOTE: modify the forward slash to a backslash if you are on Windows.
```

71. Note that if `example.xlsx` is not in the current working directory, otherwise you will want to provide the full file path.

72. We can use the `sheetnames` attribute to see what sheets are in the workbook.
```Python
# import openpyxl
import openpyxl

# load workbook
wb = openpyxl.load_workbook('data/example.xlsx')
#NOTE: modify the forward slash to a backslash if you are on Windows.

# get workbook sheet names
wb.sheetnames

# isolate a specific sheet from the workbook using its name
sheet = wb['Sheet3']
```

73. To learn more about working with Excel workbooks in Python:
- Al Sweigart, "Chapter 13, Working With Excel Spreadsheets," in *Automate the Boring Stuff With Python* (No Starch Press, 2020): 302-328.
- [`openpyxl` documentation](https://openpyxl.readthedocs.io/en/stable/)

<blockquote>Q13: How would you compare these two formats? When would you want one format vs the other?</blockquote>

# Project Prompts

## Project #1:

Create an automated program that removes first row of files for everything in `removeCsvheader.zip` folder. 

Include code + comments.

Describe a scenario in which you would want or need to do this.

## Project #2:

Navigate to an open data portal and download a `.csv` or `.xlsx` file. OpenData.gov, city of Chicago, SB, SportsReference, other. 

A few places to start:
- [Data.gov](https://www.data.gov/)
- [City of Chicago Data Portal](https://data.cityofchicago.org/)
- [City of South Bend Open Data](https://data-southbend.opendata.arcgis.com/)

Open the data in a spreadsheet program and/or text editor
- What do you see?
- How can we start to make sense of the data based on available documentation?

Load the data in Python as list/sublists and as dictionary. What challenges did you encounter? How did you address or solve them? 

### Project #3

Manually create a small dataset and write to a CSV file.

# Lab Notebook Questions

Q1: Open the `example.xlsx` file in a text editor. Describe what you see.

Q2: How does your answer to Q1 compare to what you see when you open the `example.csv` file in a text editor?

Q3: Open the `example.xlsx` file in a spreadsheet program. Save the file as a `.csv` format. What happens? Or what happens when you open the newly-created `.csv` file in a spreadsheet program or text editor?

Q4: Create a list of sublists and use the index positions to access specific values. Include code + comments.

Q5: Read in a small CSV file and access specific values. Include code + comments.

Q6: Read in a small CSV file using a `for` loop and access specific values. Include code + comments.

Q8: Modify code to load in `example.txt` file.

Q8: Define escape characters in your own words. Describe a situation in which escape characters would be needed, and how you would address that challenge using Python syntax.

Q9: Load a small CSV file that includes header info. Include code + comments.

Q10: Load a small CSV that does include headers and manually set the headers. Include code + comments.

Q11: Create your own small data structure and write it to CSV file. Describe what you expect to see and what actually happens when you look at the output. Include code + comments.

Q12: Create a small dictionary and write it to a CSV file. Include code + comments.
