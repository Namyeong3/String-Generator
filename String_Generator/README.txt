==== String Generator Documentation ====





[Updated in 20:56 08/09/2023]










==== Required Packages (Install for Windows) ====

In the command line, as an admin copy:

pip install pandas
pip install pyperclip
pip install keyboard










==== Introduction ====

These are the steps to use the program in a nutshell:

1 - Enter the template and variables (the parts of the strings that will change);

2 - The program identifies the variables and data arrays are created for each of them;

3 - The data is imported or inputed for each variable array;

4 - Select the number of strings to generate;

5 - Done! The output is displayed and copied to the clipboard.


Here's a simple example to get things started:

'''
Do you want to add advanced features? (y/n): n
Enter the template string: Test_[index]_[data]_[odd or even]
Enter input type for 'index' (txt/excel/manual/lookup/sequential/repeatn): sequential
Setting up 'index' as a sequential operation.
Enter the starting number for 'index': 1
Enter the number of values to generate for 'index' (enter a large number for no end): 5
'index' values:
['1', '2', '3', '4', '5']
Enter input type for 'data' (txt/excel/manual/lookup/sequential/repeatn): txt
Enter path to txt file for 'data': "C:\Users\namye\OneDrive\Namyeong Galaxy Book\Desktop\String_Generator_Example_Files\txt_num.txt"
'data' values:
['131', '34', '51', '304', '66', '123', '266', '1148']
Enter input type for 'odd or even' (txt/excel/manual/lookup/sequential/repeatn): manual
Enter value(s) for 'odd or even' (comma-separated for multiple values): odd,even
'odd or even' values:
['odd', 'even']
Enter the number of strings to generate: 5
Test_1_131_odd
Test_2_34_even
Test_3_51_odd
Test_4_304_even
Test_5_66_odd
Generated strings copied to clipboard.
Press 'Escape' to close the program...
'''

This is the output:

Test_1_131_odd
Test_2_34_even
Test_3_51_odd
Test_4_304_even
Test_5_66_odd










==== Variables ====

Variables are the part of the template that changes for every string in the output. Variables are arrays and the index of the array corresponds to the index of the generated string (so the first element of an array is attributed to the first string).
If an array is smaller than the requested number of strings to generate, the array will repeat, from the start.
The program considers everything inside '[]' (square brackets) as a variable. Although this can be changed if advanced features are enabled.
Everything in the template which is not a variable will stay untouched in the output strings.

E.g:

Template: http[security]:\\[domain].example\[page]

Variables: security, domain, page

Considering:

security = ['', 's', 's', '', '']

domain = ['wikipedia', 'wikimedia', 'example', 'studyin[country]', 'stackoverflow']

page = ['home', 'images', 'test', 'schoolarships', 'nerds']

The output will be (for 5 strings):

http:\\wikipedia.example\home
https:\\wikimedia.example\images
https:\\example.example\test
http:\\studyin[country].example\schoolarships
http:\\stackoverflow.example\nerds










==== Variable Input Types ====

For each variable, the user will be prompted to choose an input type.
These are methods to create the variable arrays, they're listed below:

=== Types resumed ===

txt - imports the elements from a txt file;

excel - imports the elements from an excel column;

manual - prompts the user to type the array in the terminal;

lookup - similar to excel's vlookup, compares the values to the index array, upon matching the values, it stores the index and then retrieves the values with that index from the output array; 

sequential - creates an array from an arythmetic sequence;

repeatn - creates an array by repeating several values the chosen number of times.


=== txt ===

Imports the data from a text file and converts it into an array.


Parameters:

Separator (Advanced): Type the separator used in the text file to separate elements (type 'enter' if they're sepparated by paragraphs);

Path: The path of the txt file. Both Linux and Windows paths are supported, and it can also be inside "". Tip: in windows, click the file and then press CTR + Shift + C, this will copy the path of the selected file.

Note: You must close the document so the program can import it.



=== excel ===

Imports the data from an excel document and converts the selected column it into an array.


Parameters:

Path: The path of the Excel doc. Both Linux and Windows paths are supported, and it can also be inside "". Tip: in windows, click the file and then press CTR + Shift + C, this will copy the path of the selected file.

Sheet: The sheet name in which the data is on.

Headers: If your data is in a table with headers, insert its column name. Otherwise insert the column reference (such as 'A','AB', etc.)

Note: You must close the document so the program can import it.



=== manual ===

Creates an array based on the inputed data in the terminal, the elements will be sepparated by ',' (commas).



=== lookup ===

Works like excel lookup. It needs 3 arrays:

Lookup Values - The values you want to match in the dataset;

Index array - The dataset where the lookup values are matched;

Output array - The values which will be outputed based on the indices retrieved previously.


The calculation process is the following:

1 - The program looks for the lookup array in the index array. 

2 - When it finds a match, the index in the index array is recorded.

3 - The indices are used to retrieve the values in the output array (which correspond to the index arra).

To enter these arrays you may use the txt/excel/manual/sequential input methods.

E.g:

Lookup values: ['a','b','d','g','j']

Index array: ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']

Output array: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]

Resulting array: [1, 2, 4, 7, 10]

In this example, the input is a few letters, and the output is the position of those letters in the alphabet.



=== Sequential ===

The sequential function creates an array of numbers in an arythmetic sequence.

Decimals (Advanced): Number of decimal digits, if there's too many digits they'll be rounded, if there's too little zeros are added to the right.

Step size (Advanced): The value to sum/subtract for every element in the array.

Order (Advanced): Wether to sum or subtract each step for each index (crescent or decrescent order)

Total integer digits (Advanced): The number of digits which are not decimals. If the number of digits is inferior to the one requested, the program adds zeros on the left of the number. If the number is bigger than requested, nothing will be done.
E.g:
Input: [1, 01, 123, 1.000, 2.39]
Total integer digits = 2
Output: [01, 01, 123, 01.000, 02.39]

Starting number: The number to start the sequence.

Number of values to generate: The number of elements to create in the array.



=== repeatN ===

Repeats a certain keywords an n number of times.
To do so the array must conatain an integer number (count) followed by the string (or number).
To enter these arrays you may use the txt/excel/manual input methods.

E.g:

Input array: [1,'start',3,'data',2,1,4,'example']

Output array: ['start', 'data', 'data', 'data', '1', '1', 'example', 'example', 'example', 'example']










==== Output ====

The output strings are displayed in the terminal and copied to the clipboard.
If advanced features are enableded, you'll be prompted to save the strings in a document, either txt or excel.
If you insert the path of an existing txt file, the data in it will be overwritten.
If you insert the path of an existing excel workbook, make sure to insert a sheet name which doesn't exist yet, or it will overwrite the existing one.
If you enter the path of a folder, the program creates a new document instead.










==== Advanced Features ====

The advanced features are the first thing prompted in the terminal.
This adds some features but makes the program harder (and slower) to use.

These are the advanced features:

- Choose a custom delimiter for the variables in the template;
- Disable displaying arrays everytime a new array is imported or variable created (useful if you have a very large dataset);
- Custom sepparators in 'txt' input type;
- Decimals in 'sequential' input type;
- Step size in 'sequential' input type;
- Order in 'sequential' input type;
- Total integer digits in 'sequential' input type;
- Crop array elements: Deletes the selected number of characters in the start/end of each item in a variable array;
- Export strings to doc.










==== Extensive Example ====


Do you want to add advanced features? (y/n): y
Enter the left delimiter for variables (default is '['): {
Enter the right delimiter for variables (default is ']'): }
Disable printing of array values? (y/n): n
Enter the template string: Ext_example{s}_{country}_{country code}_{serial}_{observations}_[square brackets]
Enter input type for 's' (txt/excel/manual/lookup/sequential/repeatn): repeatn
Enter input type for repeatn values of 's' (txt/excel/manual): manual
Enter repeatn values for 's' (count followed by values, comma-separated): 1,s,100,
Do you want to crop 's'? (y/n, advanced): n
's' values:
['s', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
Enter input type for 'country' (txt/excel/manual/lookup/sequential/repeatn): txt
Enter the separator for 'country' (',' for comma, ' ' for space, 'enter' for paragraphs): enter
Enter path to txt file for 'country': "C:\Users\namye\OneDrive\Namyeong Galaxy Book\Desktop\String_Generator_Example_Files\Countries_List.txt"
Do you want to crop 'country'? (y/n, advanced): n
'country' values:
['Germany', 'United Kingdom', 'China', 'Portugal', 'Japan', 'Italy', 'Norway', 'South Korea', 'Russia', 'Japan', 'France']
Enter input type for 'country code' (txt/excel/manual/lookup/sequential/repeatn): lookup
Setting up 'country code' as a lookup operation.
Enter input type for lookup values' (txt/excel/manual/sequential): txt
Enter the separator for lookup values of 'country code' (',' for comma, ' ' for space, 'enter' for paragraphs): enter
Enter path to txt file for lookup values of 'country code': "C:\Users\namye\OneDrive\Namyeong Galaxy Book\Desktop\String_Generator_Example_Files\Countries_List.txt"
Lookup Values:
['Germany', 'United Kingdom', 'China', 'Portugal', 'Japan', 'Italy', 'Norway', 'South Korea', 'Russia', 'Japan', 'France']
Enter input type for index array (txt/excel/manual/sequential): excel
Enter path to Excel file for index array: "C:\Users\namye\OneDrive\Namyeong Galaxy Book\Desktop\String_Generator_Example_Files\Test.xlsx"
Enter Excel sheet name for index array: GeoDatabase
Does the sheet have headers? (y/n): y
Enter Excel column name for index array: Country
Index Array:
['United States', 'United Kingdom', 'Germany', 'China', 'Japan', 'South Korea', 'Russia', 'France', 'Spain', 'Portugal', 'Italy', 'Netherlands', 'Norway']
Enter input type for output array (txt/excel/manual/sequential): excel
Enter path to Excel file for output array: "C:\Users\namye\OneDrive\Namyeong Galaxy Book\Desktop\String_Generator_Example_Files\Test.xlsx"
Enter Excel sheet name for output array: GeoDatabase
Does the sheet have headers? (y/n): y
Enter Excel column name for output array: Abbreviation
Do you want to crop 'country code'? (y/n, advanced): n
'country code' values:
['DE', 'GB', 'CN', 'PT', 'JP', 'IT', 'NO', 'KR', 'RU', 'JP', 'FR']
Enter input type for 'serial' (txt/excel/manual/lookup/sequential/repeatn): sequential
Setting up 'serial' as a sequential operation.
Enter the number of decimals for 'serial' (advanced): 2
Enter the step size for 'serial' (advanced): 2.367
Crescent order for 'serial'? (y/n, advanced): n
Enter the total integer digits for 'serial': 3
Enter the starting number for 'serial': 50
Enter the number of values to generate for 'serial' (enter a large number for no end): 11
Do you want to crop 'serial'? (y/n, advanced): y
Crop 'serial' at the start or end? (start/end, advanced): end
Enter the number of digits to crop for 'serial' (advanced): 1
'serial' values:
['050.0', '047.6', '045.2', '042.9', '040.5', '038.1', '035.8', '033.4', '031.0', '028.7', '026.3']
Enter input type for 'observations' (txt/excel/manual/lookup/sequential/repeatn): repeatn
Enter input type for repeatn values of 'observations' (txt/excel/manual): manual
Enter repeatn values for 'observations' (count followed by values, comma-separated): 2,ko,4,ok
Do you want to crop 'observations'? (y/n, advanced): n
'observations' values:
['ko', 'ko', 'ok', 'ok', 'ok', 'ok']
Enter the number of strings to generate: 11
Ext_examples_Germany_DE_050.0_ko_[square brackets]
Ext_example_United Kingdom_GB_047.6_ko_[square brackets]
Ext_example_China_CN_045.2_ok_[square brackets]
Ext_example_Portugal_PT_042.9_ok_[square brackets]
Ext_example_Japan_JP_040.5_ok_[square brackets]
Ext_example_Italy_IT_038.1_ok_[square brackets]
Ext_example_Norway_NO_035.8_ko_[square brackets]
Ext_example_South Korea_KR_033.4_ko_[square brackets]
Ext_example_Russia_RU_031.0_ok_[square brackets]
Ext_example_Japan_JP_028.7_ok_[square brackets]
Ext_example_France_FR_026.3_ok_[square brackets]
Generated strings copied to clipboard.
Do you want to save the generated strings to a file? (y/n): y
Enter the output file type (txt/excel): excel
Enter the path to the Excel file: "C:\Users\namye\OneDrive\Namyeong Galaxy Book\Desktop\String_Generator_Example_Files\Test.xlsx"
Enter the Excel sheet name: Outputs_Test
Warning: A new sheet 'Outputs_Test' will be created in the existing Excel file.
Results saved to 'C:\Users\namye\OneDrive\Namyeong Galaxy Book\Desktop\String_Generator_Example_Files\Test.xlsx' (Excel) in a new sheet 'Outputs_Test'.
Press 'Escape' to close the program...










==== Info ====
Python 3.11.5
Coded using python in VSCode
Written mostly by GPT 3.5, with human corrections
미
