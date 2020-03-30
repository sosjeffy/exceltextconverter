# Excel Text Converter

This miniature python program was created for RAD@UCI.

In order to properly use this program, make sure that you have your Excel file in the same folder as main.py. The saved .txt file will be in the same folder as main.py. Note that this program uses pandas in order to parse through an Excel .xlsx. 

More things to note:
1. This program was created with a particular Excel document format in mind. NAME, as in name of an individual, is stored in column A and EMAIL, as in email of an individual, is stored in column B. Column A's first entry (and thus "label") is "NAME" (without quotes) and column B's first entry (and thus "label") is "EMAIL" (without quotes). 
2. This program will take an Excel document, read through each column that you noted, and put it into a text file with tab delimiters to seperate each column entry. Note that first/last names are also separated by tabs. First and last name are reformatted.
3. You can have the program parse through as many different columns as you wish, however, name/email is always going to be within the first column in the .txt file.
4. Names of people that are not in the standard "John Doe" case (whether they were entried in as "john doe" or "JOHN DoE"), will be fixed and formatted correctly. 

Just run the script and follow on-screen directions and you're good to go!
