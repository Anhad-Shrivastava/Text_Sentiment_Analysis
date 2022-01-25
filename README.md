# Text_Sentiment_Analysis
My first experience using NLP, working in nltk library of Python

Instructions:-

To run DataFrames.py:-
delete positive_dictionary.xlsx and negative_dictionary.xlsx, the 'DataFrames.py' code  will use openpyxl to open master dictionary and get the positive and negative words and create these 2 workbooks in the same folder
Note: Workbook opening takes time as the workbook is of a medium large size

To run Main.py:-
open the "Output Data.xlsx" file and select the 'blank copy' sheet and close the file. Run main.py to get the desired data in the blank copy

Compiler Output:-
The compiler output will show program start/end, url from which the program is extracting information and the time taken by the program to process each url and store the desired variables and the total time taken by the code(approx 15-17 min on my pycharm application)

Contents of zip file:-
2- .ipynb Jupyter notebookks where codes were tried and worked
2- .py Python files, one used to get the positive and negative words dictionaries from the master dictionary
1- .txt Text file with stop words which are extracted for use
1- .txt Readme file with all instructions
7- .xlsx Excel workbooks, 4 workbooks contain dictionaries of positive, negative, constraining and uncertainty words, 1 workbook is the Master dictionary which was used to extract the positive and negative words dictionaries, a cik workbook, containing the original contents of the input xlsx file and a Output Data workbook containing the output file with all required variables filled for all the rows, the desired file

About the file:-

->The Desired Output can be found in Output Data with all the required variables calculated and textual analysis complete
->The .py files are extensively well commented with easily understandable variable names, most of the code is self-explanatory

***********************************************************************************************************************************************************************
Textual Analysis Process-

Libraries Used:
pandas,openpyxl,regex,urllib,nltk,bs4(Beautiful Soup)

PreCode:
Step1: All functions required are defined, global variables defined
Step2: Input and Output notebooks are opened for use using openpyxl
Step3: Dictionary of all stopwords, positive words, negative words etc is made. We use dictionry instead of list because the search complexity is of constant time compared to list which has a search complexty of linear time
Step4: Data is extracted, cleaned and positive, negative etc scores are stored in global variables for further use

Cleaning Data: Data is cleaned by removing all numeric data, noise, xml codes etc. Function getGlobalVariables() perorms the task of extracting the data using urllib, cleaning all the data using nltk, BeautifulSoup and finally, it finds number of positive, negative etc. words which are stored in global variables so they can be used across functions

Textual Analysis: For a given row and a list of clean data, all the required calculations are performed by this fucntion and the output values are stored in their respective columns.

File Opening: Sometimes, the webpage denies access to urllib function, so a recursive function called open file keeps sending requests till accepted, we can add more headers to the rewuest to appear less bot and more human while making a request, but this way also works

Workbook Closing: In the end, we close the workbooks that were opened for our tasks

Main Function: It simply calls the above functions to get a clean list and to use that list to perform the calculations and store the values in the appropriate cells of the output file
