# Manchester-referencing
Code to format inputs in to a Manchester Harvard Style reference list for use in UoM student reports.

Current source types software can generate: 'Book','Chapter From Edited Book','E-Book','Edited Book','Government/Corporate Publications','Illustration','Journal (electronic)','Journal (printed)','Lecturer Handout','Presentation','Report From Organisation','Self-Citation','Software','Thesis','Website','Custom'

Code requires install of docx and beautifulsoup via:

pip install python-docx

pip install bs4

For anyone unfamiliar, if using anaconda, search and open anaconda prompt then type the above code and hit enter to install packages.

When running code it will search for a file called references.docx, if such file does not exist it will create one. If it does exist it will append new references to the end of this reference list. On running the code a tkinter GUI loads up, select your desired source from the drop down menu shown and click confirm. This brings up a list of entry fields required for that source. Ensure the references file is not open on your device while operating this code otherwise the file will not be appended.

IMPORTANT:

[FORMAT THE INPUTS EXACTLY AS SHOWN IN THE EXAMPLES, THE INPUTS ARE PUNCTUATION AND CASE SENSITIVE FOR THE CORRECT FORMATTING. FOR EXAMPLE IT SHOWS AUTHOR NEEDS A COMMA AFTER THE SURNAME AND FULL STOP AFTER THE INITIAL BUT NOWHERE ELSE SAYS TO INCLUDE FULL STOPS OR COMMMAS. ALSO ENSURE CAPITALISATION IS AS SHOWN.]

The author section is the only entry field which requires punctuation, for author formatting this is as follows:

1 author: (Surname), (initial).

2 authors: (Surname), (initial). and (Surname), (initial).

3 or more authors: (Surname), (initial)., (Surname), (initial). [...] and (Surname), (initial).

When you have entered all the fields press generate reference to append a formatted reference to the end of references.docx. Currently even if you close the program and resume another time, as long as references.docx is not tampered with, the code will resume appending the numbered list the next time you generate a reference. These references, once complete can be copied and pasted over (keeping source formatting) to the end of a report.

Further buttons included are a home button to return to the source selector and a clear button to clear the entry fields if you want to do the same source a second time and a newly added ISBN search.

ISBN search is a trial feature where for any book based source you can enter the ISBN number of the book and the form will attempt to autofill all of the entry fields. It is important to check the automatically entered data is what you expect as it scrapes the data from google books so may not always be correct. If there was a fail while attempting to find the information during the search a popup will come up to advise on the next steps. If all of the fields are autofilled correctly you can then click generate and your reference will be generated as normal for manual inputs and appending the reference list.

To reiterate please check all autofilled inputs to ensure that all information is correct and all entries are formatted as per the example text for each input field denoted (e.g. [EXAMPLE]).

Search archives is a trial feature (added 13/11/2019) for reusing previously generated references. When a reference is generated in any source menu (with the exception of custom mode) as well as being added to your current reference list, the full reference will be added to archivedreferences.csv. On the home menu there is an option to search for archived references, you can type any search term(s) that is in the reference you are looking for. For example you could put the author name and/or the title of the source. If a match(es) is found it will display the source type and the full reference string in a drop down list, this allows the user to identify which match is the reference they are reusing. The displayed matches will be in order of best to worst match based on the number of search terms it matced with. Choose the reference desired from the selection and press confirm then the reference will be appended to your reference list as if you were generating it from new. This way it is possible to quickly reproduce references that you have used before to save time. If no match is found from your search terms a label is displayed below the search box to inform the user there were no matches.

If any bugs are found please flag them to me and I will address them when I can.
