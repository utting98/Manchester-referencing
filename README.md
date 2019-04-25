# Manchester-referencing
Code to format inputs in to a Manchester Harvard Style reference list for use in UoM student reports.

Code requires install of docx via:

pip install python-docx

For anyone unfamiliar, if using anaconda, search and open anaconda prompt then type the above code to install package.

When running code it will search for a file called references.docx, if such file does not exist it will create one. If it does exist it will append new references to the end of this reference list. On running the code a tkinter GUI loads up, select your desired source from the drop down menu shown and click confirm. This brings up a list of entry fields required for that source. Ensure the file is not open on your device while operating this code otherwise the file will not be appended.

IMPORTANT:

FORMAT THE INPUTS EXACTLY AS SHOWN IN THE EXAMPLES, THE INPUTS ARE PUNCTUATION AND CASE SENSITIVE FOR THE CORRECT FORMATTING. FOR EXAMPLE IT SHOWS AUTHOR NEEDS A COMMA AFTER THE SURNAME AND FULL STOP AFTER THE INITIAL BUT NOWHERE ELSE SAYS TO INCLUDE FULL STOPS OR COMMMAS. ALSO ENSURE CAPITALISATION IS AS SHOWN.

When you have entered all the fields press generate reference to append a formatted reference to the end of references.docx. Currently even if you close the program and resume another time, as long as references.docx is not tampered with, the code will resume appending the numbered list the next time you generate a reference. These references, once complete can be copied and pasted over (keeping source formatting) to the end of a report.

Further buttons included are a home button to return to the source selector and a clear button to clear the entry fields if you want to do the same source a second time.

Further versions of this code will have more sources available and if any bugs are found please flag them to me and I will address them when I can.
