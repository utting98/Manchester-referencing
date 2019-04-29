"""
Manchester Harvard Reference Generator

This code will create GUI windows to select a source to format in Manchester Harvard
Reference style. Fields will be created to enter the appropriate information. When genereate
reference is chosen the code will append a file called references.docx with numbered and
formatted references generated from the inputs. Ensure the file is not open while attempting
to generate references. Also ensure you follow the example entry formatting, only putting
punctuation if shown in the example input and capitalising as shown in the example input
to make the output correctly formatted.

A lot of this code is very repetitive in functions so one example of each will be fully
commented to explain proccesses but in interest of comment efficiency not all will be.

Joshua Utting

25/04/2019

"""

#imports list
from tkinter import *
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

try: #attempt to open a file called references.docx
    document = Document('references.docx')
    length = len(document.paragraphs) #if file opens find the number of paragraphs (i.e. references)   
    count = length #set count to start from one after the last number in the list
except:
    document = Document() #if file does not exist create a new file
    count=0 #set count to 0 to start reference numbering from 1

#define word document text style as Times New Roman 12pt
style = document.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)

#define function to get the choice from the drop down and call the appropriate reference window
def getchoice(*choices):
    global reftpyechoice, reftypevar, count
    reftypechoice = reftypevar.get() #get the value of the drop down
    if(reftypechoice=='Book'): #if the value is book etc.
        bookwin() #call the bookwindow function to show the entry fields for a book etc
    elif(reftypechoice=='Website'):
        webwin()
    elif(reftypechoice=='Journal (printed)'):
        journal1win()
    elif(reftypechoice=='Journal (electronic)'):
        journal2win()
    elif(reftypechoice=='Lecturer Handout'):
        lechandoutwin()
    elif(reftypechoice=='Thesis'):
        thesiswin()
    elif(reftypechoice=='Presentation'):
        presentwin()
    
#define the book genating reference function
#these generation functions are called by the relevant sourcewin functions below, e.g. bookwin()
def bookgen(authorstring,yearstring,titlestring,subtitlestring,editionstring,pubplacestring,publisherstring,referenceframe):
    global count
    count=count+1 #add one to the count to increase the number in the reference list
    #get the values of each of the inputs of the book window entry fields
    author = authorstring.get()
    year = yearstring.get()
    title = titlestring.get()
    subtitle = subtitlestring.get()
    edition = editionstring.get()
    pubplace = pubplacestring.get()
    publisher = publisherstring.get()
    
    if(subtitle!=''): #check if there has been an entry in the subtitle window, if so:
        newsubtitlestring = subtitle + '.' #add full stop to the end of subtitle
        subtitle = newsubtitlestring #overwrite subtitle
        newtitlestring = title + ': ' #add colon to the end of title
        title = newtitlestring+subtitle #new title is in format title: subtitle.
    else:
        newtitlestring = title + '.' #if subtitle is empty add full stop to the end of title 
        title = newtitlestring #new title is in format title.
    
    p = document.add_paragraph("[%s] " % count) #add a new paragraph to document starting with count value e.g reference [3]
    p.style = document.styles['Normal'] #write paragraph in previously defined font style
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY #justify aligning of paragraph as per required for UoM reports
    p.add_run("%s (%s). " % (author,year)) #add the title and year in brackets to the reference
    p.add_run(title).italic = True #add the title to the paragraph in italics
    p.add_run(" %s. %s: %s." % (edition,pubplace,publisher)) #add the edition. publishier location: publisher name. to the paragraph
    document.save('references.docx') #save the appending to references.docx
    success = Label(referenceframe,text=('Reference [%s] Successfully Generated' % count)).grid(row=8,column=0,columnspan=2) #create a label at the bottom of the window informing reference [number] added successfully

#define the web reference generating function
#works as bookgen but all formatting matches that required for websites
def webgen(authorstring,yearstring,titlestring,urlstring,accessstring,referenceframe):
    global count
    count=count+1
    author = authorstring.get()
    year = yearstring.get()
    title = titlestring.get()
    url = urlstring.get()
    access = accessstring.get()
    
    p = document.add_paragraph("[%s] " % count)
    p.style = document.styles['Normal']
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("%s (%s). " % (author,year))
    p.add_run(title).italic = True
    p.add_run(". Available at: %s, (Accessed: %s)." % (url,access))
    document.save('references.docx')
    success = Label(referenceframe,text=('Reference [%s] Successfully Generated' % count)).grid(row=8,column=0,columnspan=2)
    
#define the printed journal reference generating function
#works as bookgen but all formatting matches that required for printed journals
def journal1gen(authorstring,yearstring,titlearticlestring,titlestring,volumestring,issuestring,pagestring,referenceframe):
    global count
    count=count+1
    author = authorstring.get()
    year = yearstring.get()
    titlearticle = titlearticlestring.get()
    title = titlestring.get()
    volume = volumestring.get()
    issue = issuestring.get()
    page = pagestring.get()
    
    p = document.add_paragraph("[%s] " % count)
    p.style = document.styles['Normal']
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("%s (%s). ‘%s’, " % (author,year,titlearticle))
    p.add_run(title).italic = True
    p.add_run(", %s(%s), pp. %s." % (volume,issue,page))
    document.save('references.docx')
    success = Label(referenceframe,text=('Reference [%s] Successfully Generated' % count)).grid(row=8,column=0,columnspan=2)

#define the electronic journal reference generating function
#works as bookgen but all formatting matches that required for electronic journals
def journal2gen(authorstring,yearstring,titlearticlestring,titlestring,volumestring,issuestring,pagestring,urlstring,accessstring,referenceframe):
    global count
    count=count+1
    author = authorstring.get()
    year = yearstring.get()
    titlearticle = titlearticlestring.get()
    title = titlestring.get()
    volume = volumestring.get()
    issue = issuestring.get()
    page = pagestring.get()
    url = urlstring.get()
    access = accessstring.get()
    
    p = document.add_paragraph("[%s] " % count)
    p.style = document.styles['Normal']
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("%s (%s). ‘%s’, " % (author,year,titlearticle))
    p.add_run(title).italic = True
    p.add_run(", %s(%s), pp. %s. [Online]. Available at: %s (Accessed: %s)." % (volume,issue,page,url,access))
    document.save('references.docx')
    success = Label(referenceframe,text=('Reference [%s] Successfully Generated' % count)).grid(row=10,column=0,columnspan=2)

#define the lecturer handout reference generating function
#works as bookgen but all formatting matches that required for lecturer handouts
def lechandoutgen(authorstring,yearstring,titlestring,codestring,modtitlestring,institutionstring,referenceframe):
    global count
    count=count+1
    author = authorstring.get()
    year = yearstring.get()
    title = titlestring.get()
    code = codestring.get()
    modtitle = modtitlestring.get()
    institution = institutionstring.get()
    
    unitname = code + ': ' + modtitle
    
    p = document.add_paragraph("[%s] " % count)
    p.style = document.styles['Normal']
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("%s (%s). ‘%s’, " % (author,year,title))
    p.add_run(unitname).italic = True
    p.add_run(". %s. Unpublished." % (institution))
    document.save('references.docx')
    success = Label(referenceframe,text=('Reference [%s] Successfully Generated' % count)).grid(row=7,column=0,columnspan=2)

#define the thesis reference generating function
#works as bookgen but all formatting matches that required for theses
def thesisgen(authorstring,yearstring,titlestring,levelstring,institutionstring,referenceframe):
    global count
    count=count+1
    author = authorstring.get()
    year = yearstring.get()
    title = titlestring.get()
    level = levelstring.get()
    institution = institutionstring.get()
    
    p = document.add_paragraph("[%s] " % count)
    p.style = document.styles['Normal']
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("%s (%s). " % (author,year))
    p.add_run(title).italic = True
    p.add_run(". %s. %s." % (level,institution))
    document.save('references.docx')
    success = Label(referenceframe,text=('Reference [%s] Successfully Generated' % count)).grid(row=6,column=0,columnspan=2)

#define the presentation reference generating function
#works as bookgen but all formatting matches that required for presentations
def presentgen(authorstring,yearstring,titlestring,codestring,modtitlestring,urlstring,accessstring,referenceframe):
    global count
    count=count+1
    author = authorstring.get()
    year = yearstring.get()
    title = titlestring.get()
    code = codestring.get()
    modtitle = modtitlestring.get()
    url = urlstring.get()
    access = accessstring.get()
    
    p = document.add_paragraph("[%s] " % count)
    p.style = document.styles['Normal']
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("%s (%s). " % (author,year))
    p.add_run(title).italic = True
    p.add_run(" [presentation]. ")
    p.add_run("%s: %s" % (code,modtitle)).italic = True
    p.add_run(". Available at: %s (Accessed: %s)." % (url,access))
    document.save('references.docx')
    success = Label(referenceframe,text=('Reference [%s] Successfully Generated' % count)).grid(row=8,column=0,columnspan=2)    

#define bookwindow for reference inputs, called when book source is chosen to overwrite the display
def bookwin():
    global root, homeframe #make the tkinter root and the frame of the home screen globals
    homeframe.destroy() #destroy the homeframe to make the tkinter window blank
    referenceframe = Frame(root) #create a new frame for the references in root
    referenceframe.pack() #pack the new reference frame in the window
    
    #generate label at top of window reminding to enter inputs as shown in example and show in the first row and first and second columns of the grid
    infolabel = Label(referenceframe,text='Enter information for reference. Enter inputs as shown, case and punctuation sensitive.').grid(row=0,column=0,columnspan=2)
    
    #generate label to tell to input author and show example and show in second row and  first column of grid
    authorlabel = Label(referenceframe,text='Author(s) (e.g. Utting, J.)').grid(row=1,column=0)
    #define the string variable the entry will be stored to
    authorstring = StringVar()
    #define the entry window for the text and set it in the first row and second column of the grid
    authorentry = Entry(referenceframe,textvariable=authorstring).grid(row=1,column=1)
    
    #generate label to tell to input year and show example and show in third row and first column of grid
    yearlabel = Label(referenceframe,text='Publication Year (e.g. 2019)').grid(row=2,column=0)
    #define the string variable the entry will be stored to
    yearstring = StringVar()
    #define the entry window for the text and set it in the third row and second column of the grid
    yearstringentry = Entry(referenceframe,textvariable=yearstring).grid(row=2,column=1)
    
    #generate label to tell to input title and show example and show in fourth row and first column of grid
    titlelabel = Label(referenceframe,text='Title Of Book (e.g. Beginner programming)').grid(row=3,column=0)
    #define the string variable the entry will be stored to
    titlestring = StringVar()
    #define the entry window for the text and set it in the fourth row and second column of the grid
    titleentry = Entry(referenceframe,textvariable=titlestring).grid(row=3,column=1)
    
    #generate label to tell to input subtitle and show example and show in fifth row and first column of grid
    subtitlelabel = Label(referenceframe,text='Subtitle (If none leave blank, e.g. testing)').grid(row=4,column=0)
    #define the string variable the entry will be stored to
    subtitlestring = StringVar()
    #define the entry window for the text and set it in the fifth row and second column of the grid
    subtitleentry = Entry(referenceframe,textvariable=subtitlestring).grid(row=4,column=1)
    
    #generate label to tell to input edition and show example and show in sixth row and first column of grid
    editionlabel = Label(referenceframe,text='Edition (e.g. 3rd edn)').grid(row=5,column=0)
    #define the string variable the entry will be stored to
    editionstring = StringVar()
    #define the entry window for the text and set it in the sixth row and second column of the grid
    editionentry = Entry(referenceframe,textvariable=editionstring).grid(row=5,column=1)
    
    #generate label to tell to input publshing place and show example and show in seventh row and first column of grid
    pubplacelabel = Label(referenceframe,text='Place Of Publication (e.g. Manchester)').grid(row=6,column=0)
    #define the string variable the entry will be stored to
    pubplacestring = StringVar()
    #define the entry window for the text and set it in the seventh row and second column of the grid
    pubplaceentry = Entry(referenceframe,textvariable=pubplacestring).grid(row=6,column=1)
    
    #generate label to tell to input publisher and show example and show in eighth row and first column of grid
    publisherlabel = Label(referenceframe,text='Publisher Name (e.g. Pearson)').grid(row=7,column=0)
    #define the string variable the entry will be stored to
    publisherstring = StringVar()
    #define the entry window for the text and set it in the eighth row and second column of the grid
    publisherentry = Entry(referenceframe,textvariable=publisherstring).grid(row=7,column=1)
    
    #define a button to generate the reference, it has the bookgen function bound to it so will call it on button press to make a reference from the entered information, show this button in the tenth row and first column, note the ninth row skip is to display the success label in the bookgen field above this
    genrefbutton = Button(referenceframe,text='Generate Reference',command=lambda: bookgen(authorstring,yearstring,titlestring,subtitlestring,editionstring,pubplacestring,publisherstring,referenceframe)).grid(row=9,column=1)
    #define the home button to return from the current reference window back to the sourve selector if desired, calls the home function on button press, show button in tenth row and first column
    homebutton = Button(referenceframe,text='Home',command=lambda: home(referenceframe)).grid(row=9,column=0)
    #define the clear button to empty all fields (actually destroys frame and rebuilds fast enough not to notice in fullscreen), shows this button approximateky midway through (fifth row) and in the third column
    clearbutton = Button(referenceframe,text='Clear',command=lambda: clear(referenceframe,bookwin)).grid(row=4,column=2)
    
    root.update() #used instead of root.mmainloop() to allow for changes using the clear command, updates the window on each loop rather than just loops with the original layout
    return #code should not make it this far but left in for any breakages to try and return to previous function

#define webwindow for reference inputs, works as bookwin
def webwin():
    global root, homeframe
    homeframe.destroy()
    referenceframe = Frame(root)
    referenceframe.pack()
    
    infolabel = Label(referenceframe,text='Enter information for reference. Enter inputs as shown, case and punctuation sensitive.').grid(row=0,column=0,columnspan=2)
    
    authorlabel = Label(referenceframe,text='Author(s) (e.g. Utting, J.)').grid(row=1,column=0)
    authorstring = StringVar()
    authorentry = Entry(referenceframe,textvariable=authorstring).grid(row=1,column=1)
    
    yearlabel = Label(referenceframe,text='Website Publication Year (e.g. 2019)').grid(row=2,column=0)
    yearstring = StringVar()
    yearstringentry = Entry(referenceframe,textvariable=yearstring).grid(row=2,column=1)

    titlelabel = Label(referenceframe,text='Title Of Website (e.g. Beginner programming)').grid(row=3,column=0)
    titlestring = StringVar()
    titleentry = Entry(referenceframe,textvariable=titlestring).grid(row=3,column=1)

    urllabel = Label(referenceframe,text='Website URL (e.g. www.google.co.uk)').grid(row=4,column=0)
    urlstring = StringVar()
    urlentry = Entry(referenceframe,textvariable=urlstring).grid(row=4,column=1)

    accesslabel = Label(referenceframe,text='Access Date (e.g. 20th April 2019)').grid(row=5,column=0)
    accessstring = StringVar()
    accessentry = Entry(referenceframe,textvariable=accessstring).grid(row=5,column=1)
        
    genrefbutton = Button(referenceframe,text='Generate Reference',command=lambda: webgen(authorstring,yearstring,titlestring,urlstring,accessstring,referenceframe)).grid(row=9,column=1)
    homebutton = Button(referenceframe,text='Home',command=lambda: home(referenceframe)).grid(row=9,column=0)
    clearbutton = Button(referenceframe,text='Clear',command=lambda: clear(referenceframe,webwin)).grid(row=3,column=2)
    root.update()
    return

#define journal1window for printed journal reference inputs, works as bookwin
def journal1win():
    global root, homeframe
    homeframe.destroy()
    referenceframe = Frame(root)
    referenceframe.pack()
    
    infolabel = Label(referenceframe,text='Enter information for reference. Enter inputs as shown, case and punctuation sensitive.').grid(row=0,column=0,columnspan=2)
    
    authorlabel = Label(referenceframe,text='Author(s) (e.g. Utting, J.)').grid(row=1,column=0)
    authorstring = StringVar()
    authorentry = Entry(referenceframe,textvariable=authorstring).grid(row=1,column=1)
    
    yearlabel = Label(referenceframe,text='Publication Year (e.g. 2019)').grid(row=2,column=0)
    yearstring = StringVar()
    yearstringentry = Entry(referenceframe,textvariable=yearstring).grid(row=2,column=1)
    
    titlearticlelabel = Label(referenceframe,text='Title Of Article (e.g. Beginner programming)').grid(row=3,column=0)
    titlearticlestring = StringVar()
    titlearticleentry = Entry(referenceframe,textvariable=titlearticlestring).grid(row=3,column=1)

    titlelabel = Label(referenceframe,text='Title Of Journal (e.g. Manchester Programming Journal)').grid(row=4,column=0)
    titlestring = StringVar()
    titleentry = Entry(referenceframe,textvariable=titlestring).grid(row=4,column=1)

    volumelabel = Label(referenceframe,text='Volume Number (e.g. 8)').grid(row=5,column=0)
    volumestring = StringVar()
    volumestringentry = Entry(referenceframe,textvariable=volumestring).grid(row=5,column=1)

    issuelabel = Label(referenceframe,text='Issue/Part Number (e.g. 3)').grid(row=6,column=0)
    issuestring = StringVar()
    issuestringentry = Entry(referenceframe,textvariable=issuestring).grid(row=6,column=1)

    pagelabel = Label(referenceframe,text='Page Range (e.g. 55-59)').grid(row=7,column=0)
    pagestring = StringVar()
    pagestringentry = Entry(referenceframe,textvariable=pagestring).grid(row=7,column=1)

    genrefbutton = Button(referenceframe,text='Generate Reference',command=lambda: journal1gen(authorstring,yearstring,titlearticlestring,titlestring,volumestring,issuestring,pagestring,referenceframe)).grid(row=9,column=1)
    homebutton = Button(referenceframe,text='Home',command=lambda: home(referenceframe)).grid(row=9,column=0)
    clearbutton = Button(referenceframe,text='Clear',command=lambda: clear(referenceframe,journal1win)).grid(row=4,column=2)
    root.update()
    return

#define journal2window for online journal reference inputs, works as bookwin
def journal2win():
    global root, homeframe
    homeframe.destroy()
    referenceframe = Frame(root)
    referenceframe.pack()
    
    infolabel = Label(referenceframe,text='Enter information for reference. Enter inputs as shown, case and punctuation sensitive.').grid(row=0,column=0,columnspan=2)
    
    authorlabel = Label(referenceframe,text='Author(s) (e.g. Utting, J.)').grid(row=1,column=0)
    authorstring = StringVar()
    authorentry = Entry(referenceframe,textvariable=authorstring).grid(row=1,column=1)
    
    yearlabel = Label(referenceframe,text='Publication Year (e.g. 2019)').grid(row=2,column=0)
    yearstring = StringVar()
    yearstringentry = Entry(referenceframe,textvariable=yearstring).grid(row=2,column=1)
    
    titlearticlelabel = Label(referenceframe,text='Title Of Article (e.g. Beginner programming)').grid(row=3,column=0)
    titlearticlestring = StringVar()
    titlearticleentry = Entry(referenceframe,textvariable=titlearticlestring).grid(row=3,column=1)

    titlelabel = Label(referenceframe,text='Title Of Journal (e.g. Manchester Programming Journal)').grid(row=4,column=0)
    titlestring = StringVar()
    titleentry = Entry(referenceframe,textvariable=titlestring).grid(row=4,column=1)

    volumelabel = Label(referenceframe,text='Volume Number (e.g. 8)').grid(row=5,column=0)
    volumestring = StringVar()
    volumestringentry = Entry(referenceframe,textvariable=volumestring).grid(row=5,column=1)

    issuelabel = Label(referenceframe,text='Issue/Part Number (e.g. 3)').grid(row=6,column=0)
    issuestring = StringVar()
    issuestringentry = Entry(referenceframe,textvariable=issuestring).grid(row=6,column=1)

    pagelabel = Label(referenceframe,text='Page Range (e.g. 55-59)').grid(row=7,column=0)
    pagestring = StringVar()
    pagestringentry = Entry(referenceframe,textvariable=pagestring).grid(row=7,column=1)

    urllabel = Label(referenceframe,text='Jounal URL (e.g. www.google.co.uk)').grid(row=8,column=0)
    urlstring = StringVar()
    urlentry = Entry(referenceframe,textvariable=urlstring).grid(row=8,column=1)

    accesslabel = Label(referenceframe,text='Access Date (e.g. 20th April 2019)').grid(row=9,column=0)
    accessstring = StringVar()
    accessentry = Entry(referenceframe,textvariable=accessstring).grid(row=9,column=1)
        
    genrefbutton = Button(referenceframe,text='Generate Reference',command=lambda: journal2gen(authorstring,yearstring,titlearticlestring,titlestring,volumestring,issuestring,pagestring,urlstring,accessstring,referenceframe)).grid(row=11,column=1)
    homebutton = Button(referenceframe,text='Home',command=lambda: home(referenceframe)).grid(row=11,column=0)
    clearbutton = Button(referenceframe,text='Clear',command=lambda: clear(referenceframe,journal2win)).grid(row=5,column=2)
    root.update()
    return

#define lechandoutwindow for reference inputs, works as bookwin
def lechandoutwin():
    global root, homeframe
    homeframe.destroy()
    referenceframe = Frame(root)
    referenceframe.pack()
    
    infolabel = Label(referenceframe,text='Enter information for reference. Enter inputs as shown, case and punctuation sensitive.').grid(row=0,column=0,columnspan=2)
    
    authorlabel = Label(referenceframe,text='Lecturer Name (e.g. Utting, J.)').grid(row=1,column=0)
    authorstring = StringVar()
    authorentry = Entry(referenceframe,textvariable=authorstring).grid(row=1,column=1)
    
    yearlabel = Label(referenceframe,text='Distribution Year (e.g. 2019)').grid(row=2,column=0)
    yearstring = StringVar()
    yearstringentry = Entry(referenceframe,textvariable=yearstring).grid(row=2,column=1)

    titlelabel = Label(referenceframe,text='Title Of Handout (e.g. User interface programming)').grid(row=3,column=0)
    titlestring = StringVar()
    titleentry = Entry(referenceframe,textvariable=titlestring).grid(row=3,column=1)

    codelabel = Label(referenceframe,text='Module Code (e.g. PHYS101)').grid(row=4,column=0)
    codestring = StringVar()
    codeentry = Entry(referenceframe,textvariable=codestring).grid(row=4,column=1)

    modtitlelabel = Label(referenceframe,text='Title Of Module (e.g. Beginner programming)').grid(row=5,column=0)
    modtitlestring = StringVar()
    modtitleentry = Entry(referenceframe,textvariable=modtitlestring).grid(row=5,column=1)

    institutionlabel = Label(referenceframe,text='Insititution Name (e.g. University of Manchester)').grid(row=6,column=0)
    institutionstring = StringVar()
    institutionentry = Entry(referenceframe,textvariable=institutionstring).grid(row=6,column=1)

    genrefbutton = Button(referenceframe,text='Generate Reference',command=lambda: lechandoutgen(authorstring,yearstring,titlestring,codestring,modtitlestring,institutionstring,referenceframe)).grid(row=8,column=1)
    homebutton = Button(referenceframe,text='Home',command=lambda: home(referenceframe)).grid(row=8,column=0)
    clearbutton = Button(referenceframe,text='Clear',command=lambda: clear(referenceframe,lechandoutwin)).grid(row=4,column=2)
    root.update()

#define thesiswindow for reference inputs, works as bookwin
def thesiswin():
    global root, homeframe
    homeframe.destroy()
    referenceframe = Frame(root)
    referenceframe.pack()
    
    infolabel = Label(referenceframe,text='Enter information for reference. Enter inputs as shown, case and punctuation sensitive.').grid(row=0,column=0,columnspan=2)
    
    authorlabel = Label(referenceframe,text='Author(s) (e.g. Utting, J.)').grid(row=1,column=0)
    authorstring = StringVar()
    authorentry = Entry(referenceframe,textvariable=authorstring).grid(row=1,column=1)
    
    yearlabel = Label(referenceframe,text='Submission Year (e.g. 2019)').grid(row=2,column=0)
    yearstring = StringVar()
    yearstringentry = Entry(referenceframe,textvariable=yearstring).grid(row=2,column=1)

    titlelabel = Label(referenceframe,text='Title Of Thesis (e.g. Statistical modelling of neutron scattering)').grid(row=3,column=0)
    titlestring = StringVar()
    titleentry = Entry(referenceframe,textvariable=titlestring).grid(row=3,column=1)
    
    levellabel = Label(referenceframe,text='Level Of Thesis (e.g. PhD').grid(row=4,column=0)
    levelstring = StringVar()
    levelentry = Entry(referenceframe,textvariable=levelstring).grid(row=4,column=1)
    
    institutionlabel = Label(referenceframe,text='Institution Name (e.g. University of Manchester)').grid(row=5,column=0)
    institutionstring = StringVar()
    institutionentry = Entry(referenceframe,textvariable=institutionstring).grid(row=5,column=1)

    genrefbutton = Button(referenceframe,text='Generate Reference',command=lambda: thesisgen(authorstring,yearstring,titlestring,levelstring,institutionstring,referenceframe)).grid(row=7,column=1)
    homebutton = Button(referenceframe,text='Home',command=lambda: home(referenceframe)).grid(row=7,column=0)
    clearbutton = Button(referenceframe,text='Clear',command=lambda: clear(referenceframe,thesiswin)).grid(row=3,column=2)
    root.update()

#define presentationwindow for reference inputs, works as bookwin
def presentwin():
    global root,homeframe
    homeframe.destroy()
    referenceframe = Frame(root)
    referenceframe.pack()
    
    infolabel = Label(referenceframe,text='Enter information for reference. Enter inputs as shown, case and punctuation sensitive.').grid(row=0,column=0,columnspan=2)
    
    authorlabel = Label(referenceframe,text='Author(s) (e.g. Utting, J.)').grid(row=1,column=0)
    authorstring = StringVar()
    authorentry = Entry(referenceframe,textvariable=authorstring).grid(row=1,column=1)
    
    yearlabel = Label(referenceframe,text='Submission Year (e.g. 2019)').grid(row=2,column=0)
    yearstring = StringVar()
    yearstringentry = Entry(referenceframe,textvariable=yearstring).grid(row=2,column=1)

    titlelabel = Label(referenceframe,text='Title Of Presentation (e.g. Introductory programming)').grid(row=3,column=0)
    titlestring = StringVar()
    titleentry = Entry(referenceframe,textvariable=titlestring).grid(row=3,column=1)
    
    codelabel = Label(referenceframe,text='Module Code (e.g. PHYS101)').grid(row=4,column=0)
    codestring = StringVar()
    codeentry = Entry(referenceframe,textvariable=codestring).grid(row=4,column=1)

    modtitlelabel = Label(referenceframe,text='Title Of Module (e.g. Beginner python)').grid(row=5,column=0)
    modtitlestring = StringVar()
    modtitleentry = Entry(referenceframe,textvariable=modtitlestring).grid(row=5,column=1)

    urllabel = Label(referenceframe,text='Website URL (e.g. www.google.co.uk)').grid(row=6,column=0)
    urlstring = StringVar()
    urlentry = Entry(referenceframe,textvariable=urlstring).grid(row=6,column=1)

    accesslabel = Label(referenceframe,text='Access Date (e.g. 20th April 2019)').grid(row=7,column=0)
    accessstring = StringVar()
    accessentry = Entry(referenceframe,textvariable=accessstring).grid(row=7,column=1)
        
    genrefbutton = Button(referenceframe,text='Generate Reference',command=lambda: presentgen(authorstring,yearstring,titlestring,codestring,modtitlestring,urlstring,accessstring,referenceframe)).grid(row=9,column=1)
    homebutton = Button(referenceframe,text='Home',command=lambda: home(referenceframe)).grid(row=9,column=0)
    clearbutton = Button(referenceframe,text='Clear',command=lambda: clear(referenceframe,presentwin)).grid(row=4,column=2)
    root.update()

#define quitwindow function to be called by quit buttons
def quitwindow():
    global root #make access to root global
    root.destroy() #destroy root to end program

#define home function to be called by home buttons
def home(frame):
    frame.destroy() #destroy imported frame, in this case the referenceframe of the previously chosen reference
    main() #call main again to show source selector

#define clear funciton to clear current entries if the same source wants to be used again
def clear(frame,functionreturn):
    global root #make access to root global
    frame.destroy() #destroy the current frame that is given in the call of this function
    functionreturn() #call the function window to redraw the input fields of the source this function was called from

#define main function for source selection window
def main():
    global root,homeframe,reftypevar,countvar #make access to root, homeframe, referency type variable and the count variable global
    
    homeframe = Frame(root) #create a frame called homeframe in root
    homeframe.pack() #pack the frame in the root window
    
    reftypevar=StringVar() #define reference type variable as string
    choices = {'Book','Journal (electronic)','Journal (printed)','Presentation','Lecturer Handout','Thesis','Website'} #define drop down menu choices for sources
    reftypevar.set('Book') #set default value to book
    
    droplabel = Label(homeframe,text='Choose Source').grid(row=0,column=0) #define label for drop menu and pack it to the first row and first column
    dropmenu = OptionMenu(homeframe,reftypevar,*choices).grid(row=0,column=1) #define drop down menu with the previously defined choices and pack to first row and second column
    #define the confirm button which calls the function to find out what selection was made and calls the appropriate window to generate, pack it to the first row and third column
    refgenbutton = Button(homeframe,text='Confirm',command=getchoice).grid(row=0,column=2)
    #define the quitbutton which calls the quitwindow function on buttonpress and closes the program
    quitbutton = Button(homeframe,text='Quit',command=quitwindow).grid(row=0,column=3)
    #loop root over these defined objects to keep it displayed until something is done to it
    root.mainloop()

root = Tk() #define root as a tkinter window
root.title('Manchester Harvard Reference Generator') #give the window this title string
main() #cal the main function
    
    
    
    
    
