# -*- coding: utf-8 -*-
"""
Created on Wed Aug 12 14:48:39 2020

@author: Hotaru LSG
"""

from docx import Document
import os
import re
import tkinter as Tk
from tkinter import ttk, filedialog
from tkinter import Button, Label


def getChapter(s, start, end):
    if s.find(end) == -1:
        return s[s.find(start) + len(start):]
    return s[s.find(start) + len(start):s.find(end)]

def getVerse(s, start, end):
    if bool(re.search(end, s)) == False:
        return s[re.search(start, s).start():]#s[s.find(start):]
    return s[re.search(start, s).start():re.search(end, s).start()]
    #return s[s.find(start):s.find(end)]

def checkExistence(s, what):
    if s.find(what) == -1:
        return False
    return True

def trimChapter(s, end):
    if bool(re.search(end, s)) == False:#s.find(end) == -1:
        return "done"
    return s[re.search(end, s).start():]#s[s.find(end):]
    
def findnth(haystack, needle, n):
    parts= haystack.split(needle, n+1)
    if len(parts)<=n+1:
        return -1
    return len(haystack)-len(parts[-1])-len(needle)


# this is the function called when the button is clicked
def chooseSummaries():
    global fullText, book, infoStart, summariesChunks, csButton
    
    fullText = ""
    book = ""
    infoStart = 0
    
    summariesDir = filedialog.askopenfilename(initialdir = "/",title = "Choose the Summaries.txt file")
    summaries = open(summariesDir, "r", encoding="UTF-8")
    summariesLines = summaries.read()
    summariesChunks = summariesLines.split("--break--")
    



# this is the function called when the button is clicked
def outputButton():
    global outputDir
    
    outputDirectory = filedialog.askdirectory(initialdir = "/",title = "Choose a location for the Output folder")
    #make an output directory into which all subdirectories and *.md files will be placed
    os.mkdir("{}\\Output".format(outputDirectory))
    outputDir = "{}\\Output".format(outputDirectory)
    

# this is the function called when the button is clicked
def docxButton():
    global docxDir
    
    docxDir = filedialog.askdirectory(initialdir = "/",title = "Choose where your *.docx files are located.")
    
# This is a function which increases the progress bar value by the given increment amount
def makeProgress():
	progessBarOne['value']=progessBarOne['value'] + 1
	root.update_idletasks()


# this is the function called when the button is clicked
def startButton():
    
    global counter
    counter = 0
    
    #make an "All Books.md" file, as a link anchor in the output folder
    baseLink = open("{}\\All Books.md".format(outputDir), "w", encoding="UTF-8")
    baseLink.close()
    
    for s in summariesChunks:
        fullText = ""
        counter += 1
        makeProgress()
        summaryInfo = s.split("\n")
        if summaryInfo[0] == "":
            book = summaryInfo[1]
            infoStart = 2
        else:
            book = summaryInfo[0]
            infoStart = 1
        
        #make a directory for the book, into which all subdirectories and *.md files will be placed
        os.mkdir("{}\\{:02d}-{}".format(outputDir, counter, book))

        #now make a "Book of (name).md" file in that directory, which will act as a link anchor
        bookLink = open("{}\\{:02d}-{}\\{:02d}-Book of {}.md".format(outputDir, counter, book, counter, book), "w", encoding="UTF-8")
        bookLink.write("[[All Books]]")
        bookLink.close()
                
        #which document do we want to get from
        document = Document("{}\\{}.docx".format(docxDir, book))
        
        #get the full text from the document
        for i in document.paragraphs:
            fullText = fullText + " " + i.text
        
        #skip the table of contents, just cut it away
        fullText = fullText[findnth(fullText, "Chapter 1 ", 1):]
        
        #get one chapter at a time
        for c in range(1, 500):
            

            if checkExistence(fullText, C[c-1]) == False:
                break
            
            thisChapter = " " + getChapter(fullText, C[c-1], C[c])
            
            #make a subdirectory for the chapter in the output folder
            os.mkdir("{}\\{:02d}-{}\\{}.{}".format(outputDir, counter, book, CF[c-1], book))
            
            #make a "number-name.md" file in the chapter folder, which will act as a link anchor
            chapterLink = open("{}\\{:02d}-{}\\{}.{}\\{:02d}-{}.md".format(outputDir, counter, book, CF[c-1], book, c, book), "w", encoding="UTF-8")
            chapterLink.write("[[{:02d}-Book of {}]]".format(counter, book))
            chapterLink.close()
      
            #get each verse
            for v in range(0, 999):
                
                thisVerse = getVerse(thisChapter, str(V[v]), str(V[v+1]))
                
                if len(thisVerse) < 3:
                    break
                
                thisMd = open("{}\\{:02d}-{}\\{}.{}\\{}.{:02d} {}.md".format(outputDir, counter, book, CF[c-1], book, c, v+1, book), "w", encoding="UTF-8")
                thisMd.write(thisVerse.strip())
                thisMd.write("\n")
                thisMd.write("\n")
                thisMd.write("**Summery**\n")
                thisMd.write("\n")
                for i in range(infoStart, len(summaryInfo)):
                    thisMd.write(summaryInfo[i])
                    thisMd.write("\n")
                thisMd.write("[[{:02d}-{}]]".format(c, book))
                thisMd.close()
                
                thisChapter = trimChapter(thisChapter, str(V[v+1]))
                if thisChapter == "done":
                    break
            
    root.destroy()

#make a list of chapters
C = []
for x in range(1, 600):
    chapterName = "Chapter {}".format(x)
    C.append(chapterName)
    
#make a list of chapters for the filenames
CF = []
for x in range(1, 600):
    chapterName = "Chapter {:02d}".format(x)
    CF.append(chapterName)   
    
#make a list of verses
V = []
for x in range(1, 1000):
    verseName = "\s{}\s".format(x)
    V.append(verseName)


root = Tk()

# This is the section of code which creates the main window
root.geometry('790x378')
root.configure(background='#E0EEEE')
root.title('*.md Maker')


counter = 0


# This is the section of code which creates the a label
Label(root, text='Choose your Summaries.txt file:', bg='#E0EEEE', font=('arial', 16, 'normal')).place(x=26, y=23)
# This is the section of code which creates a button
csButton = Button(root, text='Choose', bg='#FFD39B', font=('arial', 16, 'normal'), command=chooseSummaries).place(x=346, y=13)
# This is the section of code which creates the a label
Label(root, text='A directory containing the output will be created with the name \'Output\'.', bg='#E0EEEE', font=('arial', 16, 'normal')).place(x=26, y=83)
# This is the section of code which creates the a label
Label(root, text='Please choose a directory in which to create that \'Output\' directory:', bg='#E0EEEE', font=('arial', 16, 'normal')).place(x=26, y=113)
# This is the section of code which creates a button
Button(root, text='Choose an output location', bg='#FFD39B', font=('arial', 16, 'normal'), command=outputButton).place(x=396, y=143)
# This is the section of code which creates the a label
Label(root, text='Finally, choose the directory where the *.docx files are located:', bg='#E0EEEE', font=('arial', 16, 'normal')).place(x=26, y=223)
# This is the section of code which creates a button
Button(root, text='Choose *.docx directory', bg='#FFD39B', font=('arial', 16, 'normal'), command=docxButton).place(x=396, y=253)
# This is the section of code which creates a button
Button(root, text='START', bg='#66CD00', font=('arial', 16, 'normal'), command=startButton).place(x=666, y=323)

# This is the section of code which creates a color style to be used with the progress bar
progessBarOne_style = ttk.Style()
progessBarOne_style.theme_use('clam')
progessBarOne_style.configure('progessBarOne.Horizontal.TProgressbar', foreground='#66CD00', background='#66CD00')


# This is the section of code which creates a progress bar
progessBarOne=ttk.Progressbar(root, style='progessBarOne.Horizontal.TProgressbar', orient='horizontal', length=620, mode='determinate', maximum=66, value=1)
progessBarOne.place(x=23, y=329)

root.mainloop()





    
    






