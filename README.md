# Visitor-log
This is my first entry into my repository
VISITOR LOG APPLICATION

1.	Introduction
The project is on developing a desktop application for visitor log. The desktop user interface is built using Tkinter framework. The content was added into database using Sqlite. This application can be used in an office front desk. The visitor will use the application to provide the following information:
•	Date
•	Name
•	Email ID
•	Phone Number
•	Person Visiting
•	Time-In
•	Time-Out

The program displays the database contents in a multi-listbox present on User Interface. The contents will be added into the database upon clicking submit and deleted using the delete button. The application provides the ability to export the information in the database as an excel file by clicking on the Export.xlsx button.

2.	Requirements
•	Python,Idle version 2.7.9
•	Download Tkinter and Tcl/Tk
a.	Tkinter does not need to be downloaded - it's an integral part of all Python distributions (except binary distributions for platforms that don't support Tcl/Tk). Sometimes when Tcl/Tk is required and it can be downloaded from http://www.tcl.tk/software/tcltk/download.html
b.	For Windows or Mac OS 9, the Python installers already contain a full Tcl/Tk and Tkinter python module installation.
c.	For Unix, both Tcl and Tk must be downloaded, built and installed. Tcl needs to be built first. Each source distribution contains a README file explaining the process. It's about as simple as building and installing Python. 
d.	For Mac OS X follow instructions and download/install Tcl/tk. https://www.python.org/download/mac/tcltk/. This program uses the latest ActiveTcl 8.5.17.0.
•	Download and install TkTreectrl 2.4.1 
•	Download and install Sqlite broswer from http://sqlitebrowser.org/ - This can be used to monitor database contents and modifications done through Sqlite in this program. 


3.	Description
The desktop application titled “Visitor Log” is built using Tkinter framework. Tkinter is part of the standard Python distribution, so it is expected to be present. It just needs to be imported. 

STEP 1: IMPORT MODULES
The modules are required for the program to function and should be imported into the python program using “import” statement. This python program gains access to the code in the following packages, modules:
•	Tkinter
•	Sqlite3
•	TkTreectrl
•	Xlsxwriter
•	tkMessageBox
•	re(regular expression)

           STEP 2: CREATE CLASS
          The class in this program is Visitorlog_app that inherits from Tk class of Tkinter module. 

STEP 3: CREATE __init__ method
The __init__ method takes arguments self (reference to current instance of class) and parent. Visitorlog_app derives from Tkinter.Tk, hence the Tkinter.Tk __init__ method is called and parent object is passed through it.
Each GUI element has a parent (its container, usually) as it is a hierarchy of objects. That's why __init__ has a parent parameter. Keeping track of parents is useful when we have to show/hide a group of widgets, repaint them on screen or simply destroy them when application exits. In this program self.parent = parent statement helps keep reference to the parent. The Visitorlog_app object’s methods are invoked.


          STEP 4: CREATE WIDGETS
a.	def label_entry(self)
Creates entry widgets for appropriate label widgets and places it on the grid at the specified row, column location.
b.	def button(self)
Creates button widgets and places it on the grid. The buttons are linked to its respective function, using command statement, which is called upon a click event.
c.	def multilistbox_scrollbar(self)
Creates a multi-listbox using TkTreeCtrl module. Column titles and scrollbar for the listbox are set.



          STEP 5: CREATE A DATABASE & ESTABLISH CONNECTION
a.	def create_db(self)
In this function, connection object which represents the database is created using sqlite3. Once the connection has been established, a cursor objected is created and its execute() method is called in order to perform SQL commands.

          STEP 6: UPDATE DATABASE AND LISTBOX
a.	def select_update(self)
This function updates database contents and populates it in the multi-listbox.

          STEP 7: BUTTON CLICK EVENT HANDLERS
a.	def OnSubmitClick(self)
This function is called upon “Submit” button-click event. When “Submit” button is clicked, the contents in the entry boxes are obtained using .get() method.  Using regular expression, the “Date”, “Name”, “Email Id” and “Phone Number” entry boxes are checked to have information in them that are in appropriate format. 
If the information entered is matched to be correct, and none of the entry boxes are empty the contents are added to the database and consequentially the database contents are updated in the listbox.
If the information entered in the entry boxes is incorrect an error message pops out.
b.	def OnDeleteClick(self)
This function is invoked when the delete button is clicked.
This function enables the user to select a row in the listbox and delete that row in the database and add the updated set of contents back in the listbox. 

c.	def OnExport_XLSXclick(self)
When the “Export.xlsx” button is clicked an excel file is created in the folder containing the .py program file and the database contents are written into an excel sheet.

        STEP 8:  MAIN METHOD: INSTANCIATING OBJECT
a.	if __name__== "__main__"
The main is created which is executed when the program is run from the command-line. In Tkinter, the class is instanciated Visitorlog_app(parent).The parent has “None” because it's the first GUI element built. The window can be titled using (app.title()). Each program is set to loop indefinitely with .mainloop() and allowed to wait for an event to take place such as click of a button.
                        
