##========================================
##Created by : Collin Preston
##Created date: 8/10/2021
##========================================

"""
WorkspaceOne Group Check

Helps Techs determite the exchange group that an end user is a part of. During the
switch from on premises exchange to cloud services. Users would need to be placed into
specific groups in Active Directory. Techs and help desk had to use a powershell command
Get-ADUser {} -Properties msExchRecipientTypeDetails, which would return a value. Depending
on the value determined which group they need to be assigned to. This application
simplifies that process by tell you the Group name and also allowing you to add the user
to the group directly. Simple to use GUI and quick add to groups.

"""


##==========================
## Imports
##==========================
import os
import tkinter as tk
from tkinter import *
import subprocess
import tkinter.font as tkFont
from tkinter import messagebox
import signal
import time

#GUI
window = Tk()
window.geometry("380x200")
fontstyle1 = tkFont.Font(family="Lucida Grande", size=20)
fontstyle3 = tkFont.Font(family="Lucida Grande", size=8)
fontstyle4 = tkFont.Font(family="Showcard Gothic", size=20)
fontstyle5 = tkFont.Font(family="Lucida Grande", size=7)
window.title("WorkspaceOne Group Check")
header = Label(window,font=fontstyle1,  text="Employee ID" )
header.place(x=120,y=5)
footer = Label(window,font=fontstyle1,  text="AD GROUP" )
footer.place(x=120,y=125)
infoboxcheck = 1
infobtn = Button(window, width=6,height=1,font=fontstyle3,text="?",bg="firebrick4",fg="white",activebackground="white",activeforeground="firebrick4", command = lambda: about())
infobtn.place(x=330,y=5) 
e1 = Entry(window,show=None, bd =3, width=50)
e1.place(x=35,y=40)
var = StringVar()
label = Entry(window,show=None, bd =3, width=50)


#Defines Relative path for files used in GUI, images/icons/ico
def img_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#Info Button, Version info, Contact information
def about3():
    global infoboxcheck
    infoboxcheck = 1

def about2(popup):
    popup.destroy()


def about():
    global infoboxcheck
    if infoboxcheck == 1:
        infoboxcheck = 0
        root_x = window.winfo_rootx()
        root_y = window.winfo_rooty()
        popup = Tk()
        win_x = root_x + 150
        win_y = root_y + 200
        popup.geometry(f'+{win_x}+{win_y}')
        popup.wm_title("Program Info")
        
        popup.configure(background='sky blue')
        infofont1 = tkFont.Font(family="Lucida Grande", size=20, weight=tkFont.BOLD)        
        infolabel = Label(popup, font = infofont1, text="WorkspaceOne Group Check\nVersion 1.2\nCopyright 2021 C.Preston",bg="sky blue",fg="black")
        infolabel.pack(side="top", fill="x", pady=10)
        infolabel2 = Label(popup, text="Program For Use Only By Those With Permission\nSend Questions/Comments/Bug Reports to\nCollin.Preston1@piedmont.org",bg="sky blue", font=fontstyle4)
        infolabel2.pack( fill="x", pady=10)
        popup.attributes("-topmost", True)
        popup.overrideredirect(True)
        B1 = Button(popup, text="Close",bg="azure", command = lambda: [about2(popup),about3()])
        B1.pack(side="bottom")
    else:
        pass


#Runs powershell command through a subprocess without console window. Returns an interger value.
def exchange():
    if e1.get() != "":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        output = subprocess.Popen(["powershell.exe", "Get-ADUser {} -Properties msExchRecipientTypeDetails|select msExchRecipientTypeDetails".format(e1.get())],universal_newlines=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE, stdin=subprocess.PIPE,startupinfo=startupinfo)
        result = str(output.communicate()[0].replace('msExchRecipientTypeDetails',"").replace('-',"").replace('\n',"").strip())
        output.terminate()

# If result is equal to 1 
        if result == "1":
            label.delete(0, END)
            label.insert(0,"Workspace_ONE_Boxer_users")
            return
# If result is equal to blank
        elif result == "":
            label.delete(0, END)
            label.insert(0,"User has no associated group")            
# If result is anything else
        else:
            label.delete(0, END)
            label.insert(0,"Workspace_ONE_Boxer_O365_Migrations")
            return
    else:
        failure = messagebox.showinfo("Error","Field Cannot Be Left Empty")
    return

#Runs powershell command through a subprocess without console window. Adds user to group in Active Directory
def boxer():
    cmds = []

    if label.get() != "":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        adadd = subprocess.Popen('powershell.exe Add-ADGroupMember -Identity {} -Members {}'.format(label.get(),e1.get().strip()),stdout=subprocess.PIPE,stderr=subprocess.PIPE, stdin=subprocess.PIPE,startupinfo=startupinfo)
        time.sleep(1)
        adadd2 = subprocess.Popen('powershell.exe Add-ADGroupMember -Identity {} -Members {}'.format(label.get(),e1.get().strip()),stdout=subprocess.PIPE,stderr=subprocess.PIPE, stdin=subprocess.PIPE,startupinfo=startupinfo)
        success = messagebox.showinfo("SUCCESS","User has been added to {}.".format(label.get()))
        adadd3 = subprocess.Popen('powershell.exe Add-ADGroupMember -Identity {} -Members {}'.format(label.get(),e1.get().strip()),stdout=subprocess.PIPE,stderr=subprocess.PIPE, stdin=subprocess.PIPE,startupinfo=startupinfo)
        adadd4 = subprocess.Popen('powershell.exe Add-ADGroupMember -Identity {} -Members {}'.format(label.get(),e1.get().strip()),stdout=subprocess.PIPE,stderr=subprocess.PIPE, stdin=subprocess.PIPE,startupinfo=startupinfo)
        adadd. terminate()
        adadd2. terminate()
        adadd3. terminate()
        adadd4. terminate()
        return
    else:
        failure = messagebox.showinfo("Error","Field Cannot Be Left Empty")        
        return
    return


#GUI Buttons for functions boxer and exchange
btn1 = Button(window, text="Get Exchange \nGroup", command = lambda: exchange())
btn1.place(x=90, y=80)
btn2 = Button(window, text="Add User to \nADGroup", command = lambda: boxer())
btn2.place(x=200, y=80)

#ICO
window.iconbitmap(default=img_resource_path('boxer.ico'))
label.place(x=35,y=160)

window.mainloop()
