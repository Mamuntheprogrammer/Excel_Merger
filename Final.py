from tkinter import *
import tkinter as tk
from tkinter.font import Font
import webbrowser
from tkinter import ttk
from tkinter import filedialog,messagebox
import time
import glob
import os
import pandas as pd
from sys import exit
from pandas import ExcelWriter
x=''



#main inheret classes
#----------------------------------------------------------------
win =tk.Tk()
main_menu=tk.Menu(win)

#-----------------main gui title-----------------
win.title("Excel Merger")

wi_gui=700
hi_gui=370

wi_scr=win.winfo_screenwidth()
hi_scr=win.winfo_screenheight()

x=(wi_scr/2)-(wi_gui/2)
y=(hi_scr/2)-(hi_gui/2)

win.geometry('%dx%d+%d+%d'%(wi_gui,hi_gui,x,y))

win.iconbitmap(r'C:\Users\HujurHacker\Desktop\Final gui\images\xlsx.ico')


#----------------- Main gui size -------------------

#-------------------------------All functions------------------------------------
#---- progress bar -----
max=100
step=tk.DoubleVar()
step.set(0)
progbar=ttk.Progressbar
st=0

def add_progbar():
	global progbar
	progbar=ttk.Progressbar(
		pframe,
		orient=tk.HORIZONTAL,
		mode='determinate',
		variable=step,
		maximum=max,
		length=600
		)
	progbar.pack(fill=X,expand=True)


#----- FOR MENU ----
#---------Button for select directory-------------
def daction():
	entry_d.delete(0, 'end')
	
	folder_selected = filedialog.askdirectory(initialdir="/",title='Please select a directory')
	if not folder_selected:
		folder_selected=entry_d_var.get()
	else:
		entry_d.insert(0,folder_selected)
	try:
		pp=os.chdir(str(folder_selected))
	except:
		messagebox.showerror("Error", "Empty or wrong Directory")

def empt():
	button_m.configure(state=DISABLED)


def mfunc():
	
	btn_txt.set("Merging")
	pattern ='*.xlsx'
	xllis=glob.glob(pattern)

		

	#xllis = os.listdir(pp)
	xllis.sort()
	file_identifier = "*.xlsx"
	if not xllis:
		messagebox.showerror("Error", "Wrong directory or There is no XLSX file found")
		btn_txt.set("Merge")
	else:
		df2 = pd.DataFrame()

		add_progbar()

		global progbar
		
		pa=selct_typ1.get()
		
		button_m.configure(state=DISABLED)
		for x in range(len(xllis)):

			if pa==1 or pa==3:
				df = pd.read_excel(xllis[x])
			else:
				df = pd.read_excel(xllis[x],header=None)
			
			for y in range(0,len(df)):
				
				step.set(y)
				time.sleep(.00000001)
				win.update()
				df2=df2.append(df.iloc[[y,],:])
			df.iloc[0:0]
		writer = ExcelWriter('Merged-'+str(x)+'.xlsx', engine='xlsxwriter')

		if pa==1:
			df2.to_excel(writer,sheet_name='merged',index=False,header=None)
		elif pa==3:
			df2.to_excel(writer,sheet_name='merged',index=False)

		else:
			df2.to_excel(writer,sheet_name='merged',index=False,header=None)
		 
		writer.save()
		progbar.destroy()
		messagebox.showinfo("Information","Merge Complete")
		btn_txt.set("Merge")
		path=os.getcwd()
		button_m.configure(state=NORMAL)
		webbrowser.open(path)
		

	

#---- FOR LINK ----
def opnlink(url):
    webbrowser.open_new(url)

#developer
def dev():
	nwin = Toplevel()
	wi_gui=480
	hi_gui=300

	wi_scr=nwin.winfo_screenwidth()
	hi_scr=nwin.winfo_screenheight()

	x=(wi_scr/2)-(wi_gui/2)
	y=(hi_scr/2)-(hi_gui/2)

	nwin.geometry('%dx%d+%d+%d'%(wi_gui,hi_gui,x,y))

	nwin.iconbitmap(r'C:\Users\HujurHacker\Desktop\Final gui\images\xlsx.ico')
	nwin.title("Developer")
	photo2 = PhotoImage(file=r'C:\Users\HujurHacker\Desktop\Final gui\images\developer.png')
	lbl2 = Label(nwin, image = photo2)
	lbl2.pack()
	nwin.mainloop()
	

def sft():
	nwin = Toplevel()
	wi_gui=400
	hi_gui=144

	wi_scr=nwin.winfo_screenwidth()
	hi_scr=nwin.winfo_screenheight()

	x=(wi_scr/2)-(wi_gui/2)
	y=(hi_scr/2)-(hi_gui/2)

	nwin.geometry('%dx%d+%d+%d'%(wi_gui,hi_gui,x,y))

	nwin.iconbitmap(r'C:\Users\HujurHacker\Desktop\Final gui\images\xlsx.ico')
	nwin.title("PyGems Excel Merger")
	photo2 = PhotoImage(file=r'C:\Users\HujurHacker\Desktop\Final gui\images\software.png')
	lbl2 = Label(nwin, image = photo2)
	lbl2.pack()
	nwin.mainloop()
def rme():
	nwin = Toplevel()
	wi_gui=400
	hi_gui=144

	wi_scr=nwin.winfo_screenwidth()
	hi_scr=nwin.winfo_screenheight()

	x=(wi_scr/2)-(wi_gui/2)
	y=(hi_scr/2)-(hi_gui/2)

	nwin.geometry('%dx%d+%d+%d'%(wi_gui,hi_gui,x,y))

	nwin.iconbitmap(r'C:\Users\HujurHacker\Desktop\Final gui\images\xlsx.ico')
	nwin.title("PyGems Excel Merger")
	photo2 = PhotoImage(file=r'C:\Users\HujurHacker\Desktop\Final gui\images\software.png')
	lbl2 = Label(nwin, image = photo2)
	lbl2.pack()
	nwin.mainloop()

def abt():
	opnlink("http://www.pygems.com")
def tuto():
	opnlink("http://www.pygems.com")
def tip():
	nwin = Toplevel()
	wi_gui=600
	hi_gui=480
	wi_scr=nwin.winfo_screenwidth()
	hi_scr=nwin.winfo_screenheight()

	x=(wi_scr/2)-(wi_gui/2)
	y=(hi_scr/2)-(hi_gui/2)

	nwin.geometry('%dx%d+%d+%d'%(wi_gui,hi_gui,x,y))

	nwin.iconbitmap(r'C:\Users\HujurHacker\Desktop\Final gui\images\xlsx.ico')
	nwin.title("Tips")
	photo2 = PhotoImage(file=r'C:\Users\HujurHacker\Desktop\Final gui\images\tips.png')
	lbl2 = Label(nwin, image = photo2)
	lbl2.pack()
	nwin.mainloop()


def pdmrgr():
	opnlink("http://www.pygems.com")
def pdimg():
	opnlink("http://www.pygems.com")
def xlsplt():
	opnlink("http://www.pygems.com")
def rnam():
	opnlink("http://www.pygems.com")
def pdsplt():
	opnlink("http://www.pygems.com")


def pygems():
	messagebox.showinfo("Pygems Excel Merger","Merge Complete")


#--------------ALL FRAMES WILL BE HERE-----------------------
rframe=Frame(win)
dframe=Frame(win)

#frame for merge button
mframe = Frame(win,	bd=0)

#--- progress bar 
pframe=Frame(win)


#------------------- FRAME CONTENT ----------------------

label_1=Label(text="PyGems Excel Merger",
	bd=0,
	bg="#393e46",
	fg='#F4511E',
	font='Times 20',
	width=0,
	height=0	
	)

label_2_font=Font(family='Times',size=8,underline=1)

label_2=Label(text="WWW.PYGEMS.COM",
	bd=1,
	font=label_2_font,
	bg='#393e46',
	fg='#FFEB3B',
	width=0,
	height=0,
	cursor="hand2",
	)

statusbar =Label(win, text="Click here to visit : www.pygems.com",
 bd=1,
  relief=SUNKEN,
   bg="#37474F",
   fg='#fcf9ec',
   height=2,
   font="Times 13",
   cursor="hand2"
   )

#-----------Radio Button All----------- 


selct_typ1=tk.IntVar()
selct_typ1.set(2)


radiobtn1 = ttk.Radiobutton(rframe,text="Ignore Header" ,value=1,variable=selct_typ1)

radiobtn2 = ttk.Radiobutton(rframe,text="Default" ,value=2,variable=selct_typ1)

radiobtn3 = ttk.Radiobutton(rframe,text="Same Header" ,value=3,variable=selct_typ1)


qbutton=Button(rframe,command=tip,cursor="hand2")
photo=PhotoImage(file=r'C:\Users\HujurHacker\Desktop\Final gui\images\question.png')


#------------- directory entry------------
entry_d_var = StringVar()
entry_d=Entry(dframe,width=80,textvariable=entry_d_var,bg='#dedede')
entry_d_txt = entry_d_var.get()


#------------directory Button----------
button_d=tk.Button(dframe,relief=RAISED,font=('Times 10 bold'),text='Select Folder' ,fg='#fcf9ec',bg='#132238',command=daction)


#----------merge button-------------------
btn_txt=StringVar()
button_m=tk.Button(mframe,textvariable=btn_txt,command=mfunc,relief=GROOVE,font=('Times 10 bold'),width=22,fg='#fcf9ec',bg='#132238')
btn_txt.set("Merge")

#-------------------- ALL PACK ------------------


label_1.pack(fill=X)
label_2.bind("<Button-1>", lambda e: opnlink("http://www.pygems.com"))
label_2.pack(fill=X)

#statusbar pack

statusbar.bind("<Button-1>", lambda e: opnlink("http://www.pygems.com"))
statusbar.pack(side=BOTTOM, fill=X)


#radio pack

radiobtn1.pack(side=LEFT,padx=20)
radiobtn2.pack(side=LEFT,padx=20)
radiobtn3.pack(side=LEFT,padx=20)

qbutton.pack(side=LEFT,ipadx=3,ipady=3)

entry_d.pack(ipady=4,side=LEFT,pady=13)
entry_d.focus()

button_d.pack(side=LEFT,padx=10,ipady=2,pady=13)
button_m.pack(pady=20)



#frame pack
#radio button pack
rframe.pack(pady=15)

#directory entry pack
dframe.pack(padx=0)

#merge button pack
mframe.pack(pady=0)


#progress bar
pframe.pack(pady=5)

#----------- Menu --------------------------------------------------------------

#About menu
about_menu=tk.Menu(main_menu,tearoff=0)
about_menu.add_command(label='Developer',command=dev)
about_menu.add_command(label='Software',command=sft)

#Contact menu
cntct_menu=tk.Menu(win,tearoff=0)
cntct_menu.add_command(label='About Us',command=abt)

#help menu
hlp_menu=tk.Menu(win,tearoff=0)
hlp_menu.add_command(label='Tutorial',command=tuto)
hlp_menu.add_command(label='Tips',command=tip)

#readme_menu
rd_menu=tk.Menu(win,tearoff=0)
rd_menu.add_command(label='Read Me',command=tip)

#more tools
mrtools=tk.Menu(win,tearoff=0)
mrtools.add_command(label='Pdf Merger V-1.0.1',command=pdmrgr)
mrtools.add_command(label='File Renamer V-1.0.1',command=rnam)
mrtools.add_command(label='Excel Spliter V-1.0.1',command=xlsplt)
mrtools.add_command(label='Pdf Spliter V-1.0.1',command=pdsplt)
mrtools.add_command(label='Click for More Tools',command=mfunc)


#menu cascade
main_menu.add_cascade(label="About",menu=about_menu)
main_menu.add_cascade(label='Contact',menu=cntct_menu)
main_menu.add_cascade(label='Help',menu=hlp_menu)
main_menu.add_cascade(label='More Tools',menu=mrtools)
main_menu.add_cascade(label='Read Me First',menu=rd_menu)


#--------------------------------- configure -------------------------------------
win.config(menu=main_menu)
win.configure(bg='#393e46')

#questionbutton
qbutton.config(image=photo,width='30',height='30',bg='#132238')

#directory entry
dframe.config(bg='#393e46')
mframe.config(bg='#393e46')
#-------------------- mainloop to open window always -----------------
win.mainloop()





