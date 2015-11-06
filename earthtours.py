# This Script allows one to create a script that records a tour in Google Earth
# It makes control of the Eye smoother and easier then hand controls, making it a
# better tool for controlling the Eye when recording a Tour.
# If you want kml, use this to record a tour, start the script, then record the tour using the built in tool.
# Silas Toms, GIS Technician 3/09
# Official Version 1


#!/usr/bin/python

import Tkinter, sys , os
from Tkinter import *
import time
import win32com.client           #you need to install Mark Hammond's PythonWin, or at least the win32com module, for this script to work

g = 'Enter File Name and Path'
n = 'Pause between Views'
header = '''import time, win32com.client\n
ge = win32com.client.Dispatch('GoogleEarth.ApplicationGE')\n

    #http://earth.google.com/comapi/interfaceIApplicationGE.html
    #SetCameraParams ([in] double lat,[in] double lon,[in] double alt,[in] AltitudeModeGE altMode,
    #[in] double range,[in] double tilt,[in] double azimuth,[in] double speed)'''
x = 50
s = 1
b = 10



def openge(): #this function opens Google Earth
	ge = win32com.client.Dispatch("GoogleEarth.ApplicationGE")

def tour(): # this function gets the necessary information from Google Earth and writes it to a script
	
	ge = win32com.client.Dispatch("GoogleEarth.ApplicationGE")
	lat = float(ge.GetCamera(0).FocusPointLatitude)
	long = float(ge.GetCamera(0).FocusPointLongitude)
	alt = float(ge.GetCamera(0).FocusPointAltitude)
	mode = 1
	range = float(ge.GetCamera(0).Range)
	tilt = float(ge.GetCamera(0).Tilt)
	az = ge.GetCamera(0).Azimuth

	k = speed(s)
	param = "\nge.SetCameraParams(%f,%f,%f,%d,%f,%f, %f, %f)\n" % (lat, long, alt, mode, range, tilt, az, k)
	pause1 = "time.sleep(%f)" % (float(pause(b)))
	fo = filewrite()
	if fo != 'fo':	
		fo.writelines(param)
		fo.writelines(pause1)
		fo.close()
	else:
		m = 'Set File First'
		v.set(m)


def quit():
	root.destroy()



def filewrite(): # This function either creates a new script or opens an old one
	try:
		fl= e.get()
		fs = fl.split('.')
		fc = fs[0].count('\\')
		fc1 = fs[0].count('/')
		f = fc + fc1
		if os.path.exists(fl) and fs[-1] == 'py':
			fo = open(fl,'a')
		elif  f == 0 and fs[-1] != '': 
			if os.path.exists('C:')and fs[-1]=='py':
				fl = 'C:/' + fl
				fo = open(fl,'a')
				fo.write(header)
			elif os.path.exists('C:')and fs[-1]!='py':
				fl = 'C:/' + fl + '.py'
				fo = open(fl,'a')
				fo.write(header)
			else:
				w = 'Enter Absolute File Path'
				v.set(w)
				e.delete(0)
				fo = 'fo'
				return fo
		
		elif fl == '':
			w = 'No File Path Entered'
			v.set(w)
			e.delete(0)
			fo = 'fo'
			return fo
		else:
			if fs[-1] == 'py':
				fo = open(fl,'a')
				fo.write(header)
			else:
				fl = fs[0] + '.py'
				fo = open(fl,'a')
				fo.write(header)
		o = 'File: %s' % fl #(e.get())
		v.set(o)
		e.delete(0, len(e.get()))
		e.insert(0,fl)
		return fo
	except:
		o = 'Re-enter File Path'
		v.set(o)
	
def pause(b): # this function adjusts the time.sleep function in the created script, making the time between views longer or shorter
	
	p = int(b.get())
	return p

def speed(s):
	sp = s.get()
	s = (float(sp)/2) /10
	return s

def execute(): # this function runs the script that has been created, if an existing script has been entered
	fl = e.get()
	
	if os.path.exists(fl):
		execfile(fl)
		w = 'Running the script'
		v.set(w)
	
	else:
		w = 'No Script Created'
		v.set(w)

root =Tk()
root.title('Google Earth Tours')
root.geometry('240x370')
root.geometry('+250+70')
icon = 'example.ico'  # you can put an icon's path in this string and map the GUI prettier. It replaces the default Tkinter icon. 
if os.path.exists(icon):
	root.iconbitmap(icon)

e = Entry(root, width=25)
e.pack(side=TOP, padx=0, pady =5)
e.focus_set()
v= StringVar()
Label (textvariable=v,
	   font = 'Helvetica -13 bold',
	   fg = 'dark blue').pack(side=TOP,padx=10,pady=2)
v.set(g)

Button(root, text="File Create or Open", width=20,
	   font = 'Helvetica -12 bold', bg = 'white', fg=  'dark blue',
	   command=filewrite).pack(side=TOP, padx =x,pady=2)

separator = Frame(height=6, bd=1, relief=SUNKEN)
separator.pack(fill=X, padx=5, pady=5)
Button(root, text='Open Google Earth', width=20,
	   font = 'Helvetica -12 bold', bg = 'white',
	   fg=  'dark blue',command= openge).pack(side=TOP, padx =x,pady=2)
separator = Frame(height=6, bd=1,relief=SUNKEN)
separator.pack(fill=X, padx=5, pady=5)
Button(root, text='Run Script',width=20,
	   font = 'Helvetica -12 bold', bg = 'white',
	   fg=  'dark blue',command=execute).pack(side=TOP,padx=x,pady=2)


separator = Frame(height=6, bd=1, relief=SUNKEN)
separator.pack(fill=X, padx=5, pady=5)
Button(root, text='Record View',width=20,height =2,
	   font = 'Helvetica -17 bold', bg = 'white',
	   fg=  'dark blue', command= tour).pack(side=TOP, padx =x)


b= IntVar()
s = IntVar()
Scale(root, variable = b, width = 20, length = 100,
	  orient = 'vertical', from_ = 5, to = 30, resolution = 1.5,
	  tickinterval = 5,font = 'Helvetica -11 bold', fg = 'dark blue',
	  label = 'Pause', command = pause(b)).pack(side= LEFT)
Scale(root, variable = s, width = 20, length = 100,
	  orient = 'vertical', from_ = 1, to = 11, resolution = 1,
	  tickinterval = 2, font = 'Helvetica -11 bold', fg = 'dark blue',
	  label = "Speed", command = speed(s)).pack(side= RIGHT)


root.wm_attributes("-topmost", 1)


root.mainloop()




