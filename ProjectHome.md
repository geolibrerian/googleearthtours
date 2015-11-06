This script is a GUI program that allows the user to create a tour in Earth, which can then be recorded using the built-in tour tool. The advantage of this tool is that, by balancing the pause and the camera speed, you can create a smooth and swooping tour, better than by hand.

To start, enter a file path and a file name (with the .py extension, of course) into the top entry box and push File Create or Open; you can enter the name of an existing file to append to it. Then, start taking 'snapshots' of the areas you want to highlight. Be sure to adjust the pause to make it linger in an area longer, and the speed to make fly faster/slower between areas. By adjusting these two scale bars, you can make it more or less smooth.

You need Mark Hammond's win32com module, available by downloading PythonWin or installing the module (available here: http://sourceforge.net/projects/pywin32/ ), if you are not installing it using the MSI.

PS, this was a very early script of mine, so please ignore the crappy variable names and noodle code. It works, and that's that best part.


Oh, and of course you need Python!

I have no association with Google or Google Earth, though I love their product!

PostScript: Because this program uses COM interaction with Google Earth, which is being phased out, you might not be able to use the program with Google Earth versions beyond 5.2