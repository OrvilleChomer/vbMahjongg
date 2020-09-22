# vbMahjongg

This is the source code for an old Mahjongg Solitaire game that I wrote in Visual Basic which I provided as *freeware* which ran on Windows machines. 
I think it was even included on one of those floppy disks back in the day that contained a collection of freeware written by various authors! That's pretty cool!
I got a lot of nice comments from people over the years about the game.

**Screen Shot of Original Game Below:**

![Screen Shot](http://chomer.com/wp-content/mahjongg_screen_full1.png)


It was thorougly debugged, and I played many many delightful games on it myself! It had some limitations, for example the window size was fixed, and was for the screen resolution I was running in at the time I wrote it. This resolution was much lower than screen resolutions are today!  But it was still playable!

I have a sneaking suspicion that the original executable will not run on a modern version of Windows without some sort of adapter code or such thing.

To give you an idea how old this code is, I believe it was written in Visual Basic 5 Professional Edition. It compiled in the Visual Basic IDE in an executable file.

Notice also in the screen shot above how I never went to the trouble to create custom icons for the buttons, or a custom icon for the app itself. Sigh!

I am currently working on and off on a *web version* which I plan to have running on CodePen and Glitch.com! I *may* do an SVG version of the tile images, or I might stick to using just bitmap images.

- *VBMahjongg.vbp* is the project file.
- The *clsMahjongg.cls* VB class file contains the main game logic.
  - The `ShuffleTiles()` function is very interesting to me. It's logic provides a quick way to shuffle the contents of an array. Rewriting the core logic of this function in other programming languages could be a very useful piece of code indeed! A more generic name for the routine like: `shuffleArray()` would probably be good.
- *frmMain.frm* and *frmMain.frx* are the files that define the *main window* of the application. The `frm` file is a text file defining different controls on the form and their property values.  The `frx` file contains the binary data for the form. Stuff like bitmap image data. This includes a bitmap for all the different possible tile images there are, as well as a background images (which were photos I actually took at the Japanese Garden located in the Chicago Botanic Gardens)!
- The game had sound effects which could be heard by the code playing `wav` files.
  - There was no native way in VB to play wav files (or any other format for audio files), so I used a Window API call to do it. This code was in a wrapper function called: `PlayWav()` which is in the *clsMahjongg.cls* file. You pass it the name of the `wav` file and it takes care of the rest!
