# vbMahjongg
This is the source code for an old Mahjongg Solitaire game that I wrote in Visual Basic which I provided as *freeware* which ran on Windows machines. 
I think it was even included in one of those floppy disks back in the day that contained a collection of freeware! That's pretty cool!
I got a lot of nice comments from people over the years about the game.

It was thorougly debugged, and I played many many delightful games on it myself! It had some limitations, for example the window size was fixed, and was for the screen resolution I was running in at the time I wrote it. This resolution was much lower than screen resolutions are today!  But it was still playable!

I have a sneaking suspicion that the original executable will not run on a modern version of Windows without some sort of adapter code or such thing.

To give you an idea how old this code is, I believe it was written in Visual Basic 5 Professional Edition. It compiled in the Visual Basic IDE in an executable file.

I am currently working on and off on a web version which I plan to have running on CodePen and Glitch.com!

- *VBMahjongg.vbp* is the project file.
- The *clsMahjongg.cls* VB class file contains the main game logic.
  - The `ShuffleTiles()` function is very interesting to me. It's logic provides a quick way to shuffle the contents of an array. Rewriting the core logic of this function in other programming languages could be a very useful piece of code indeed!
- The game had sound effect which could be heard by the code playing `wav` files.
  - There was no native way in VB to play wav files, so I used a Window API call to do it. This code was in a wrapper function called: `PlayWav()` which is in the *clsMahjongg.cls* file.
