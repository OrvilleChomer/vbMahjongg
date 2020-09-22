# vbMahjongg
This is the source code for an old Mahjongg Solitaire game that I wrote in Visual Basic which I provided as *freeware* which ran on Windows machines. 
To give you an idea how old this code is, I believe it was written in Visual Basic 5 Professional Edition. It compiled in the Visual Basic IDE in an executable file.

- *VBMahjongg.vbp* is the project file.
- The *clsMahjongg.cls* VB class file contains the main game logic.
  - The `ShuffleTiles()` function is very interesting to me. It's logic provides a quick way to shuffle the contents of an array. Rewriting the core logic of this function in other programming languages could be a very useful piece of code indeed!
- The game had sound effect which could be heard by the code playing `wav` files.
  - There was no native way in VB to play wav files, so I used a Window API call to do it. This code was in a wrapper function called: `PlayWav()`.
