# vbamacros
Here are some of my Excel VBA Macros which I've found to be helpful.

### Nonce_Values: 
A Macro that queries a blockchain API and returns the nonce values of any range of blocks you want, based on inputted block heights! _Note: this Macro only works on Windows computers, but a Mac version is on the way!_

#### Instructions:
1. Make sure to add two references: Microsoft WinHTTP Services and Microsoft Scripting Runtime. In the VBA editor on your Excel (ALT + F11), click 'Tools' in the upper-left hand corner, then 'References'. Scroll down to Microsoft WinHTTP Services, check the box, and click 'OK'. (It might be a bit of scrolling.) Then do the same for Microsoft Scripting Runtime.
2. Copy and paste the entire code (lines 1-165) into your VBA editor. To do this, press ALT + F11, then click 'Insert' in the upper left corner, then 'Module'. Now copy and paste this code into there.
3. Download the JsonConverter.bas file from https://github.com/VBA-tools/VBA-JSON, then drag the file into the "Modules" section of the VBA editor. It should add itself automatically as the JsonConverter module.
4. Click on the cell in your Excel where you want the table to be placed. (That cell will be the top-left-most cell in the table.)
5. Finally, run the Macro by pressing ALT+F8 at the same time. Click on Nonce_Values, then click Run! Follow the instructions from there.

This is my first Macro using an Error Handler, very exciting stuff.

### Charts_to_PPT: 
A Macro that takes any charts you've made in Excel and directly imports them, one chart per slide, into a PowerPoint Presentation. Easy instructions and customizable parts!

#### Instructions:
1. Since this Macro creates a PowerPoint and is not only limited to Excel, make sure to add a reference to the Microsoft PowerPoint library, or else this may not run properly. In the VBA editor on your Excel (ALT + F11), click 'Tools' in the upper-left hand corner, then 'References'. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and click 'OK'. (It might be a bit of scrolling.)
2. Copy and paste the entire code (lines 1-128) into your VBA editor. To do this, press ALT + F11, then click 'Insert' in the upper left corner, then 'Module'. Now copy and paste this code into there.
3. Finally, X out of the VBA editor and click back into your Excel workbook. Press ALT + F8, then click on 'Charts_to_PPT', and finally 'Run'. Follow the instruction prompts, and the Macro should create a PowerPoint!
4. If anything needs tweaking after the PPT has been made, either change it on the PPT, or if it won't let you, change it in Excel and run the Macro again.
5. Final message: If you want to spice up the slides, click 'Design' in the top PowerPoint banner, then 'Design Ideas'. This will sometimes suggest a few beautiful layouts for your slide.

The most common error, I'm guessing, will be a 'User defined type' error. This just means that you have to redo Step 1: enable the reference. (ALT + F11 > Tools > References > check the PowerPoint box > OK)

Hope this can save you some time!

### More Macros on the way...
Stay tuned!

### Final notes:
If there are any other errors, potential improvements, or anything else you'd like to see, please let me know! Criticism is the best way to learn, and I appreciate it greatly. Thanks!

#### -Reilly
