RFD Python Automation Guidance
Contents: 
•	Installing
•	Running
•	Updating


Installing:
1.	Install Microsoft Visual Studio (VS) Code (from Microsoft store)
2.	Install Python version 3.13 or higher (from Microsoft store)
3.	Open the code file in visual studio code.
4.	Open the terminal in visual studio code (see picture below) and 1 by 1 paste the following instructions in, then press enter:
a.	pip install pandas
b.	pip install xlwings
 
 

Running:
Update the file locations by copying the full path from File Explorer:
•	the path to the excel source file is on line 450 (called source file)
•	the path to the excel template is on line 454 (called template file)
•	the path to the folder you want the PDF’s to be saved to is on line 458 (called output folder)
•	MAKE SURE TO KEEP THE ‘r’ AT THE FRONT OF THE PATH

Close all excel files (you can continue using other applications while running the code, but if you have the excel file that you want to run during the code open when you start the code this will lead to an issue
The code should take 10-15 minutes to run, it will provide updates on its progress in the terminal as it runs, if there is no sign of progress, you may want to use control C to cancel the running of the code and to start again
If you get an error, there is copilot in visual studio code, or paste the entirety of the code and the entirety of the error into ChatGPT (or Claude – a better version of ChatGPT), ask it why you are getting this error and how to correct it, 9/10 it will be able to solve this.
Make sure once the code has finished running to go back and check for errors in specific files. If the code finishes running fine, but there are specific lines that have caused an error, you probably don’t need to change the code. It is likely to be an issue in excel. If there is no error in excel and you try running the code again and it still doesn’t work for a specific line, you may need to manually create the RFD for that file. 


 
Updating:
If the template updates or changes materially, you can simply update where data is being drawn from or sent to:
Data is drawn from the initial excel document in rows 74 – 93. Update the column references, or add/remove particular rows with their column reference if new areas of the template need to be filled

Data is then inserted into the template in rows 220 – 238. Follow the same process to update this information as you did for updating where the data was drawn from.
 
If you want to update the way the files are named, this is done on line 192 - 195
 
