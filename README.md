# Schedule Maker

## Overview

The **Schedule Maker** is a Python program that automates the transformation of appointment data (from Appointy) into a customizable schedule. Originally developed during my assistant manager role at a tutoring center, this tool was designed to save time and reduce manual work, enabling us to focus more on what mattered most: teaching students.

## The Beginning: Automating a Repetitive Task

I had to create a schedule at the end of each week which consisted of me manually inputing and colouring cells in an excel spreadsheet. It took around three hours. 

At the time, I had no formal programming experience, but I had an inkling that I could use Python.

## Overall Idea

- **Overall Idea**: Coverting Raw Data to a cool looking excel spreadsheet

## Iterating and Adding GUI Features

As time went on, I realized that the program needed to be updated. My boss asked for revisions as the schedule formatting changed, and I also saw that adding a graphical user interface (GUI) could make it even easier for the team to use.

- **Revisions**: Over the course of time, I updated the program multiple times to adapt to the changing requirements of the schedule.
- **Adding a GUI**: I added a GUI that made it easier for the team to interact with the program without writing any code. It was a huge improvement, but the code was still rigid, messy, and difficult to maintain.

## The Turning Point: A New Understanding

After 2 years of working part-time at the tutoring center, I was asked to revise the tool before I left for good. By this point, I had gained more experience in programming and a much deeper understanding of Python, especially with libraries like Pandas. My earlier code, which had become inefficient and ugly, needed to be reworked.


Looking back, the code was actually horrible, down to the logic choices I had made. Why on earth would I use dictionaries and transfrom strings in the csv files instead of using pandas and creating a dataframe? It accomodated excels infamous date format conversions, but it really didn't need to. I should have read the file and then wrote a new one, not doing some weird gymnastics on the file I had at hand. Also I had some really convoluted logic that created the columns, it really makes me wanna throw up. Anywaysssss.

  
With gaining a brain cell, I realized I could simplify everything by doing all the transformations in Pandas, rather than relying on Excel or complicated logic. The code would be shorter, more efficient, and more flexible.

## Rewriting the Code: From 700 Lines to 350

With this new brain cell, I rewrote the program in just **350 lines of code** (down from 700+). I made the following improvements:

The transformations now work in a dynamic way, which means the program can adjust more easily to changing data formats. It in a sense is more rigid. It would place down a students name at a corresponding time, not calculating the cell which it should be in which sounds crazy but thats how I did it the first way. 

I eliminated the old, confusing logic and replaced it with cleaner, more efficient Pandas code.


The new version of the program is far more maintainable and easier to update, ensuring that it can adapt to future changes.

## The Final Version

The updated program does exactly what it’s supposed to: **it takes the output from Appointy (a scheduling system), processes the CSV data, and creates an Excel sheet that resembles Appointy’s UI**. This feature was especially important during the COVID-19 pandemic,


### Key Features:
- **Automated CSV to Excel Conversion**: Turned CSV data into a schedule format.
- **Flexibility**: Allowed our team to handle the scheduling more efficiently, reducing errors.
- **Added a Graphical User Inserface**: Used Tinker for better user friendly experience
- **Used a more Dynamical Logic**: Better for frequent changes

## Lessons Learned

Looking back at the journey of this project, I’ve learned a lot about:
- **Improving code efficiency**: By leveraging powerful libraries like Pandas, I was able to reduce the amount of code and make it much more efficient.
- **Adapting to change**: I changed as a person. A person who creates garabage to someone who can find the more logical step in a problem. I fundamentally approach coding differently now
- **Automation’s Impact**: The impact of automating repetitive tasks cannot be understated. It made our work much easier and allowed us to focus on what really mattered - helping students. I think coding for effciency is an art. Like even figuring out which tasks can be automated is a skill that can be attained, and I feel like its never taught. Like the short version is, if anything, litterally anything on the computer at least is repetitive in some way, you can automated it period. My hatred for repretitive tasks fueled this understanding.

## Conclusion

This project represents more than just an automation tool-it’s a testament to how much I’ve grown as a programmer. What started as a simple program to help with a tedious task turned into a dynamic and valuable tool that continues to help my former team. 

I’m proud to have been able to leave a parting gift that will save time and reduce manual work for others. This project will always remind me of my journey in programming and the power of learning and adapting. I'm so proud of my newly attained brain cell

Thanks for reading this
