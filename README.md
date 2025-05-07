# Schedule Maker

## Overview

The **Schedule Maker** is a Python program that automates the transformation of appointment data into a customizable schedule. Originally developed during my assistant manager role at a tutoring center, this tool was designed to save time and reduce manual work, enabling us to focus more on what mattered most: teaching students.

## The Beginning: Automating a Repetitive Task

This program started as a solution to a repetitive task I faced daily as an assistant manager at a tutoring center. My job involved manually transforming appointment data into a schedule format that could be used by our team. It was tedious and time-consuming. I decided to automate the process.

At the time, I had no formal programming experience, but I was determined to find a solution. I wrote my very first useful program that took an output file from our scheduling system (Appointy) and transformed it into an Excel sheet.

### Key Features in the Beginning:
- **Automated CSV to Excel Conversion**: Turned CSV data into a schedule format.
- **Flexibility**: Allowed our team to handle the scheduling more efficiently, reducing errors.

## Iterating and Adding GUI Features

As time went on, I realized that the program needed to be updated. My boss asked for revisions as the schedule formatting changed, and I also saw that adding a graphical user interface (GUI) could make it even easier for the team to use.

- **Revisions**: Over the course of time, I updated the program multiple times to adapt to the changing requirements of the schedule.
- **Adding a GUI**: I added a GUI that made it easier for the team to interact with the program without writing any code. It was a huge improvement, but the code was still rigid, messy, and difficult to maintain.

## The Turning Point: A New Understanding

After 2 years of working part-time at the tutoring center, I was asked to revise the tool before I left for good. By this point, I had gained more experience in programming and a much deeper understanding of Python, especially with libraries like Pandas. My earlier code, which had become inefficient and ugly, needed to be reworked.

- **Reflection on the Old Code**: Looking back, the code was a mess! I had used Excel’s built-in functions and manual string manipulations to process the data. I had created a rigid structure of transformations and dictionary-based cleaning that was hard to modify and scale.
  
- **The Decision to Use Pandas**: With my newfound knowledge, I realized I could simplify everything by doing all the transformations in Pandas, rather than relying on Excel or complicated logic. The code would be shorter, more efficient, and more flexible.

## Rewriting the Code: From 700 Lines to 350

With this new perspective, I rewrote the program in just **350 lines of code** (down from 700+). I made the following improvements:
- **Dynamic Transformations**: The transformations now work in a dynamic way, which means the program can adjust more easily to changing data formats.
- **No More Weird Logic**: I eliminated the old, confusing logic and replaced it with cleaner, more efficient Pandas code.
- **Simplicity and Maintainability**: The new version of the program is far more maintainable and easier to update, ensuring that it can adapt to future changes.

## The Final Version

The updated program does exactly what it’s supposed to: **it takes the output from Appointy (a scheduling system), processes the CSV data, and creates an Excel sheet that resembles Appointy’s UI**. This feature was especially important during the COVID-19 pandemic when we needed to:
- **Reorder cells**: So we could easily assign students to desks and teachers.
- **Adapt quickly to changes**: As the format of the schedule was updated regularly due to changing circumstances.

### Key Features:
- **Customizable Scheduling**: Allows for easy reordering of the schedule data.
- **Excel Output**: Outputs data in an Excel format that closely resembles the Appointy UI, making it easy to view and manipulate.
- **Dynamic and Scalable**: The program is flexible enough to accommodate changes to the scheduling system without requiring major rewrites.

## Lessons Learned

Looking back at the journey of this project, I’ve learned a lot about:
- **Improving code efficiency**: By leveraging powerful libraries like Pandas, I was able to reduce the amount of code and make it much more efficient.
- **Adapting to change**: The program evolved from a simple script to a full-featured, dynamic tool. This is a great reminder of how important it is to be flexible and keep learning.
- **Automation’s Impact**: The impact of automating repetitive tasks cannot be understated. It made our work much easier and allowed us to focus on what really mattered—helping students.

## Conclusion

This project represents more than just an automation tool—it’s a testament to how much I’ve grown as a programmer. What started as a simple program to help with a tedious task turned into a dynamic and valuable tool that continues to help my former team. 

I’m proud to have been able to leave a parting gift that will save time and reduce manual work for others. This project will always remind me of my journey in programming and the power of learning and adapting.
