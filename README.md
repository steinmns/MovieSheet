# MovieSheet
VBA enabled spreadsheet that tracks and analyzes movie watching trends over time.

# Repository Format
The latest copy of the entire workbook (EntertainmentLogGraphing - DevCopy.xlsm) is available above as well as the code for the excel macros themselves (MacroCodeV1) and the code for the buttons that call the macros (ButtonCodeV1).

# Operation
The workbook is comprised of five different worksheets:
  1. Movie Log: This is the bread and butter of the sheet. Movies are entered here with columns for the title, date watched, genre, rating      comments the viewer had, whether it should be re-watched, and whether it was watched in a theater or at home. This data is then            visualized in the "Graphs and Stats" worksheet.
  
  2. Movies to Watch: This worksheet is currently used to record movies the user wants to watch based on title, genre, and where the movie      is available (theaters/streaming services).
  
  3. TV Log: The "TV Log" worksheet is very similar to the "Movie Log" worksheet. The primary differences are that the "TV Log" tracks          which episode was completed last and does not have columns for re-watch or location of viewing. The "TV Log" data is not currently        accessed in the "Graphs and Stats" worksheet.
  
  4. TV Shows to Watch: Same as "Movies to Watch" worksheet.

  5. Graphs and Stats: The "Graphs and Stats" worksheet is where the magic happens. In the worksheet, the user is able to see how many          movies of each genre that they have watched and have it also displayed as a pie chart, see their average movie rating, their highest      and lowest movie ratings, how many movies they watched in theaters or at home, and the total amount of movies they've watched each        year since 2016 (when the workbook was first created). To gather data on movie genres watched and create the pie chart, click the          button labeled "Generate Pie Chart." To gather and display all other statistics, click the button labeled "Misc. Stats."

# Development
The code for the workbook is comprised of five macros and two button subs. To access the macros, navigate to the  "View" tab in excel and select "Macros" in the far right. This will bring up the list of macros where the user then has access to edit or execute them. 


# Future Improvements/Ideas (Subject to change)
 -User form data entry
 
 -Duplicate entry warnings
 
 -Better filtering
 
