# Outline:

- A time logging productivity tool that helps you log time spent on activities, storing data in an excel sheet through openpyxl and analysing the data with matplotlib. A more detailed description, alongside sample outputs lies below.

# Purpose:

- Productivity is something a lot, if not most of us struggle with - myself included - and therefore when my friend approached me and asked me to build him a productivity tool, I was extremely intrigued. I wanted to build a minimal yet useful product that I would end up utilising myself and so I arrived at my build - a means to log time spent on different activities throughout the day, which can be very effectively visualised using an excel sheet where the rows are labelled to different times, by default 07:00 to 24:00 in 15 minute increments, and each column is a particular day. The build also helps analyse the time logged - showing you how much of your day you have logged and what have you logged the most as well as a pie chart of all the logged data.
- Building this system was wholly engaging and challenging - teaching me not only how to better work with modules such as openpyxl and matplotlib but also, and in my opinion more importantly, deciding how to keenly store data and reuse it in the most efficient way, alongside how to make the entire program more usable by allowing the users to make mistakes and explicitly handling errors. I loved programming this project.

# Description:

- First, while looking at the inner workings of the program, it is necessary to also visualise the final product, and see the program as it runs, hence I have included several snapshots of runtime in a folder in this repository.
- In this build, the initialise_logger program is meant to be run first and only needs to be run once, this sets up the excel sheet and enters the times into it, formatting it etc. This is what the sheet would then look like:

    ![alt text]()

- In this program, the start and end times can be altered, and for each time - the cell adjacent to it represents the time spent on a task from that time to that time + 15 minutes, upper bound not inclusive. For example, if cell B3 is 07:00 and cell C3 reads 'Programming' that means that the activity 'Programming' spanned from 07:00 to 07:15 (not including 07:15).
- This program also initialises certain key data structures that are integral to this program and shelves them for later use. I will go into further detail about the choice of my data structures shortly.
- The reason for why this data is shelved and retrieved from the shelve is an interesting one and allows me to move on to the main time_logger_gui program. This program creates a GUI through which the users can easily interact with the program and log their information. A challenge that arose was that the program should be able to store data for the day and wipe it when the day changed and while that would be fairly easy if the program was allowed to run infinitely, in most cases, it is not so and therefore the program had to be able to manage data storage in cases where the user quit and restarted the program many times a day.
- This dilemma was solved by shelving data. The program, on start-up, retrieves all of its data from the shelved files, which holds all the key data structures as well as a string containing the date that the files were shelved, achieved by the datetime module. If the date of the day that the program is running is the same as the one shelved, then it means that the program has restarted on the same day and so it uses all the data stored in the shelve, and if it doesn't match, it means the date has changed and therefore it initialises its structures with some default values. If the user quits the program, either through the quit button or the tk 'X' button (through remapping of the tk protocol), all relevant data in the program is first shelved in the same shelve file and then the program quits so the next time that the program is launched, it has data to work with.
- The data structures that are used throughout this program and serve as its very foundation are:
    - times_list = a list of all the times that are in the sheet ranging from the start time to the end time with 15 minute increments.
    - activities = a list of all the activities that the user could possible log, the activity at the 0 position is None and the user can add upto 12 other activities - comes pre-equipped with generic activities. The reason for the none activity will be covered shortly.
    - act_data = an array the size of the times list where the value at the ith index position reveals the activity that was being performed at the time at the ith position of the times_list. Here values range from 0 to the length of the activities list where the value corresponds to the index position of the activity in the activities array. This index is initialised with the value 0, indicating that no activity has been logged for any time. The choice for this representation of activities is simply being more memory efficient as storing a list of repeated strings does not seem very logical, therefore this list serves as a quasi-relational database. The decision to include the value None at index 0 and not simply storing None in act_data was more of a stylistic choice as it meant that I did not have to validate for the None type and could perform integer operations on the values of this list carefree.
    - added_sheet = a list of the size of times_list with boolean values. A value of True at an index position means that  an activity has been logged for the time corresponding to the same index position in the times_list array. Used so that if a cell has already had some value logged to it, the process is not redundantly repeated.
    - merged = a list of strings containing the cells that have been merged and need to be unmerged in case of a conflicting input.
    - colours = a list of colours to use for each activity.
- Here is the an image of the GUI main page:

    ![alt text]()

- This page allows users to enter data that is logged by choosing their activity, entering a start and end time - both in military format. In case of a conflict with data, there is a popup that emerges and informs the user. If the users enter a time for an activity that spans more than one cell, the relevant cells are merged. All activities are coloured with a distinct colour in the colour_list.
- In a conflict, all cells are unmerged, the cells are reformatted and the added_sheet is reinitialised with all False values so the act_data can be re-entered with the new correct values.
- Here is an image of the log sheet after some activities have been inputted:

    ![alt text]()

- The second page of the GUI allows users to edit the activities, giving them the option to add or delete any, and in case of deletion, all logs for that activity are removed and the activity is popped from the activities list. Therefore all index positions higher than it must be reduced in the act_data list as the size of the activities list has reduced. The value in the activities list must be popped and not simply made None to avoid the activities list outgrowing the colours list while actual activities may remain less than 12.
- Here is an image of the second page/tab of the GUI:

    ![alt text]()

- The third page of the GUI uses matplotlib to construct a pie chart for all the activities have been logged as well as displaying some other information, such as number of hours logged, activity that has been logged maximum etc. It also allows users to save the analytics report through the module pillow which grabs the entire tkinter window and saves it as a png file, with a result as displayed below:

    ![alt text]()

# P.S:

- The program initialise_logger is to be run once and before any other program! 
- The file_paths program is imported by the other programs and is used to make changing of the paths for storage a lot easier, and therefore if anyone wants to use this build, they simply have to create a file path program and add in their own paths for the shelved_data, log_sheet (the excel sheet) and the image_file (the analytics report saved by pillow).
- While it is possible to view the excel file as the image runs, it is currently not possible to edit it while you view it owing to permission errors and the openpyxl module - therefore the log any information or delete an activity, please make sure the excel file is not open!
