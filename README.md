# Warehouse control application

Created by [Rosvaldas Šlekys](https://github.com/RosSlek) 

This project was made to deepen my knowledge about python, learn new things, better understand what I already knew and apply it to real life situations.

## The main goals for this project were to:
#### •	Create app, which could be used for small business and make day-to-day job easier
#### • Ensure traceability in warehouse and production
#### • Track material quantities
#### • Learn about Tkinter
#### • Have fun
#### • Improve

## Main steps:
#### •	Learn about Tkinter, create nice looking UI.
#### • Add functionality and comfortability.
#### •	Adapt app to real business example

## Result:
### Created UI to represent daily workers activities.
![image](https://github.com/RosSlek/Sandelio-valdymo-programa/assets/149397027/3d2e3d2b-92b6-4a74-bab2-e77ea8d0b3cb)

### Created functions to add materials to the warehouse, transfer materials to production with optional comment fields and timestamps when action occured. Functions to show remaining inventory in warehouse and production, informational popup windows.

![image](https://github.com/RosSlek/Sandelio-valdymo-programa/assets/149397027/00d29c9b-e37f-479a-b06a-3a8f7c507266)

![image](https://github.com/RosSlek/Sandelio-valdymo-programa/assets/149397027/0bc7e0ce-3b24-4ce6-bb35-a738dbbe82ee)

### Created function to registrate work orders, to write off used materials, added safety features to fill correct data.

![image](https://github.com/RosSlek/Sandelio-valdymo-programa/assets/149397027/df0b63bd-4ead-45b4-a317-ceea0c779c9a)

### Added switches to change view of inventory in warehouse/production, light/dark mode.

![image](https://github.com/RosSlek/Sandelio-valdymo-programa/assets/149397027/8a33fb17-da54-4e43-bb53-49eb0e946430)

![image](https://github.com/RosSlek/Sandelio-valdymo-programa/assets/149397027/5a623b46-0762-401f-aaf4-400f9803a8d7)

### Program creates excel files (not CSV, so it would be more friendly to lithuanian language and user) with history of every action, in case of mistakes it can be easily corrected manually. Positive numbers show added materials, negative removed or consumed.
![image](https://github.com/RosSlek/Sandelio-valdymo-programa/assets/149397027/8fd71ede-b352-4a9d-9e48-0c0636605a79)

## Conslusion
It was a really nice experience to work on this project, a bunch of new things learned, tons of bugs and puzzles to solve. I`m quite happy with the result and will continue to develop my skills. You can download this apllication and try it yourself [here](https://www.dropbox.com/scl/fi/vfekzijr6ds3hh8i2ra6b/Sand-lio-programa.rar?rlkey=svv8ing3xq0fnenmi4ja4289h&dl=0).
*Excel and exe files have to be in the directory they are, create shortcuts if needed. There also master files created deeper in the '_internal' folder with warehouse and production history so they could be restored if main excel files become corrupted.
