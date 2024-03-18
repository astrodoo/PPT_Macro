# Progress bar (continuous or discrete) on the PowerPoint slides

![progressbar](images/slides_bar.png)

Using Visual Basic in the Microsoft PowerPoint, it produces discrete or continuous progress bar at the targeted location (either of top or bottom of each slide). 

## How to use

1) Open macro_test.ppm file and select "Enable Mcros".
2) Open your powerpoint file.
3) In your powerpoint file, click "Macros" under "View" tab.
4) Chose "Macros in:" to "macro_test"
5) Select "ProgressBar_discrete" (for discrete progress bar) or "ProgressBar" (for continuous progress bar) and click "Edit" button
6) In the code, you can set the configurations:
 - no_bar_slide = slide numbers that do not contain the bar (e.g., title slide, hidden slide ...)
 - no_progres_slide: slide numbers that have the same size of bar as the previous one
 - bar_loc_x0: left most position of the progress bar (only for the continous one)
 - bar_loc_top: (True or False): True -> Top / False -> Bottom
 - bar_loc_y_margin: y-margin of the bar from top or bottom of each slide
 - bar_height: height of the progress bar
 - bar_color: color of the progress bar
 - bar_trp: transparency of the bar for the remaining slides (only discrete one)
 - bar_gap: gap width between bars, the length is relative to the bar size (e.g., bar_gap=1: the gap width is the same as the bar length); only for discrete one
7) After editing the code, run the macro with "Macros" > "Run" 
8) If you need to erase the progress bar for all slides, run the macro, "RemoveMacros" (for continuous bar) or "RemoveMacros_multiple" (for discrete bar) and do the step 4-7 if you want to re-draw the progress bar. 

