# Progress bar (continuous or discrete) on the PowerPoint slides

![progressbar](images/slides_bar.png)

Using Visual Basic in Microsoft PowerPoint, it produces a discrete or continuous progress bar at the targeted location (either at the top or bottom of each slide). 

## How to use

1) Open the macro_test.ppm file and select "Enable Macros".
2) Open your PowerPoint file.
3) In your PowerPoint file, click "Macros" under the "View" tab.
4) Chose "Macros in:" to "macro_test"
5) Select "ProgressBar_discrete" (for discrete progress bar) or "ProgressBar" (for continuous progress bar) and click the "Edit" button
6) In the code, you can set the configurations:
 - no_bar_slide = slide numbers that do not contain the bar (e.g., title slide, hidden slide ...)
 - no_progres_slide: slide numbers that have the same size of the bar as the previous one
 - bar_loc_x0: left most position of the progress bar (only for the continuous one)
 - bar_loc_top: (True or False): True -> Top / False -> Bottom
 - bar_loc_y_margin: y-margin of the bar from the top or bottom of each slide
 - bar_height: height of the progress bar
 - bar_color: the color of the progress bar
 - bar_trp: transparency of the bar for the remaining slides (only discrete one)
 - bar_gap: gap width between bars, the length is relative to the bar size (e.g., bar_gap=1: the gap width is the same as the bar length); only for discrete one
7) After editing the code, run the macro with "Macros" > "Run" 
8) If you need to erase the progress bar for all slides, run the macro, "RemoveMacros" (for a continuous bar) or "RemoveMacros_multiple" (for a discrete bar), and do steps 4-7 if you want to re-draw the progress bar. 

