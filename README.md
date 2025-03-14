# This version of this tool is outdated. [The new version can be found here](https://github.com/FTimm90/custclrtool)
---

<img src="./doc_img/tool_GUI.png">

# CustClrTool for PowerPoint
A handy little tool for adding up to 50 additional colors (custom colors) to a PowerPoint presentation or template

---

## Background

In PowerPoint, you can generally define 10 colors for a template. The first four are for text and background colors, and the following six are used, for example, when a new graph is created. However, sometimes that number of colors is not enough. Right below the theme colors, there are some additional colors that are automatically generated from the theme colors. Unfortunately, you have no control over these, and they are often unsuitable, especially in a professional context where a company CI has clear rules for what colors to use and what percentage levels of those colors can be used. Fortunately, there is a way to add additional colors.<br>
This is where the CustClr tool comes into play. It offers a simplistic GUI where you can choose a file (either .pptx or .potx) and add up to 50 additional colors that will be located below the theme colors.<br>
Initially I planned on adding more functionality and that may still happen in the future, that's the reason for the Icon and "Custom Colors" on the left of the GUI.

<img src="./doc_img/colors_example.png">

## Features

•	Detecting already existing custom colors and loading them into the GUI<br>
•	Changing every color as desired or removing them completely<br>
•	Adding or removing names of custom colors<br>

## How to Use

Simply run the tool and choose a file. The current file name will be visible right next to the “choose file” button. After that, you must choose a theme file you want to change. If you don’t choose a theme, it won’t work. The name of the theme will also show up right next to the dropdown menu. Usually, there is just one theme file within the PowerPoint file, but depending on how old the presentation or template is, there is always a chance for duplicate themes due to slides that have been copied over from other presentations. The theme name being visible should help you identify if you picked the correct one.<br>
Once you have chosen a theme and potentially existing colors have been found, you can start editing them. Above the color preview tile, you can enter a name. If you do not wish to enter a name, you have two options. You can either leave the name entry field completely empty, which will result in PowerPoint showing the RGB values while hovering the cursor over the color, or you can simply enter a space, which will result in an empty field while hovering over.<br>
The switch between the color name and the preview tile is to activate or deactivate the color. If it’s deactivated, the color will not be added to the file. If there are any custom colors in the file and you deactivate them all, they will all be removed.
Right below the color preview tile, you can enter the hex value for the color you want to change or add. Hex values are what PowerPoint internally works with while dealing with colors. If you enter the desired value, you need to confirm by hitting the enter key.<br>
Once you’re done configuring all your colors, simply press the “confirm” button and that’s it.

<img src="./doc_img/tool_GUI_filled.png">

## Known Issues/Current Limitations

If you are using a version of PowerPoint that is older than the 2016 version, there is a high chance that this tool will not work, since the internal structure has slightly changed.<br>
I also need to note that since this is my very first programming project, it is far from perfect. It does what it’s supposed to do, but there are some things that at this point I simply don’t know how to fix or work around. I STRONGLY advise you to make a copy of your PowerPoint file first and work with that to have a backup.<br>
For this tool to work, the PowerPoint file needs to have its file extension changed to .zip. The tool does that automatically to detect theme files and existing custom colors, and in the end, if you confirm your changes, it gets changed back to its original extension. If you load a file into it and close the tool without confirming any changes, your PowerPoint file will still have the .zip extension. You can simply change it back to what it was before, and it should work without any issues.
You can only make changes to one theme at a time! If you want to change colors for multiple themes within the same PowerPoint file, you currently must confirm your changes, load the same file again, and then choose a different theme file. I would also advise you to restart the tool.

## How it Works

The entry point for this tool is the fact that you can simply change the file extension of a PowerPoint file to .zip and access all parts that it’s made of, which are mostly .xml files and the media files (images, videos, etc.) that it contains.
Within the structure of the .zip, there are theme.xml files. There is always at least one (theme1.xml), but the number is not limited, and it refers to the PowerPoint master(s) that are defined within the file.<br>
This tool changes the file extension from .pptx or .potx to .zip, scans for theme.xml files, and then opens the desired one, loading the name and all found colors directly. The pptxClass stores things like the file name, file extension, and the found themes, basically everything just relating to the presentation. The CustClrClass stores what’s found within the theme file, such as the hex values and names of the custom colors, and also the switches and everything that is needed to construct the tiles that make up the 50 additional colors.<br>
Once everything is set up and the “confirm” button is pressed, the tool searches through the theme file again, deletes any existing custom colors to avoid any confusion or doubling of any kind. Then the new custom colors are constructed with the necessary .xml tags and injected into the correct position within the .xml file.<br>
Since it’s unfortunately not that easy to just save a new and changed .zip file, it then gets duplicated with all the same content as the original, but with the updated theme file. And as a last step, the initially stored filename and file extension are used to change the file back from .zip format to its initial extension.

<img src="./doc_img/structure.png">
