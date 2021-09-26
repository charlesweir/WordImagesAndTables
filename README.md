# Microsoft Word Plug-in for Image and Table Layout

This plug-in improves Microsoft Word, making it easier to work with cross-references, images and tables. 

On its own, Microsoft Word is poor at positioning images and tables. You have to position images manually; captions don't work very well; and when you add text or change anything, everything ends up in a mess. But it doesn't have to be like that. Other word processing packages are rather good at positioning images and tables: LaTeX users point with pride at its clever image positioning; FrameMaker flows everything around pictures; InDesign and QuarkExpress do wonders! Yet Microsoft Word is programmable, so why not make Word do the same?

This add-in does just that. It makes it easy to insert pictures and tables within frames, and can lay out all those frames in your document in a pleasing way so that each frame is as close as possible to the main reference to it, without being constrained to be on the same page.

The plug-in also addresses two frustrations using Word layouts:

*  **Preserving image size and cropping**: If you have an graphic-creating tool that doesn't support Word embedding (and there are lots nowadays), then you have to save the graphic as a PNG, JPEG, SVG, or (on a Mac) PDF, then import it into Word as a picture, then crop and size it to suit your needs. That's fine. But when you make changes to the graphic, Word's 'Change Picture' function forgets the size and cropping you've so carefully set up, so you have to crop and size it again: every time! The **Replace Picture** function solves that problem by remembering the size and crop for the picture (and it's border settings too). It makes changing external graphics as easy as object embedding!

* **Updating**: Word doesn't update its fields consistently, especially cross references. Only when you print does it update the fields, so only then do you see the errors from lost field references (and even then with some fields you may need to print twice!). The **Update All** function solves that problem by fully updating every field.

The plug-in works on the latest (2020) Microsoft Office installations: Word for Windows version 16, and Word for Mac version 16. It does not support Word 365 online, since that doesn't support VBA.

## The Functions (and When to Use Them)

The plug-in creates six new buttons in the *Layout* tab:

**New Figure**: creates a new figure in a frame with a figure caption, plus a reference to it at the insertion point. It uses a placeholder figure; use the **Replace Picture** button to chose another. It puts the frame at the top of the page; use the **Reposition** button to move the frame around, or **Relayout Document** to rearrange all the frames.

**New Table**: creates a new table in a frame with a caption, plus a reference to it at the insertion point. Replace the table with what you want. The frame starts at the bottom of the page; again, use  **Reposition** or **Relayout Document** to move it around.

**Replace Picture**: does the same as Word's **Change Picture** button, but keeps the same size and cropping. Select an image before use. 

**Reposition**: moves a frame around the page consistently with the LaTeX formatting rules. So big frames go at the top or bottom of the page; small frames in a two-column page go at the top or bottom of a column; small frames in a single-column page go to the left or right of the text. Clicking the button twice moves the frame to another position: top vs bottom, or left vs right. The operation doesn't move the frame's anchor, so the frame always stays on the same page. Just select a frame and try it!

**Update All**: Updates all the fields in the document, reliably. To check for cross referencing errors, search for "Error!" and " 0" afterwards.

**Relayout Document** is the most complex of the functions here. It looks through the current document for all the frames with references to them (as created by **New Figure** and **New Table**), checks that you want to go ahead, then arranges the frames at the top and bottom of columns and pages, each as close as possible to its reference, according to the LaTeX formatting rules. To make this possible, it moves Word's *Anchor Points* for each frame (unlike **Reposition**). **Relayout Document** ignores images and tables that are not in frames, frames without references to them, frames that are positioned *Left* or *Right* (in a single column page), and references that are in different sections from their corresponding frames. It takes a second or two per frame laid out, but the results can be excellent.

## Using the Functions

All the functions support **Undo**, so experiment as much as you like!

We recommend doing **Update All** after **Relayout Document**, as the figure, table and page numbering may change.

## How to Install the Plug-in

Download the latest version of ImagesAndTableSupport.dotm by clicking [**Download** here](https://github.com/charlesweir/WordImagesAndTables/releases/latest/download/ImageAndTableSupport.dotm). Copy that file as follows:

### On Windows

Copy *ImageAndTableSupport.dotm* to *%AppData%\Microsoft\Word\Startup*  
To go to that directory in Windows File Explorer, type the above string into the address bar and hit enter.  

### On Mac

Copy *ImageAndTableSupport.dotm* to *~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word*  
To get to that folder in Finder, use Cmd-Shift-G, paste the above string into the dialog and click OK. If it's not there, try *~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word*

## Upgrading and Uninstalling.

To upgrade, simply download the latest as above, and overwrite the previous version. To uninstall, delete the file *ImageAndTableSupport.dotm* in the directory given above.

## Developer instructions

The test suite is in *WordSupportTest.docm*. The introduction part of the document also contains basic instructions how to edit, test and and debug the package.

## Troubleshooting

Sometimes **Relayout Document** may fail to identify some of the images and table frames to layout, and the result is usually messy. A good way to spot the problem is to check the count in the "Found <count> Figures and Tables to layout" dialog and see if that corresponds to the number of figures + tables you want laid out. Omitted frames can be because of a Word documentation corruption somewhere. If a frame is omitted, check:
* That the image or table concerned is in a frame and has a caption within the frame. **Relayout Document** doesn't lay out floating images -- even ones with captions -- unless they're in a frame.
* That the frame you want laid out has a corresponding reference to its caption *in the same section* in the main text (this feature allows you to have forward references to images or tables in earlier sections)
* That the invisible bookmark in the caption has not got lost; do **Update All**, and fix any references that show **Error! Reference source not found.**
* That (in a single-column section of the document) the frame isn't set to align *left* or *right*. This feature to permit small figures with text wrapped around them.


