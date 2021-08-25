# Microsoft Word Plug-in for Image and Table Layout

This plug-in makes Word into a desktop publisher, making it easier to put images and tables in frames and lay them out automatically. 

On its own, Microsoft Word is poor at positioning images and tables. You have to position images manually; captions don't work very well; and when you add text or change anything everything ends up in a mess. But it doesn't have to be like that. Other word processing packages are rather good at positioning images and tables: Latex users point with pride at its clever image positioning; FrameMaker flows everything around pictures; InDesign and QuarkExpress do wonders! Yet Microsoft Word is programmable, so why not make Word do the same?

This add-in does just that. It makes it easy to insert pictures and tables within frames, and lays out all those frames in your document in a pleasing way so that each frame is as close as possible to the main reference to it, without being constrained to be on the same page.

The plug-in also addresses two frustrations using Word layouts:

*  If you have an graphic-creating tool that doesn't support Word embedding, like many nowadays, then you have to save the graphic as a PNG, JPEG, or even PDF on a Mac, then import it into Word as a picture; then you have to crop and size it to suit your needs. And when you need to make changes to the graphic, Word's 'change picture' function looses all the size and cropping you've so carefully set up and you have to crop and size it again--every time! The **Replace Picture** function solves that problem by remembering the size and crop for the picture. It makes changing external graphics as easy as object embedding!

* Word doesn't update its fields consistently, especially cross references. Only when you print does it update the fields, so only then do you see the errors from lost field references (and even then with some fields you may need to print twice!). The **Update All** function solves that problem by fully updating every field.

The plug-in works on the latest (2020) Microsoft Office installations: Word for Windows version 16, and Word for Mac version 16. It does not support Word 365 online, since that doesn't support VBA.

## The Functions (and When to Use Them)

The plug-in creates six new buttons in the *Layout* tab:

**New Figure**: creates a new figure in a frame with a figure caption, plus a reference to it at the insertion point. It uses a placeholder figure; use the **Replace Picture** button to chose another. It puts the frame at the top of the page; use the **Reposition** button to move the frame around.

**New Table**: creates a new table in a frame with a caption, plus a reference to it at the insertion point. Replace the table with what you want. The frame starts at the bottom of the page; use the **Reposition** button to move it around.

**Replace Picture**: does the same as Word's **Change Picture** button, but keeps the same size and cropping. Select an image before clicking it, of course! 

**Reposition**: moves a frame around the page consistently with the Latex formatting rules. So big frames go at the top or bottom of the page; small frames in a two-column page go at the top or bottom of a column; small frames in a single-column page go to the left or right of the text. Clicking the button twice moves the frame to the other position--top to bottom, left to right--and back. The operation doesn't move the frame's anchor, so the frame always stays on the same page. Just select a frame and try it!

**Update All**: Updates all the fields in the document, reliably. To check for cross referencing errors, search for "Error!" and " 0" afterwards.

**Relayout Document** is the big daddy of the functions here. It looks through the current document for all the frames with references to them (as created by **New Figure** and **New Table**), checks that you want to go ahead, then arranges the frames at the top and bottom of columns and pages, each as close as possible to its reference, according to the Latex formatting rules. To make this possible, it moves Word's *Anchor Points* for each frame (unlike **Reposition**). **Relayout Document** ignores images and tables that are not in frames, frames without references to them, frames that are positioned *Left* or *Right*, and references that are in different sections from their corresponding frames. It takes a while, but the results are usually excellent.

All the functions support **Undo**, so experiment as much as you like!

## How to Install the Plug-in

Download the latest version of ImagesAndTableSupport.dotm by clicking the **Download** button in the page [here](https://github.com/charlesweir/WordImagesAndTables/blob/master/ImageAndTableSupport.dotm). Copy that file as follows:

### On Windows

Copy *ImageAndTableSupport.dotm* to *%AppData%\Microsoft\Word\Startup*  
To go to that directory in Windows File Explorer, type the above string into the address bar and hit enter.  

### On Mac

Copy *ImageAndTableSupport.dotm* to *~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word*  
To get to that folder in Finder, use Cmd-Shift-G, paste the above string into the dialog and click OK.

## Upgrading and Uninstalling.

To upgrade, simply download the latest as above, and overwrite the previous version. To uninstall, delete the file ImageAndTableSupport.dotm in the directory given above.

## Developer instructions

The test suite is in WordSupportTest.docm. The introduction part of the document also contains basic instructions how to edit, test and and debug the package.

