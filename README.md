# Excellent Microsoft Word Plug-in for Images and Tables

This plug-in makes Word into a desktop publisher, making it easy to put images and tables in frames and laying them out automatically. 

On its own, Microsoft Word is pretty poor at positioning images and tables. Normally you have to position images manually; captions don't work very well; and when you add text or change anything, everything ends up in a mess. But it doesn't have to be like that. Other word processing packages are rather good at positioning images and tables: Latex users point with pride at its clever image positioning; FrameMaker flows everything around pictures. But Microsoft Word is programmable, so why not make Word do the same?

This add-in does just that. It makes it easy to insert pictures and tables within frames, and lays out all those frames in your document in a pleasing way so that each frame is as close as possible to the main reference to it, without being constrained to be on the same page.

The plug-in also solves some other frustrations with Word:

*  If you have an graphic-creating tool that doesn't support Word embedding (like many nowadays), then you have to save the graphic as a PNG/JPEG (for Windows) or PDF (for Mac) then import it into Word as a picture, then crop and size it to suit your needs. And when you need to make changes to the graphic, Word's 'change picture' function looses all the size and cropping you've so carefully set up and you have to crop and size it again--every time! The **Replace Picture** function solves that problem by remembering the size and crop for the picture. It makes the changing external graphics as easy as object embedding!

* Word doesn't update its cross reference fields very consistently. Only when you print does it update the fields and you see the errors from lost field references (and sometimes not even then). The **Update All** function solves that problem by doing a full update when you press the button. So then you can search for "Error!" and " 0" to look for missing references.

* If you use Mendeley or Xotero, that the standard formatting for the table of references may not be exactly what you want; references break across pages, for example. The **Update All** button assigns a standard Word style ("Bibliography") to the references, so you can change the formatting as you like.

The plug-in works on the latest (2020) Microsoft Office installations: that is Word for Windows version 16, and Word for Mac version 16, but not Word 365 online.

## The Functions (and When to Use Them)

The plug-in creates six new buttons in the *Layout* tab:

**New Figure**: creates a new figure in a frame at the top of the page with a figure caption, plus a reference to it at the insertion point. It uses a placeholder figure; use the **Replace Image** button to chose another, and the **Reposition** button to move the frame around.

**New Table**: creates a new table in a frame at the bottom of the page containing a table caption, plus a reference to it at the insertion point. Replace the table with what you want, and use the **Reposition** button to move the frame around.

**Replace Picture**: does the same as Word's **Change Picture** button, but keeps the same size and cropping. Select an image before clicking it, of course!  It saves a lot of effort when you're using external tools to create images.

**Reposition**: moves a frame around the page consistently with the Latex formatting roles. So big frames go at the top or bottom of the page; small frames in a two-column page go at the top or bottom of a column; small frames in a single-column page go to the left or right of the text. Clicking the button twice moves the frame to the other position--top to bottom, left to right--and back. The operation doesn't move the frame's anchor, so the frame always stays on the same page. Just select a frame and try it!

**Update All**: Updates all the fields in the document, reliably; it also assigns the "Bibliography" style to any Mendeley or Xotero references.

**Relayout Document** is the big daddy of the functions here. It looks through the current document for all the frames with references to them (as created by **New Figure** and **New Table**), checks 
 that you want to go ahead, then arranges the frames at the top and bottom of columns and pages, each as close as possible to its reference, according to the Latex formatting rules. To make this possible, it moves Word's *Anchor Points* for each frame (unlike **Reposition**). **Relayout Document** ignores images and tables that are not in frames, frames without references to them, and frames that are positioned *Left* or *Right* (and references that are in different sections from their corresponding frames). It takes a while, but the results are usually excellent.

All the functions support **Undo**, so experiment as much as you like!

## How to Install the Plug-in

Download the latest release [here](https://github.com/charlesweir/WordSupport/releases/download/V2.0/ReleaseV2.0.zip)

Unzip the package, and copy the file *ImageAndTableSupport.dotx* as follows:

### On Windows

Copy *ImageAndTableSupport.dotx* to *\<Your Home Directory\>/AppData/Roaming/Microsoft/Word/STARTUP*

You can find your home directory using Windows - R, CMD , and it is shown in the prompt. *AppData* is normally hidden, so [here are brief instructions](https://support.microsoft.com/en-gb/help/4028316/windows-view-hidden-files-and-folders-in-windows-10) how to show it in File Explorer. 

### On Mac

 Copy *ImageAndTableSupport.dotx* to *~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word* .
*Library* is normally hidden. In Finder, Use Cmd + Shift + . to reveal it.
