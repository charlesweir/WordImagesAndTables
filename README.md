# WordSupport
Microsoft Word VBA Macros and autotext to add missing features.

Download the latest release [here](https://github.com/charlesweir/WordSupport/releases/download/V1.1/ReleaseV1.2.zip)

## WordPicture.docm
This contains two autotext entries, to support proper floating images and tables. It allows you to arrange them as Latex or other page layout tools would do it. There is VBA macro support, both to manage the layout, and to make it easy to replace images with new versions. 

Full instructions how to set up and use the items are in the document.

## DocumentSetup.bas
This contains two VBA macros that are useful with almost any document. 
To install them go to the Visual Basic Editor (Alt-F11), then Project Window - Normal - Modules - Right click Import File... 
I recommend assigning the two macros to buttons (as described in WordPicture.docm).

**SetupMasterForEditing**: Documents, especially master documents, tend to open in particular modes and zoom settings that aren't necessarily what you want. This changes them to more suitable modes (including opening all sub-documents), and can be modified to your requirements.

**UpdateAllFields**: Word's field updating is a bit haphazard; some of it happens only when you print; others at other times; some only happen manually. This macro updates every field and all tables of contents in the document. So you can find 'Error!' messages and fix them easily. For Mendeley users (and Zotero) it also updates all the inline references and the bibliography. It also deals with the problem that neither allows you to reformat the bibliography, by assigning a style to it.
