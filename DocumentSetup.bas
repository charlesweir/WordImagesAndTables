Attribute VB_Name = "DocumentSetup"
Sub SetupMasterForEditing()
'
' The document comes up in a default mode with subdocuments locked and showing comments. Change it back.
' Note. Sometimes with large documents the Word libraries take too long to load, and this macro
'    fails with a popup box: Abort or Debug.
'    Just chose Abort, and invoke it again.
'
    ChangeView (wdOutlineView)
    ActiveDocument.Subdocuments.Expanded = True
    ChangeView (wdPrintView) ' Change to your preference.
    With ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = wdRevisionsViewFinal
    End With
    ActiveWindow.DocumentMap = True  ' Restore the Navigation Pane
    ZoomTo (200) ' Zoom Percentage. Change to your preference.
    ' FixStyles
End Sub

Sub ChangeView(View As Integer)
' Change to view View - see https://msdn.microsoft.com/en-us/library/office/ff836365.aspx for values
    ActiveWindow.ActivePane.View.Type = View
    DoEvents
End Sub

Sub ZoomTo(ZoomSetting As Integer)
'
' Set document to preferred zoom level. Change level to your own preference.
'
    ActiveWindow.ActivePane.View.Zoom.Percentage = ZoomSetting
End Sub

Sub UpdateAllFieldsIn(doc As Document)
' Updates all fields, and tables of contents.
' Assign to button or keyboard shortcut.
' Also updates Mendeley references and sets the table of references style (for IEEE) to MendeleyReference
' Todo: update header and footer?

    Application.StatusBar = "Updating fields..." 'N.b. Doesn't seem to work despite what the documentation says.
    
    ' Do this twice. Figure numbers seem to update the first time, references to them the second time
    Dim i As Long
    For i = 1 To 2
        '' Update tables. We do this first so that they contain all necessary
        '' entries and so extend to their final number of pages.
        Dim toc As TableOfContents
        For Each toc In doc.TablesOfContents
            toc.Update
        Next toc
        Dim tof As TableOfFigures
        For Each tof In doc.TablesOfFigures
            tof.Update
        Next tof
        '' Update fields everywhere. This includes updates of page numbers in
        '' tables (but would not add or remove entries). This also takes care of
        '' all index updates.
        Dim sr As Range
        For Each sr In doc.StoryRanges
            sr.Fields.Update
            While Not (sr.NextStoryRange Is Nothing)
                Set sr = sr.NextStoryRange
                '' FIXME: for footnotes, endnotes and comments, I get a pop-up
                '' "Word cannot undo this action. Do you want to continue?"
                sr.Fields.Update
            Wend
        Next sr
        
    Next i
    
    ' Now update references
     ' If you use Zotero or similar, replace this with their update method.
     
    ' For Mendeley. Needs Tools - References - MendeleyPlugin ticked.
    Refresh
    
    ' And change the bibliography to be the 'Bibliography' style.
    RestyleBibliography
End Sub
'' Update all the fields, indexes, etc. in the active document.
'' This is a parameterless subroutine so that it can be used interactively.
Sub UpdateAllFields()
    UpdateAllFieldsIn ActiveDocument
End Sub

Sub RestyleBibliography()
'
' Changes the style of text within the 'Bibliography' bookmark to 'Bibliography'. Useful for Mendeley and Zotero,
' that use hard-coded formatting.
'
' To set up, select the whole bibliography field, and
' (Mac) Menu - Insert - Bookmark... - Bibliography; (PC) Insert - Bookmark - Bibliography
'
    Dim currentPosition As Range
    
    ' If there's no Bibliography style or bookmark, skip this (though actually I think Bibliography is a built-in style)
    On Error GoTo ExitSub
    BibliographyStyle = ActiveDocument.Styles("Bibliography")
    Set currentPosition = Selection.Range ' save current cursor position
    With ActiveDocument.Bookmarks("Bibliography").Range
        .Select
        Selection.ClearParagraphDirectFormatting
        .style = BibliographyStyle
    End With
    currentPosition.Select ' return cursor to original position
    
ExitSub:
    On Error GoTo -1 ' VBA Magic to reset error handling
End Sub

