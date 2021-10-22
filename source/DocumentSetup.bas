Attribute VB_Name = "DocumentSetup"
' Microsoft Word Setup and Update macros
'
' Copyright (c) 2020 Charles Weir
'
' This contains two VBA macros that are useful with almost any document.
' SetupMasterForEditing: Changes the view and layout (especially for master documents)
'     to something more useful than the default.
' UpdateAllFields: Updates every field, contents table, and reference in the document.
'
'This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

' Uncomment this to have the setup happen automatically on open.
'Sub Document_Open()
'    SetupMasterForEditing
'End Sub

Sub UpdateButtonPressed(control As IRibbonControl)
    UpdateAllFields
End Sub

'' Update all the fields, indexes, etc. in the active document.
Sub UpdateAllFields()
    
' Updates all fields, and tables of contents.
' Assign to button or keyboard shortcut.

    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Do this twice. Figure numbers seem to update the first time, references to them the second time
    Dim i As Long
    For i = 1 To 2
        Application.StatusBar = "Updating fields. Pass " & i & " of 2"
        DoEvents
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
        
        '' For footnotes, endnotes and comments, we get a pop-up
        '' "Word cannot undo this action. Do you want to continue?". Prevent it.
        Application.DisplayAlerts = wdAlertsNone
        Dim sr As Range
        For Each sr In doc.StoryRanges
            sr.Fields.Update
            While Not (sr.NextStoryRange Is Nothing)
                Set sr = sr.NextStoryRange
                sr.Fields.Update
            Wend
        Next sr
        Application.DisplayAlerts = wdAlertsAll
        
    Next i
    Application.StatusBar = " "
    DoEvents
        
    ' We can also do further processing, e.g. to update references.
    ' This does nothing if the function DoAdditionalDocumentUpdates doesn't exist.
    
    On Error Resume Next
    Application.Run "DoAdditionalDocumentUpdates"
    On Error GoTo 0
    
End Sub


Sub SetupMasterForEditing()
'
' A master document comes up in a default mode with subdocuments locked and showing comments. Change it back.
' And for every document, move not to show comments, and to show headings in the left hand window.
'
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
End Sub

Sub ChangeView(View As Integer)
' Change to view View - see https://msdn.microsoft.com/en-us/library/office/ff836365.aspx for values
    ActiveWindow.ActivePane.View.Type = View
    DoEvents
End Sub

Sub ZoomTo(ZoomSetting As Integer)
'
' Set document to preferred zoom level.
'
    ActiveWindow.ActivePane.View.Zoom.Percentage = ZoomSetting
End Sub

