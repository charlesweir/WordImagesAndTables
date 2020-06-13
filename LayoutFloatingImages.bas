Attribute VB_Name = "LayoutFloatingImages"
' Function to relayout the image frames in a document according to the Latex rules.
'
' Copyright (c) 2020 Charles Weir
'
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Option Explicit ' All variables must be defined.


Sub LayoutFloatingImages()
    LayoutFloatingImagesFor ActiveDocument.StoryRanges(wdMainTextStory)
End Sub

Public Sub LayoutFloatingImagesFor(region As Range)
    ' Moves every text frame containing a captioned Image or Table to be anchored at its first reference.
    '
    ' Word for Mac 16 sometimes moves a wrong frame (and sometimes crashes).
    ' There seem to be some odd bugs in its copy and paste too.

    ShowStatusBarMessage ("Analysing frames to reposition")
    Dim AnchoredFrameClass As New anchoredFrame
    Dim oAnchoredFrame As anchoredFrame
    Dim myAnchoredFrames As Collection
    Set myAnchoredFrames = New Collection
    Dim myFramesToLayout As Collection
    Set myFramesToLayout = New Collection
    
    ' First, find all the frames and their hidden cross reference bookmarks:
    Dim currentFrame As Shape
    For Each currentFrame In region.ShapeRange
        If AnchoredFrameClass.IsValidFrame(currentFrame) Then
            Set oAnchoredFrame = New anchoredFrame
            Set oAnchoredFrame.Frame = currentFrame
            myAnchoredFrames.Add item:=oAnchoredFrame, key:=oAnchoredFrame.BookmarkId
        End If
    Next
    
    ' Now look for the first reference to each of those bookmarks, and construct a list of them:
        

    Dim ReferencingField As field
    Dim bookmarkName As String
    Dim previousField As field
    
    For Each ReferencingField In region.Fields
        ' Word can take several minutes to sort the fields after you open a new document.
        ' If this next assertion fails, get on with something else and come back to do this again later.
        If Not previousField Is Nothing Then Debug.Assert ReferencingField.Result.Start >= previousField.Result.Start
        
        bookmarkName = AnchoredFrameClass.BookmarkIdFromField(ReferencingField)
        If bookmarkName <> "" And ContainsKey(myAnchoredFrames, bookmarkName) Then
            Set oAnchoredFrame = myAnchoredFrames(bookmarkName)
            
            ' Only pair references within a section. (I have forward references to figures at the start of my Thesis)
            If (ReferencingField.Code.Information(wdActiveEndSectionNumber) = oAnchoredFrame.SectionNumber) Then
                Set oAnchoredFrame.RefField = ReferencingField
                myFramesToLayout.Add oAnchoredFrame
                ' And remove it, so that we ignore later references.
                myAnchoredFrames.Remove bookmarkName
            End If
        End If
        Set previousField = ReferencingField
    Next ReferencingField
    
    ' OK. So now we have all the frames we want to reposition, and the locations of their first relevant reference:
    ' Move the frames to the end of the document.
    
    For Each oAnchoredFrame In myFramesToLayout
        ShowStatusBarMessage ("Stashing " & oAnchoredFrame.Name & " of " & myFramesToLayout.count)
        oAnchoredFrame.Stash
    Next oAnchoredFrame
    
    ' OK So all the frames are out of the way. Here's the interesting stuff.
    ' Take each one and position it as close as possible to its reference according to the Latex (Knuth?) algorithm.
    
    Dim clsColumnLayout As New ColumnLayout
    clsColumnLayout.Initialise

    For Each oAnchoredFrame In myFramesToLayout
        ShowStatusBarMessage ("Positioning " & oAnchoredFrame.Name)
        clsColumnLayout.PositionFrame oAnchoredFrame
    Next oAnchoredFrame
    
    EmptyCutBuffer
    ShowStatusBarMessage ("Repositioned " & myFramesToLayout.count & " frames")
End Sub

Private Sub EmptyCutBuffer()
    ' Empty cut buffer to stop extra "Do you want to save clipboard?" on exit
    Dim aDataObject As New DataObject
    aDataObject.SetText Text:=Empty
    aDataObject.PutInClipboard
End Sub

    
Private Function ContainsKey(col As Collection, key As String) As Boolean
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    ContainsKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0 ' Reset
End Function

Private Sub ShowStatusBarMessage(message As String)
    If message = "" Then
        Application.StatusBar = " "
    Else
        Application.StatusBar = message
    End If
    DoEvents
End Sub
Private Sub Say(message As String)
    If MsgBox(message, vbYesNo) = vbNo Then End
End Sub


