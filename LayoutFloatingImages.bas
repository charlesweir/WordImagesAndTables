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
    
    ShowStatusBarMessage ("Analysing frames to reposition")
    Dim clsAnchoredFrame As New AnchoredFrame
    Dim oAnchoredFrame As AnchoredFrame
    Dim myFramesToLayout As Collection
    
    Set myFramesToLayout = clsAnchoredFrame.RepositionableFramesInRegion(region)

    ' OK. So now we have all the frames we want to reposition, and the locations of their first relevant reference:
    ' Move the frames to the end of the document.
    
    Application.ScreenUpdating = False
    For Each oAnchoredFrame In myFramesToLayout
        ShowStatusBarMessage ("Stashing " & oAnchoredFrame.Name & " of " & myFramesToLayout.count)
        oAnchoredFrame.Stash
    Next oAnchoredFrame
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    
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
    aDataObject.SetText Text:=" " ' For Mac - don't ask!
    aDataObject.PutInClipboard
End Sub


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


