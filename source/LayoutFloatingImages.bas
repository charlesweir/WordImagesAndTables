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


Public Sub AnalyseImagesToLayoutInDocument(imagesToLayout As collection, imagesIgnored As collection)
' Fills collections of (1) AnchoredFrame objects representing all those needing layout in the current document, and (2) those that a user might expect to be laid out but will not be.
    Dim clsAnchoredFrame As New AnchoredFrame
    ShowStatusBarMessage ("Analysing frames to reposition")
    clsAnchoredFrame.AnalyseFramesInRegion ActiveDocument.StoryRanges(wdMainTextStory), imagesToLayout, imagesIgnored
End Sub

Public Sub LayoutFloatingImagesFor(region As Range)
    ' Re-lays out all the AnchoredFrames in the given region (only used for testing)
    
    Dim imagesToLayout As collection
    Set imagesToLayout = New collection
    Dim clsAnchoredFrame As New AnchoredFrame
    clsAnchoredFrame.AnalyseFramesInRegion region, imagesToLayout, New collection
    LayoutTheseFloatingImages imagesToLayout
End Sub

Public Sub LayoutTheseFloatingImages(myFramesToLayout As collection)
    ' Repositions each of the given AnchoredFrames according, roughly, to the Latex rules.
    
    Dim oAnchoredFrame As AnchoredFrame

    ' Move the frames to the end of the document.
    
    Application.ScreenUpdating = False
    For Each oAnchoredFrame In myFramesToLayout
        ShowStatusBarMessage ("Stashing " & oAnchoredFrame.Name & " of " & myFramesToLayout.Count)
        oAnchoredFrame.Stash
    Next oAnchoredFrame
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    
    ' OK So all the frames are out of the way. Here's the interesting stuff.
    ' Take each one and position it as close as possible to its reference according to the LaTeX (Knuth?) algorithm.
    
    Dim clsColumnLayout As New ColumnLayout
    clsColumnLayout.Initialise

    For Each oAnchoredFrame In myFramesToLayout
        ShowStatusBarMessage ("Positioning " & oAnchoredFrame.Name)
        clsColumnLayout.PositionFrame oAnchoredFrame
    Next oAnchoredFrame

    ShowStatusBarMessage ("Repositioned " & myFramesToLayout.Count & " frames")
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


