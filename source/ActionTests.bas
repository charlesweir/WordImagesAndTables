Attribute VB_Name = "ActionTests"
' Tests for LayoutFloatingImages.

'Option Explicit ' All variables must be defined.

Sub ActionTests()
    ' Reset the undo buffer, and fix the selection:
    ActiveDocument.UndoClear
    Set cursorLocation = Selection.Range
    
    Figure1Top = ActiveDocument.Shapes(1).Top

    For Each mySection In ActiveDocument.Sections
        Dim oRangeTested As Range
        Set oRangeTested = mySection.Range
        If Not oRangeTested.Text Like "*Resulting Name;*" Then GoTo nextMySection
        ' Does the repositioning, then checks everything's right.
        
        ' First, load the test specs from the first paragraph in the section:
        tests = Split(oRangeTested.Paragraphs(1).Range.Text, Chr(11))
    
        Dim myFrames As Collection
        Set myFrames = New Collection
        
        ' Can't just call this on mac if the template file is in the Startup directory.
        Application.Run "LayoutFloatingImages.LayoutFloatingImagesFor", oRangeTested
        
        ' Now, find all the frames and their hidden bookmarks generated by the Cross Reference to Figure or Table
        For Each shp In oRangeTested.ShapeRange
            If shp.Type = msoTextBox Then
                Set bookmarkSet = shp.TextFrame.TextRange.Bookmarks
                bookmarkSet.ShowHidden = True
                Debug.Assert bookmarkSet.Count > 0 ' If not, we've messed up our references somewhere.
                Debug.Assert bookmarkSet(1).Name Like "_Ref##*" ' If not, I don't know what's going on.
                myFrames.Add Item:=shp, key:=bookmarkSet(1).Range.Text ' Gives "Figure NN"
                
            End If
        Next
        
        ' Check the layout was as expected:
        For Each x In tests
            If Not (x Like "Figure *" Or x Like "Table *") Then GoTo NextX ' There's no continue in this version of VBA
     
            Dim expectedName As String
            expectedName = Split(x, ";")(0)             ' E.g. Figure 1
            expectedLocation = Split(x, ";")(1)         ' E.g. top of column
            expectedColumn = Split(x, ";")(2)           ' E.g. 2
            expectedFirstParaOnPage = Split(x, ";")(3)  ' E.g  12
            
            Set shp = myFrames(expectedName) ' May fail if we've lost that frame.
            
            actualColumn = "" & ColumnNumber(shp.Anchor)
            Debug.Assert actualColumn = expectedColumn
            
            actualLocation = IIf(shp.Top = wdShapeTop, "top", IIf(shp.Top = wdShapeBottom, "bottom", "other"))
            Debug.Assert expectedLocation Like actualLocation & "*"
            
            actualHLocation = IIf(shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn, "column", _
                                IIf(shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin, "page", "other"))
            Debug.Assert expectedLocation Like "*" & actualHLocation
            
            Set firstPara = shp.Anchor
            Set firstPara = firstPara.GoTo(What:=wdGoToPage, Count:=shp.Anchor.Information(wdActiveEndPageNumber))
            actualFirstParaOnPage = firstPara.ListFormat.ListString
            Debug.Assert actualFirstParaOnPage = expectedFirstParaOnPage & "."
NextX:
        Next x
nextMySection:
    Next mySection
    
    ' And reset everything.
    ActiveDocument.Undo (1000)
    
    ' Check Undo worked. If Figure 1 is back as before, then presumably everything else is.
    Set F1 = ActiveDocument.Shapes(1)
    Debug.Assert F1.Top = Figure1Top
    
    ' And reset the cursor.
    cursorLocation.Select
    
    ' Check the list of frames to be laid out...
    Set myFrames = New Collection
    Dim ignoredFrames As New Collection
    Application.Run "LayoutFloatingImages.AnalyseImagesToLayoutInDocument", myFrames, ignoredFrames
    Debug.Assert ContainsKey(myFrames, "Figure 1")
    Debug.Assert Not ContainsKey(myFrames, "Figure 8") ' Right aligned
    Debug.Assert Not ContainsKey(myFrames, "Figure 18") ' Reference is in different section.
    Debug.Assert Not ContainsKey(ignoredFrames, "Figure 1")
    Debug.Assert ContainsKey(ignoredFrames, "Figure 8")
    Debug.Assert ContainsKey(ignoredFrames, "Figure 18")
    
    MsgBox ("All tests completed")
End Sub


Private Function ColumnNumber(rng As Range) As Integer
    ColumnNumber = 1 ' default
    Set currentPageSetup = ActiveDocument.Sections(rng.Information(wdActiveEndSectionNumber)).PageSetup
    ' In the left hand column, the distance from the page edge is the distance from the page boundry plus the left margin.
    ' So if we're further away, we're in the right hand column
    If currentPageSetup.TextColumns.Count > 1 And _
           rng.Information(wdHorizontalPositionRelativeToPage) > rng.Information(wdHorizontalPositionRelativeToTextBoundary) + currentPageSetup.LeftMargin + 1 Then
        ColumnNumber = 2
    End If
End Function


Private Sub ShowStatusBarMessage(message As String)
    If message = "" Then
        Application.StatusBar = " "
    Else
        Application.StatusBar = message
    End If
    DoEvents
End Sub

Private Function ContainsKey(col As Collection, key As String) As Boolean
' Answers true if Collection col contains key
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    ContainsKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0 ' Reset
End Function
