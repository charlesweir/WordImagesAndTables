Attribute VB_Name = "FiguresAndTables"
' Functions for image manipulation in Word
'
' Copyright (c) 2020 Charles Weir
'
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
Option Explicit


Sub ButtonPressed(control As IRibbonControl)
    'Callback when a button is pressed on the Ribbon
    
    Dim objUndo As UndoRecord ' Combines all undo information for this activity into one.
    
    'Begin the custom undo record and provide a name for the record
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord (control.ID)

    Select Case control.ID
    Case "Figure", "Table"
        InsertFrame (control.ID)
        
    Case "Reposition"
        If SelectedFloatingShape Is Nothing Then
            MsgBox "Please select a floating image, shape or textbox first."
            Exit Sub
        End If
        RepositionFloatingImage
        
    Case "ChangePicture"
        Debug.Print "ChangePicture called"
        Dim OK As Boolean
        OK = False ' No ElseIf construct in this version of VBA
        If Selection.InlineShapes.Count <> 0 Then
            If Selection.InlineShapes(1).Type = wdInlineShapePicture Then OK = True
        End If
        If Not OK Then
            MsgBox "Please select an inline shape first", vbOKOnly
            Exit Sub
        End If
        
        ChangePicture
        
    Case "RelayoutDocument"
        Dim framesToLayout As New collection
        Dim framesIgnored As New collection
        LayoutFloatingImages.AnalyseImagesToLayoutInDocument framesToLayout, framesIgnored

        
        If MsgBox("Found " & framesToLayout.Count & " Figures and Tables to layout" & vbCrLf & vbCrLf & _
                "Omitting: " & vbCrLf & _
                ListOfFrameNames(framesIgnored) & vbCrLf & _
                "Do you wish to continue?", vbYesNo) <> vbYes Then
            Exit Sub
        End If
        LayoutFloatingImages.LayoutTheseFloatingImages framesToLayout
        
    Case Else
        Debug.Assert False ' Unexpected button ID
    End Select

    'End the custom undo record
    objUndo.EndCustomRecord
    
End Sub


Sub InsertFigure()
' Inserts a floating figure and reference to it.
    InsertFrame "Figure"
End Sub
Sub InsertTable()
' Inserts a floating table and reference to it.
    InsertFrame "Table"
End Sub
Sub CopyImageFormat()
Attribute CopyImageFormat.VB_Description = "Copy Image Format"
Attribute CopyImageFormat.VB_ProcData.VB_Invoke_Func = "Project.FiguresAndTables.CopyImageFormat"
    ' For button or keyboard shortcut (suggest Ctrl-Sh-Cmd C).
    ' Copy the layout of the current image or shape
    
    PreserveImageCroppingAndSizing (False)
End Sub
Sub PasteImageFormat()
    ' For button or keyboard shortcut (suggest Ctrl-Sh-Cmd V).
    ' Paste the layout of the current image or shape
    
    PreserveImageCroppingAndSizing (True)
End Sub

Private Static Function MacFileSelectDialog() As String
    ' Puts up a file selection dialog on Mac, remembering the location between calls.
    ' Answers a string starting "-" on error.
    
    Dim sDefaultLocation As String ' Preserved between calls by the function's Static-ness

    If Trim(sDefaultLocation & vbNullString) = vbNullString Then ' Check for every manner of null string...
        sDefaultLocation = ActiveDocument.Path
        If (sDefaultLocation = vbNullString) Or (sDefaultLocation Like "http*") Then ' New document with no filename, or online document?
            sDefaultLocation = Options.DefaultFilePath(wdDocumentsPath)
        End If
    End If
    
    Dim sMacScript As String
    sMacScript = "try " & vbNewLine & _
        "set theFile to (choose file " & _
        "with prompt ""Please select a file"" default location  """ & _
        sDefaultLocation & """ multiple selections allowed false) as string" & vbNewLine & _
        "on error errStr number errorNumber" & vbNewLine & _
        "return errorNumber " & vbNewLine & _
        "end try " & vbNewLine & _
        "return POSIX path of (theFile as text)"
        
    Debug.Print "Actioning: " & sMacScript
    MacFileSelectDialog = MacScript(sMacScript)
    
    Debug.Print "Returned: " & MacFileSelectDialog
    
    ' Errors are returned as negative numbers:
    If MacFileSelectDialog Like "-*" Then Exit Function
    
    sDefaultLocation = Left(MacFileSelectDialog, InStrRev(MacFileSelectDialog, "/"))

End Function

Private Sub ChangePicture()
' Changes the selected image to a new one chosen by the user, preserving size and cropping
    PreserveImageCroppingAndSizing (False)
    Dim sFileName As String
    
    #If Mac Then
    
        sFileName = MacFileSelectDialog
        If sFileName Like "-*" Then Exit Sub
        
    #Else
    
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        If Application.FileDialog(msoFileDialogOpen).Show = 0 Then Exit Sub
        sFileName = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)

    #End If

    ' This deletes the old image. The selection ends up somewhere after the new image.
    Selection.InlineShapes.AddPicture FileName:=sFileName, LinkToFile:=False, SaveWithDocument:=True

    ' So select the new image:
    Selection.GoToPrevious wdGoToGraphic
    Selection.Expand wdCharacter
    PreserveImageCroppingAndSizing (True)
End Sub

Private Sub InsertFrame(frameType As String)
' Inserts a frame (Figure or Table) at the current selection point.
    Dim newAnchoredFrame As AnchoredFrame
    Set newAnchoredFrame = New AnchoredFrame
    Application.ScreenUpdating = False
    newAnchoredFrame.InitWithNewFrameAt Selection.Range, frameType
    newAnchoredFrame.Update
    Application.ScreenUpdating = True
    Application.ScreenRefresh
End Sub

Private Static Sub PreserveImageCroppingAndSizing(IsPaste)
' Support for updating an image to a new version while preserving cropping and sizing.
' Typically used, for example, with PDFs or jpegs generated by an external tool
'
' Called with IsPaste=False: Takes the selected inline image and copies its size and cropping
' Called with IsPaste=True: Applies the saved size and cropping to the selected image
'
    ' Call is static to preserve value of all variables:
    Dim cropLeft, cropRight, cropTop, cropBottom, pictureHeight, pictureWidth, ScaleHeight, ScaleWidth, Height, Width, lineVisible, lineStyle, lineWeight, lineColor As Variant
    
    With Selection.InlineShapes(1)
        If Not IsPaste Then ' copy
            ScaleHeight = .ScaleHeight
            ScaleWidth = .ScaleWidth
            Height = .Height
            Width = .Width
            With .PictureFormat
                cropLeft = .cropLeft
                cropRight = .cropRight
                cropTop = .cropTop
                cropBottom = .cropBottom
                pictureHeight = .Crop.pictureHeight
                pictureWidth = .Crop.pictureWidth
            End With
            lineVisible = .Line.Visible
            lineStyle = .Line.Style
            lineWeight = .Line.Weight
            lineColor = .Line.ForeColor
        Else ' Paste
            With .PictureFormat
                .cropLeft = cropLeft
                .cropRight = cropRight
                .cropTop = cropTop
                .cropBottom = cropBottom
            End With
            
            ' Now, we want the same width; height the same proportional scaling
            .Width = Width ' Which sets .ScaleWidth, I think:
            .ScaleHeight = ScaleHeight * .ScaleWidth / ScaleWidth
            
            .Line.Visible = lineVisible
            If lineVisible Then
                .Line.Style = lineStyle
                .Line.Weight = lineWeight
                .Line.ForeColor = lineColor
            End If
        End If
    End With
End Sub


Sub RepositionFloatingImage()
'
' Note, doesn't work well as keyboard shortcut, since with keystroke multiple invocations don't seem to work.
'
' Implements support for Latex-like floating pictures near their anchor point.
'
' Specifically, either resets the positioning of the selected floating image or text frame:
'   In two column mode:
'       to be top or bottom of its column or page
'       (first to top, then when called again, to bottom etc)
'   In single column mode:
'      If large, float top or bottom of the page
'      If small, float Left or right, near the anchor
'
    Dim SingleColumnWidth, NumColumns, InsideMarginWidth, MaxSingleColumnFrameWidth As Single
    With Selection.Sections(1).PageSetup
        SingleColumnWidth = .TextColumns.Width
        NumColumns = .TextColumns.Count
        InsideMarginWidth = .PageWidth - .LeftMargin - .RightMargin - .Gutter
        MaxSingleColumnFrameWidth = SingleColumnWidth * 1.05 ' Frames less wide than this are treated as intended to be in a column
    End With
    
    With SelectedFloatingShape
        ' Special: Small frames in a single column page layout toggle left or right near the anchor.
        If NumColumns = 1 And .Width < SingleColumnWidth / 2 Then
            .WrapFormat.Type = wdWrapSquare
            .RelativeVerticalPosition = wdRelativeVerticalPositionLine
            .Top = wdShapeTop  'Or wdShapeCenter, but that looked a bit odd.
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
            .Left = IIf(.Left = wdShapeRight, wdShapeLeft, wdShapeRight)
        Else
            ' Toggle top/bottom of page or column
            .WrapFormat.Type = wdWrapTopBottom
            .RelativeHorizontalPosition = IIf(.Width > MaxSingleColumnFrameWidth, wdRelativeHorizontalPositionMargin, wdRelativeHorizontalPositionColumn)
            .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
            .Left = wdShapeCenter
            .Top = IIf(.Top = wdShapeTop, wdShapeBottom, wdShapeTop)
            ' Now make the frame have the correct width:
            .Width = IIf(NumColumns = 2 And .Width > MaxSingleColumnFrameWidth, InsideMarginWidth, SingleColumnWidth)
        End If
        
        ' And set height to fit the contents:
        .TextFrame.AutoSize = -1
        .TextFrame.WordWrap = -1
        
        ' And lock the anchor. It makes dragging not work, but we don't want that to happen by accident.
        .LockAnchor = -1
        
        ' Make it visible
        ActiveWindow.ScrollIntoView .TextFrame.TextRange
        
        ' Selection seem to go wrong. Put the cursor at the start of the frame so doing it again will work even with a keystroke.
        Dim newSelection As Range
        Set newSelection = .TextFrame.TextRange
        newSelection.Collapse
        newSelection.Select
        
    End With
    
End Sub

Public Function SelectedFloatingShape() As Shape
    ' Answers the floating shape intended by the user
    ' either a selected floating shape or frame, or a text frame containing the cursor
    If Selection.ShapeRange.Count > 0 Then
        Set SelectedFloatingShape = Selection.ShapeRange(1)
        Exit Function
    End If
    
    ' Find the text frame containing the selection. No easy way I've found...
    Dim currentFrame As Shape
    For Each currentFrame In ActiveDocument.StoryRanges(wdMainTextStory).ShapeRange
        If currentFrame.Type = msoTextBox Then
            If Selection.inRange(currentFrame.TextFrame.TextRange) Then
                Set SelectedFloatingShape = currentFrame
                Exit Function
            End If
        End If
    Next currentFrame
    
    Set SelectedFloatingShape = Nothing
End Function

Private Function ListOfFrameNames(frames As collection) As String
    ' Answers a string listing all the frame names in the given collection.
    ListOfFrameNames = ""
    Dim frame As AnchoredFrame
    For Each frame In frames
        ListOfFrameNames = ListOfFrameNames & frame.Name & vbCrLf
    Next

End Function
