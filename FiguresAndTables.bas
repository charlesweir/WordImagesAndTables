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
        Dim framesToLayout As Collection
        Set framesToLayout = LayoutFloatingImages.ImagesToLayoutInDocument
        If MsgBox("Found " & framesToLayout.Count & " Figures and Tables to layout" & vbCrLf & vbCrLf & _
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
    Dim OldPicture As Variant
    Dim NewPicture As Variant
    Dim Where As Range
    Set OldPicture = Selection.InlineShapes(1)
    
    
    Dim sFileName As String
    
    #If Mac Then
    
        sFileName = MacFileSelectDialog
        If sFileName Like "-*" Then Exit Sub
        
    #Else
    
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        If Application.FileDialog(msoFileDialogOpen).Show = 0 Then Exit Sub
        sFileName = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)

    #End If

    ' Add the new image before the old one, keeping the old one
    Set Where = Selection.Range.Duplicate
    Where.Collapse wdCollapseStart
    Set NewPicture = Selection.InlineShapes.AddPicture(FileName:=sFileName, LinkToFile:=False, SaveWithDocument:=True, _
                                        Range:=Where)

    PreserveImageCroppingAndSizing OldPicture, NewPicture
    ' Now delete it.
    OldPicture.Range.Delete
    ' And select the new image
    NewPicture.Range.Select
    ' And oddly, sometimes we seem to scroll away, so
    ActiveWindow.ScrollIntoView Selection.Range, True
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

Private Sub PreserveImageCroppingAndSizing(OP As Variant, NP As Variant)
'
' Copies the size and other information from inline image OP to inline image NP

    Dim i As Integer

    With NP
        With .PictureFormat
            .cropLeft = OP.PictureFormat.cropLeft
            .cropRight = OP.PictureFormat.cropRight
            .cropTop = OP.PictureFormat.cropTop
            .cropBottom = OP.PictureFormat.cropBottom
        End With
        .ScaleHeight = OP.ScaleHeight
        .ScaleWidth = OP.ScaleWidth
        .Height = OP.Height
        .Width = OP.Width
        ' Copying the borders doesn't work. No obvious reason why not (https://shaunakelly.com/word/formatting/border-basics.html)
        ' but setting LineWidth always gives error 5843
        ' Even without setting it, we don't get a border.
        'For i = wdBorderTop To wdBorderRight Step -1: ' -4 to -1
            '.Borders(i).Visible = OP.Borders(i).Visible
            '.Borders(i).LineStyle = OP.Borders(i).LineStyle
            'If OP.Borders(i).LineStyle <> wdLineStyleNone Then
                '.Borders(i).Color = OP.Borders(i).Color
                '.Borders(i).LineWidth = OP.Borders(i).LineWidth
            'End If
        'Next i
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

    Dim AnchorParagraph As Paragraph

    With SelectedFloatingShape
        If Selection.Sections(1).PageSetup.TextColumns.Count > 1 Then
            ' Column layout. In column if small enough, else page. Toggle top/bottom
            Dim MaxSingleColumnImageWidth As Single
            MaxSingleColumnImageWidth = (Selection.Sections(1).PageSetup.TextColumns.Width * 1.05) ' Little bit of leeway.
            .WrapFormat.Type = wdWrapTopBottom
            If .Width > MaxSingleColumnImageWidth Then
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
            Else
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
            End If
            .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
            .Left = wdShapeCenter
            If .Top = wdShapeTop Then
                .Top = wdShapeBottom
            Else
                .Top = wdShapeTop
            End If
        Else
            ' One column.
            Dim HalfPageWidth As Single
            HalfPageWidth = Selection.Sections(1).PageSetup.TextColumns.Width / 2
            If .Width < HalfPageWidth Then
                'Small picture. Put near anchor, wrap around. Toggle left/right
                .WrapFormat.Type = wdWrapSquare
                .RelativeVerticalPosition = wdRelativeVerticalPositionLine
                .Top = wdShapeTop  'Or wdShapeCenter, but that looked a bit odd.
                
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
                If .Left = wdShapeRight Then
                    .Left = wdShapeLeft
                Else
                    .Left = wdShapeRight
                End If
            Else
                ' Big picture: Toggle top/bottom of page
                .WrapFormat.Type = wdWrapTopBottom
                .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
                .Left = wdShapeCenter
                If .Top = wdShapeTop Then
                    .Top = wdShapeBottom
                Else
                    .Top = wdShapeTop
                End If
            End If
        End If
            
    End With

End Sub

Private Function SelectedFloatingShape() As Shape
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
