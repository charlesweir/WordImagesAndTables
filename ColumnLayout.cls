VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    ' ColumnLayout: Represents the layout for the current column.
Option Explicit

Private Enum PageFillStatus
    pfsUnfilled = 0 ' No frames in this page
    pfsTopFilled = 1 ' Frame at the top
    pfsBottomFilled = 2 ' Frame at the top and the bottom
    pfsOverFilled = 3 ' Definitely need to position on the next page.
End Enum

Dim workingPage As Integer
Dim workingPageStatus As PageFillStatus
Dim workingPosition As Range
Dim nextPage As Range
Dim spaceUsedOnCurrentPage As Single

    
Sub Initialise()
    workingPage = 0
    workingPageStatus = pfsUnfilled
End Sub

        
Sub PositionFrame(newFrame As anchoredFrame)
    With newFrame
        ShowStatusBarMessage ("Positioning " & .Name)
        
        ' Work out where to try to put it:
        If .ReferencePageNumber > workingPage Then
            workingPage = .ReferencePageNumber
            workingPageStatus = pfsUnfilled
            Set workingPosition = .PlaceNearReference
        ElseIf workingPageStatus = pfsTopFilled Then
            ' If the reference is ahead of the working position, move the working position forward.
            If workingPosition.Start < .ReferenceFieldStart.Start Then
                Set workingPosition = .PlaceNearReference
            Else
                workingPosition.Move wdCharacter ' Move past the last paste.
            End If
        ElseIf workingPageStatus = pfsBottomFilled Then
            workingPageStatus = pfsOverFilled ' We'll have to put it on the next page.
        End If
                    
        ' If we have a top or bottom slot free:
        If (workingPageStatus = pfsUnfilled) Or (workingPageStatus = pfsTopFilled) Then
        
            ' Try moving it to the working position, at the top or bottom of the page:
            .MoveAnchorTo workingPosition
    
            ' Todo: Add special case. Top of page is continued table, in which case we want to be bottom.
            If workingPageStatus = pfsUnfilled Then
                .SetToFloatAt (wdShapeTop)
                workingPageStatus = pfsTopFilled
            Else
                .SetToFloatAt (wdShapeBottom)
                workingPageStatus = pfsBottomFilled
            End If
    
            '  If the anchor is now on a different page number, or Word has misbehaved and put both overlapping anyway.
            ' Word also seems to behave oddly when there's two frames and only one line of text (and it looks pretty odd too):
            ' Let's say we we want at least 10% of page for text when there are two frames on the page.
            If (.FramePageNumber > workingPage) _
                 Or ((workingPageStatus = pfsBottomFilled) And (spaceUsedOnCurrentPage + .Height > (.UsableDocumentHeight * 0.9))) Then
                ' And we'll treat it as though we had no space in the first place.
                workingPageStatus = pfsOverFilled
            End If
        End If
        
        ' If there wasn't space enough on that page, put it on the next one:
        If workingPageStatus = pfsOverFilled Then
            .Stash ' Get it out of the way.
            
            ' Find the first para (line?) on the next page
            ' Skip over any tables, as tables make a right mess.
            Debug.Assert workingPosition.Information(wdActiveEndPageNumber) = workingPage
            Set nextPage = workingPosition.GoTo(wdGoToPage, wdGoToNext)
            ' Kludge. The above sometimes goes wrong (but not when stepping through in the debugger):
            If nextPage.Information(wdActiveEndPageNumber) = workingPage + 1 Then
                Set workingPosition = nextPage
            Else
                While workingPosition.Information(wdActiveEndPageNumber) = workingPage
                    workingPosition.Move wdParagraph
                Wend
            End If
            
            While workingPosition.Information(wdWithInTable) _
                     And (Not workingPosition.Next(Unit:=wdParagraph) Is Nothing)
                workingPosition.Move wdParagraph
            Wend
    
            ' N.B. There might not be enough paras to take all the images. Also we might end up in the middle of the bibliography.
            ' Both would be unusual in academic papers. Forget it.
            
            ' And move it there
            .MoveAnchorTo workingPosition
            .SetToFloatAt (wdShapeTop)
            workingPageStatus = pfsTopFilled
        End If
        
        workingPage = workingPosition.Information(wdActiveEndPageNumber)
        If workingPageStatus = pfsTopFilled Then
            spaceUsedOnCurrentPage = .Height
        Else
            spaceUsedOnCurrentPage = 0
        End If
    End With

End Sub

Private Property Get NumColumns() As Integer
    NumColumns = CurrentPageSetup.TextColumns.count
End Property

Private Property Get CurrentPageSetup() As pageSetup
    Set CurrentPageSetup = ActiveDocument.Sections(workingPosition.Information(wdActiveEndSectionNumber)).pageSetup
End Property


Private Function ColumnNumber(rng As Range) As Integer
    ' Answers the column number of rng (1 or 2)
    ColumnNumber = 1 ' default
    ' In the left hand column, the distance from the page edge is the distance from the page boundry plus the left margin.
    ' So if we're further away, we're in the right hand column
    If NumColumns > 1 And _
           rng.Information(wdHorizontalPositionRelativeToPage) > rng.Information(wdHorizontalPositionRelativeToTextBoundary) + CurrentPageSetup.LeftMargin + 1 Then
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

