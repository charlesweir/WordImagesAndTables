VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnchoredFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' AnchoredFrame:
' A Floating Frame containing an caption (Figure or Table), plus the main reference in the document to that caption.
Option Explicit

Private myFrame As Shape
Private myReferencingField As field
Private myOriginalSectionNumber As Integer ' we get moved to the document end. Remember our section number.

Public Property Set Frame(newFrame As Shape)
    Set myFrame = newFrame
    myOriginalSectionNumber = myFrame.Anchor.Information(wdActiveEndSectionNumber)
End Property

Public Property Set RefField(newField As field)
    Set myReferencingField = newField
End Property

Public Property Get Height() As Single
    Height = myFrame.Height
End Property

Public Property Get FramePageNumber() As Integer
    ' Answers the page number containing the frame.
    FramePageNumber = myFrame.Anchor.Information(wdActiveEndPageNumber)
End Property

Public Property Get ReferencePageNumber() As Integer
    ' Answers the page number containing the (end of the) referencing field.
    ReferencePageNumber = myReferencingField.Result.Information(wdActiveEndPageNumber)
End Property

Public Property Get ReferenceColumnNumber() As Integer
    ' Answers the column number containing the referencing field.
    ReferenceColumnNumber = ColumnNumberFor(myReferencingField)
End Property

Public Property Get SectionNumber() As Integer
    ' Answers the page number containing the frame.
    SectionNumber = myOriginalSectionNumber
End Property

Public Property Get NumColumns() As Integer
    ' Answers the number of text columns on the page (1 or 2)
    NumColumns = CurrentPageSetup.TextColumns.count
    Debug.Assert NumColumns <= 2
End Property

Private Property Get CurrentSection() As Section
    ' The Word Section object containing the frame.
    Set CurrentSection = ActiveDocument.Sections(SectionNumber)
End Property

Private Property Get CurrentPageSetup() As pageSetup
    Set CurrentPageSetup = CurrentSection.pageSetup
End Property

Public Property Get PlaceNearReference() As Range
    ' Answers a safe place to put an anchor beside the primary reference.
    Dim workingPosition As Range
    Set workingPosition = myReferencingField.Result
    workingPosition.Collapse (wdCollapseEnd)
    Set workingPosition = workingPosition.Next(Unit:=wdCharacter)
    workingPosition.Collapse
    Set PlaceNearReference = workingPosition
End Property
            
Public Property Get ReferenceFieldStart() As Range
    ' Answers the location at the start of the reference field
    Dim workingPosition As Range
    Set workingPosition = myReferencingField.Result
    workingPosition.Collapse
    Set ReferenceFieldStart = workingPosition
End Property

Public Property Get FirstParaOnSamePage() As Range
    ' Answers the first location at the top of the page containing the frame.
    Dim firstPara As Range
    Set firstPara = myFrame.Anchor
    Set FirstParaOnSamePage = firstPara.GoTo(What:=wdGoToPage, count:=myFrame.Anchor.Information(wdActiveEndPageNumber))
End Property
 
Private Property Get FrameBookmark() As Bookmark
    ' The bookmark in the frame caption.
    Dim bookmarkSet As Bookmarks
    Set bookmarkSet = myFrame.TextFrame.TextRange.Bookmarks
    bookmarkSet.ShowHidden = True
    Debug.Assert bookmarkSet.count > 0 And bookmarkSet(1).Name Like "_Ref##*"
    Set FrameBookmark = bookmarkSet(1)
End Property

Private Property Get FirstCaption() As Range
    ' Answers full paragraph that is the first caption in the frame.
    Dim para As Paragraph
    For Each para In myFrame.TextFrame.TextRange.Paragraphs
        If para.style = "Caption" Then GoTo FoundPara ' Could be Like, if several caption types
    Next para: Debug.Assert False ' No caption found
FoundPara:
    Set FirstCaption = para.Range
End Property

Public Property Get Name() As String
    ' Answers the name, such as "Figure 1" or "Table 3"
    Name = FrameBookmark.Range.Text
End Property

Public Property Get UsableDocumentHeight() As Single
    ' Answers the height of the formatted region containing the frame
    With CurrentPageSetup
        UsableDocumentHeight = .PageHeight - .TopMargin - .BottomMargin
    End With
End Property

Public Property Get BookmarkId() As String
    ' Answers the bookmark ID, such as "_Ref12345678"
    BookmarkId = FrameBookmark.Name
End Property

Public Property Get FitsInColumn() As Boolean
    ' Does the frame fit in a single column?
    ' Answers true for single column mode in case an image is a little wide for the text.
    FitsInColumn = (NumColumns = 1) Or (myFrame.Width < CurrentPageSetup.TextColumns.Width)
End Property

Public Sub MoveAnchorTo(workingPosition As Range)
    ' Moves the anchor of the frame to workingPosition, or as soon after it as possible.
    myFrame.Anchor.Cut
    workingPosition.Collapse (wdCollapseEnd) ' After previous pastes.
    
    ' If we're in a field, the frame could be wiped out by an update.
    While workingPosition.Information(wdInFieldCode) Or workingPosition.Information(wdInFieldResult)
        Set workingPosition = workingPosition.Next(Unit:=wdCharacter)
        workingPosition.Collapse wdCollapseEnd
    Wend
    workingPosition.Paste
End Sub

Public Sub Stash()
    ' Move the frame out of the way so it doesn't affect the layout process.
    If FramePageNumber <> ActiveDocument.Content.Information(wdActiveEndPageNumber) Then ' Not already stashed?
        MoveAnchorTo ActiveDocument.Content
    End If
End Sub

Public Sub SetToFloatAt(topOrBottom As Long)
    ' Sets the frame to float at the center of the top or bottom of the page
    With myFrame
        ' Optimisation. Word re-lays out the document each time you set any of these, whether or not the value has changed. (actually doesn't make much difference)
        If .RelativeVerticalPosition <> wdRelativeVerticalPositionMargin Then .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
        .RelativeHorizontalPosition = IIf(FitsInColumn And (NumColumns > 1), wdRelativeHorizontalPositionColumn, wdRelativeHorizontalPositionMargin)
        If .Left <> wdShapeCenter Then .Left = wdShapeCenter
        If .WrapFormat.Type <> wdWrapTopBottom Then .WrapFormat.Type = wdWrapTopBottom
        If .Top <> topOrBottom Then .Top = topOrBottom
    End With
End Sub

Public Sub Update()
    ' Updates all fields in the frame, and the anchor field.
    myFrame.TextFrame.TextRange.Fields.Update
    myReferencingField.Update
End Sub
       
Public Sub InitWithNewFrameAt(oRng As Range, autotextName As String)
' Set myself up as a a new frame and reference to it, based on the given autotext (Figure or Table)
    Application.ScreenUpdating = False
    Dim rngInsertedAnchor As Range
    Set rngInsertedAnchor = ActiveDocument.AttachedTemplate.AutoTextEntries(autotextName).Insert(where:=oRng, RichText:=True)
    Debug.Assert rngInsertedAnchor.ShapeRange.count = 1 ' Autotext contains floating frame
    Set Frame = rngInsertedAnchor.ShapeRange(1)
    InsertCrossReferenceBesideAnchor autotextName
    Application.ScreenUpdating = True
    Application.ScreenRefresh
End Sub

Private Sub InsertCrossReferenceBesideAnchor(autotextName As String)
' Inserts a reference field, such as "Figure 2" or "Table 7", just after the anchor.
'
' Word VBA doesn't support cross references to text in text frames (not text boxes), so we
' move the caption to the main document, create the reference, then move it back again.
'
' To find the caption in the list of cross-referenceable items, we add a random tag to it.

    Dim tag As String
    tag = "TAG" & Int(99999999 * Rnd)
    
    ' Remember where the caption was.
    Dim captionLocation As Range
    Set captionLocation = FirstCaption
    If autotextName = "Figure" Then captionLocation.InsertAfter (vbCrLf)
    captionLocation.Collapse

    ' Move the caption para into the main document and tag it
    Dim theCaption As Range
    Set theCaption = FirstCaption
    theCaption.Cut
    Set theCaption = ActiveDocument.Content
    theCaption.Collapse wdCollapseEnd
    theCaption.InsertBefore vbCrLf
    theCaption.Collapse wdCollapseEnd
    theCaption.Paste
    Dim tagLocation As Range
    Set tagLocation = theCaption
    tagLocation.Collapse wdCollapseEnd
    tagLocation.Move wdCharacter, -1
    tagLocation.InsertAfter tag ' sets tagLocation to the range of the inserted tag

    ' Create the cross reference to it after the anchor:
    Dim items() As String
    items = ActiveDocument.GetCrossReferenceItems(autotextName)
    ' Answers an array of strings, "123. Figure 1/ The Caption", plus all the numbered paras etc.
    Dim crossRefItem As Integer
    For crossRefItem = LBound(items) To UBound(items)
        If InStr(items(crossRefItem), tag) <> 0 Then GoTo ItemFound
    Next crossRefItem: Debug.Assert False ' Tag not found. Ouch!
ItemFound:

    Dim afterAnchor As Range
    Set afterAnchor = myFrame.Anchor
    afterAnchor.Collapse wdCollapseEnd
    
    afterAnchor.InsertCrossReference ReferenceType:=autotextName, ReferenceKind:= _
        wdOnlyLabelAndNumber, ReferenceItem:="" & crossRefItem, InsertAsHyperlink:=False, _
        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
    
    afterAnchor.Expand wdSentence
    Set RefField = afterAnchor.Fields(1) ' Might not actually be correct, but who cares? We only use it for updating...
    
    ' There's a bizarre Word error here, where the caption paragraph isn't editable for a short time after the InsertCrossReference command.
    ' Turns out we can simply wait for Word to sort it out...
    On Error Resume Next
    Do
        DoEvents
        tagLocation.Delete
    Loop While Err.Number = 4605 Or Err.Number = 91
    On Error GoTo 0
    
    ' Finally, move the caption para back, and tidy up.
    theCaption.Expand wdParagraph
    theCaption.Cut
    theCaption.Delete wdCharacter, Len(vbCrLf)   ' Ending CRLF
    captionLocation.Paste
    ' In the case of the table, we somehow lose the caption style (not sure why). Just put it back.
    captionLocation.style = "Caption"
    captionLocation.Collapse wdCollapseEnd
    captionLocation.Delete wdCharacter ' the following empty paragraph.
End Sub
' Static
Public Function IsValidFrame(shp As Shape) As Boolean
    ' Is this a text frame containing a caption that we need to layout?
    IsValidFrame = False
    If shp.Type <> msoTextBox Then Exit Function
    
    ' Don't layout frames aligned left or right.
    If shp.Left = wdShapeLeft Or shp.Left = wdShapeRight Then Exit Function
    
    Dim bookmarkSet As Bookmarks
    Set bookmarkSet = shp.TextFrame.TextRange.Bookmarks
    bookmarkSet.ShowHidden = True
    If bookmarkSet.count = 0 Then Exit Function
    If Not bookmarkSet(1).Name Like "_Ref##*" Then Exit Function
    
    IsValidFrame = True
End Function

' Static
Public Function BookmarkIdFromField(fld As field) As String
    ' If this field is a reference to a caption, answers the reference key, else null string.
    BookmarkIdFromField = ""
    If fld.Type <> wdFieldRef Then Exit Function
    If Not fld.Code.Text Like "* REF _Ref##* *" Then Exit Function
    BookmarkIdFromField = Split(fld.Code.Text, " ")(2)
End Function
 
' Static
Public Function ColumnNumberFor(rng As Range) As Integer
    ' Answers the column number of rng (1 or 2)
    ColumnNumberFor = 1 ' default
    ' In the left hand column, the distance from the page edge is the distance from the page boundry plus the left margin.
    ' So if we're further away, we're in the right hand column
    Dim ps As pageSetup
    Set ps = ActiveDocument.Sections(rng.Information(wdActiveEndSectionNumber)).pageSetup
    
    If ps.TextColumns.count > 1 And _
            rng.Information(wdHorizontalPositionRelativeToPage) > rng.Information(wdHorizontalPositionRelativeToTextBoundary) + _
                    ps.LeftMargin + 1 Then
        ColumnNumberFor = 2
    End If
End Function

            
