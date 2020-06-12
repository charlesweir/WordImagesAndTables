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

Public Property Set Frame(newFrame As Shape)
    Set myFrame = newFrame
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
    ' Answers the page number containing the frame.
    ReferencePageNumber = myReferencingField.Result.Information(wdActiveEndPageNumber)
End Property

Public Property Get SectionNumber() As Integer
    ' Answers the page number containing the frame.
    SectionNumber = myFrame.Anchor.Information(wdActiveEndSectionNumber)
End Property
Private Property Get NumColumns() As Integer
    ' Answers the number of text columns on the page (1 or 2)
    NumColumns = CurrentPageSetup.TextColumns.count
    Debug.Assert NumColumns <= 2
End Property

Private Property Get CurrentSection() As Section
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
    ' Answers the text (not numbering) part of the first caption in the frame.
    Dim para As Paragraph
    For Each para In myFrame.TextFrame.TextRange.Paragraphs
        If para.style = "Caption" Then GoTo FoundPara ' Could be Like, if several caption types
    Next para: Debug.Assert False ' No caption found
FoundPara:
    Set FirstCaption = para.Range
    FirstCaption.Collapse
    FirstCaption.MoveWhile " FigureTabl" ' Skip the Figure or table part
    Do While FirstCaption.Information(wdInFieldCode)
        FirstCaption.Move wdCharacter
    Loop
    FirstCaption.MoveEnd wdParagraph
    FirstCaption.MoveEnd wdCharacter, -1 ' Not cr at end.
End Property

Public Property Get Name() As String
    ' Answers the name, such as "Figure 1" or "Table 3"
    Name = FrameBookmark.Range.Text ' Gives "Figure NN"
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
    FitsInColumn = myFrame.Width < CurrentPageSetup.TextColumns.Width
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
    'Debug.Assert Not (workingPosition.Information(wdInFieldCode) Or workingPosition.Information(wdInFieldResult))
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
       
' Static
Public Sub InitWithNewFrameAt(oRng As Range, autotextName As String)
    ' Create a new frame and reference to it at the given point by inserting the specified autotext (which must be a floating frame)
    ' ToDo Can't get word to omit the caption from the
    Dim rngInsertedAnchor As Range
    Set rngInsertedAnchor = ActiveDocument.AttachedTemplate.AutoTextEntries(autotextName).Insert(where:=oRng, RichText:=True)
    Debug.Assert rngInsertedAnchor.ShapeRange.count = 1 ' Autotext contains floating frame
    Set Frame = rngInsertedAnchor.ShapeRange(1)
    InsertCrossReferenceBesideAnchor
End Sub

Private Sub InsertCrossReferenceBesideAnchor()
    Dim oldCaption As String
    oldCaption = FirstCaption.Text
    FirstCaption.Text = "TAG" & Int(99999999 * Rnd)
    Dim items() As String
    items = ActiveDocument.GetCrossReferenceItems("Figure")
    ' Answers an array of strings, "123. Figure 1/ The Caption", plus all the numbered paras etc.
    
    Dim item As Variant
    Dim searchString As String
    searchString = "*" & FirstCaption.Text & "*"
    For Each item In items
        If item Like searchString Then GoTo ItemFound
    Next item: Debug.Assert False ' Tag not found. Ouch!
ItemFound:

        ' Also wdCaptionTable
    Dim afterAnchor As Range
    afterAnchor.Collapse
    
    afterAnchor.InsertCrossReference ReferenceType:=wdCaptionFigure, ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=item
    Set RefField = afterAnchor.Fields(1)
    myFrame.TextFrame.TextRange.Fields.Update
    myReferencingField.Update
End Sub
' Static
Public Function IsValidFrame(shp As Shape) As Boolean
    ' Is this a text frame containing a caption that we need to layout?
    If shp.Type <> msoTextBox Then IsValidFrame = False: Exit Function
    
    ' Don't layout frames aligned left or right.
    If shp.Left = wdShapeLeft Or shp.Left = wdShapeRight Then IsValidFrame = False: Exit Function
    
    Dim bookmarkSet As Bookmarks
    Set bookmarkSet = shp.TextFrame.TextRange.Bookmarks
    bookmarkSet.ShowHidden = True
    If bookmarkSet.count = 0 Then IsValidFrame = False: Exit Function
    If Not bookmarkSet(1).Name Like "_Ref##*" Then IsValidFrame = False: Exit Function
    
    IsValidFrame = True
End Function

' Static
Public Function BookmarkIdFromField(fld As field) As String
    ' If this field is a reference to a caption, answers the reference key, else null string.
    If fld.Type <> wdFieldRef Then BookmarkIdFromField = "": Exit Function
    If Not fld.Code.Text Like "* REF _Ref##* *" Then BookmarkIdFromField = "": Exit Function
    BookmarkIdFromField = Split(fld.Code.Text, " ")(2)
End Function
 

 ' shp.Top = IIf(testLocation = "top", wdShapeTop, wdShapeBottom)
            
            
