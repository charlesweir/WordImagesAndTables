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
    ' And gives a 5% leeway, since there's usually a bit of whitespace in a frame.
    FitsInColumn = (NumColumns = 1) Or (myFrame.Width < (CurrentPageSetup.TextColumns.Width * 1.05))
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
    Dim rngInsertedAnchor As Range
    Set rngInsertedAnchor = ActiveDocument.AttachedTemplate.AutoTextEntries(autotextName).Insert(where:=oRng, RichText:=True)
    Debug.Assert rngInsertedAnchor.ShapeRange.count = 1 ' Autotext contains floating frame
    Set Frame = rngInsertedAnchor.ShapeRange(1)
    InsertCrossReferenceBesideAnchor autotextName
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
    ' Turns out in Windows we can simply wait for Word to sort it out...
    On Error Resume Next
    Do
        DoEvents
        tagLocation.Delete
        Debug.Assert Err.Number = 4605 Or Err.Number = 5904 Or Err.Number = 0 ' or 91?  4605 is Windows, 5904 is Mac, but doesn't seem to clear - Ignore it.
    Loop While Err.Number = 4605
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
       
' Static
Function RepositionableFramesInRegion(region As Range) As Collection
' Answers a collection of relevant AnchoredFrames - ones with references - in the given region.

    Dim oAnchoredFrame As AnchoredFrame
    Dim allFramesInRegion As Collection
    Set allFramesInRegion = New Collection
    
    Set RepositionableFramesInRegion = New Collection
    
    ' First, find all the frames and their hidden cross reference bookmarks:
    Dim currentFrame As Shape
    For Each currentFrame In region.ShapeRange
        Debug.Assert Not currentFrame Is Nothing ' If this fires, Word is still sorting out itself. Wait and try again later.
        If IsValidFrame(currentFrame) Then
            Set oAnchoredFrame = New AnchoredFrame
            Set oAnchoredFrame.Frame = currentFrame
            allFramesInRegion.Add Item:=oAnchoredFrame, key:=oAnchoredFrame.BookmarkId
        End If
    Next
    
    ' Now look for the first reference to each of those bookmarks, and construct our collection with them:
    Dim ReferencingField As field
    Dim bookmarkName As String
    Dim previousField As field
    
    For Each ReferencingField In region.Fields
        ' Word can take several minutes to sort the fields after you open a new document.
        ' If this next assertion fails, get on with something else and come back to do this again later.
        If Not previousField Is Nothing Then Debug.Assert ReferencingField.Result.Start >= previousField.Result.Start
        
        bookmarkName = BookmarkIdFromField(ReferencingField)
        If bookmarkName <> "" And ContainsKey(allFramesInRegion, bookmarkName) Then
            Set oAnchoredFrame = allFramesInRegion(bookmarkName)
            
            ' Only pair references within a section. (I have forward references to figures at the start of my Thesis)
            If (ReferencingField.Code.Information(wdActiveEndSectionNumber) = oAnchoredFrame.SectionNumber) Then
                Set oAnchoredFrame.RefField = ReferencingField
                RepositionableFramesInRegion.Add oAnchoredFrame
                ' And remove it, so that we ignore later references.
                allFramesInRegion.Remove bookmarkName
            End If
        End If
        Set previousField = ReferencingField
    Next ReferencingField
End Function

' Static
Private Function ContainsKey(col As Collection, key As String) As Boolean
' Answers true if Collection col contains key
    On Error Resume Next
    col (key) ' Just try it. If it fails, Err.Number will be nonzero.
    ContainsKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0 ' Reset
End Function
