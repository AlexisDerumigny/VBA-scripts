Attribute VB_Name = "GatherNotes"
Option Explicit

Sub GatherAllNotesInWordFile()

    Dim path, name, docName As String
    Dim WordBasic, myDoc As Object
    
    'path = InputBox("path to save")
    path = ActivePresentation.path
    name = ActivePresentation.name
    docName = path & "\" & name & ".docx"
    
    ' Start Microsoft Word.
    Set WordBasic = CreateObject("Word.Application")
    ' Add a new document
    WordBasic.Documents.Add
    Set myDoc = WordBasic.Documents(1)
    myDoc.Activate
    
    Dim allSlides, mySlide, notePageShape As Object
    Dim NotesText As String
    ' Go through each slides
    ' check if any shape in the notesPage has text, and if it is the case
    ' copy it on the Word document using SaveNote
    Set allSlides = ActivePresentation.Slides
    For Each mySlide In allSlides
        For Each notePageShape In mySlide.NotesPage.Shapes
            ' We are interested only in shapes that have a text frame
            ' and we exclude the shape that holds only the slide number as it is not interesting
            If notePageShape.HasTextFrame And _
                notePageShape.PlaceholderFormat.Type <> ppPlaceholderSlideNumber Then
                
                If notePageShape.TextFrame.HasText Then
                
                    NotesText = notePageShape.TextFrame.TextRange.text
                    SaveNote myDoc:=myDoc, slideIndex:=mySlide.slideIndex, text:=NotesText
                    
                End If
            End If
        Next
    Next
    
    myDoc.SaveAs2 FileName:=docName
    WordBasic.Quit
End Sub


' Save the note "text" on slide "slideIndex" in the document "myDoc"
Function SaveNote(myDoc, slideIndex, text)
    Dim rangeToWrite As Object
    Dim paragraphCounts, lastParagraphEnd As Long
    
    ' We first write the index of the slide in bold
    paragraphCounts = myDoc.Paragraphs.Count
    lastParagraphEnd = myDoc.Paragraphs(paragraphCounts).Range.End - 1
    Set rangeToWrite = myDoc.Range(Start:=lastParagraphEnd, End:=lastParagraphEnd)
    With rangeToWrite
        .InsertAfter text:="Slide " & CStr(slideIndex)
        .InsertParagraphAfter
        With .Font
            .Size = 12
            .Bold = True
        End With
    End With
    
    ' Then we write the "text" argument
    paragraphCounts = myDoc.Paragraphs.Count
    lastParagraphEnd = myDoc.Paragraphs(paragraphCounts).Range.End - 1
    Set rangeToWrite = myDoc.Range(Start:=lastParagraphEnd, End:=lastParagraphEnd)
    With rangeToWrite
        .InsertAfter text:=text
        .InsertParagraphAfter
        With .Font
            .Size = 12
            .Bold = False
        End With
    End With
End Function


