Attribute VB_Name = "GatherNotesv2"
Option Explicit

Sub GatherAllNotesInWordFilev2()

    Dim path, name, docName, slideNumber As String
    Dim WordBasic, myDoc As Object
    
    'path = InputBox("path to save")
    path = ActivePresentation.path
    name = ActivePresentation.name
    docName = path & "\" & name & "V2.docx"
    
    ' Start Microsoft Word.
    Set WordBasic = CreateObject("Word.Application")
    ' Add a new document
    WordBasic.Documents.Add
    Set myDoc = WordBasic.Documents(1)
    myDoc.Activate
    
    Dim allSlides, mySlide, notePageShape As Object
    Dim NotesText As String
    Dim currentSlideNumber As Integer
    ' Go through each slides
    ' check if any shape in the notesPage has text, and if it is the case
    ' copy it on the Word document using SaveNote
    currentSlideNumber = 2
    Set allSlides = ActivePresentation.Slides
    For Each mySlide In allSlides
        ' We first update the slide number if necessary
        slideNumber = CStr(mySlide.slideIndex)
        FindNextSlidesNumbered mySlide:=mySlide, _
            currentSlideNumber:=currentSlideNumber, currentText:=slideNumber
        For Each notePageShape In mySlide.NotesPage.Shapes
            ' We are interested only in shapes that have a text frame
            ' and we exclude the shape that holds only the slide number as it is not interesting
            If notePageShape.HasTextFrame And _
                notePageShape.PlaceholderFormat.Type <> ppPlaceholderSlideNumber Then
                
                If notePageShape.TextFrame.HasText Then
                    NotesText = notePageShape.TextFrame.TextRange.text
                    SaveNote myDoc:=myDoc, slideNumber:=slideNumber, text:=NotesText
                    
                End If
            End If
        Next
    Next
    
    myDoc.SaveAs2 FileName:=docName
    WordBasic.Quit
End Sub


' Save the note "text" on slide "slideNumber" in the document "myDoc"
Function SaveNote(myDoc, slideNumber, text)
    Dim rangeToWrite As Object
    Dim paragraphCounts, lastParagraphEnd As Long
    
    ' We first write the index of the slide in bold
    paragraphCounts = myDoc.Paragraphs.Count
    lastParagraphEnd = myDoc.Paragraphs(paragraphCounts).Range.End - 1
    Set rangeToWrite = myDoc.Range(Start:=lastParagraphEnd, End:=lastParagraphEnd)
    With rangeToWrite
        .InsertAfter text:="Slide " & slideNumber
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



' Look at a given slide to see whether a given slide number
' is present. If not, do nothing. If yes, increase the slide
' number and add it to the text that it is given in parameter
' Note that this slide number does not correspond to the real slide number
' of mySlide in the presentation.
' This is because some slides may not have a number and are therefore skipped.
Function FindNextSlidesNumbered(mySlide, currentSlideNumber, currentText)

    Dim text As String
    Dim myShape As Object

    For Each myShape In mySlide.Shapes
        ' We are interested only in shapes that have a text frame
        ' and we exclude the shape that holds only the slide number as it is not interesting
        If myShape.HasTextFrame Then
            If myShape.TextFrame.HasText Then
                If IsNumeric(myShape.TextFrame.TextRange.text) Then
                    If CInt(myShape.TextFrame.TextRange.text) = currentSlideNumber Then
                        ' We found it!
                        currentText = currentText + " -> (" + myShape.TextFrame.TextRange.text + ") "
                        currentSlideNumber = currentSlideNumber + 1
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function


