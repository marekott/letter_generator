Attribute VB_Name = "LetterGenerator"
Option Explicit

Sub Main()

    Dim rowIterator As Integer

    For rowIterator = 2 To 3
        If ReadCellValue(rowIterator, 1) = "yes" Then
            Call CreateLetter(ReadCellValue(rowIterator, 8), rowIterator)
        End If
    Next rowIterator

    MsgBox ("Done, your letter can be find in " & GetFilledTemplatesPath())

End Sub

Sub CreateLetter(fileName As String, rowNumber As Integer)  'dodaj jak error to zamkniecie pliku

   Dim objWord: Set objWord = CreateObject("Word.Application")
   objWord.Visible = False
   Dim objDoc: Set objDoc = objWord.Documents.Open(GetEmptyTemplatesPath() & fileName & ".docx")
   Dim ValuesToFill As New LetterData
   ValuesToFill.Constructor (ReadValues(rowNumber))
   
    Call FillTemplate(objDoc, ValuesToFill)
    Call SaveFile(objDoc, ValuesToFill.GetAddressee(), fileName)
    Call CloseFile(objWord)

End Sub

Function GetEmptyTemplatesPath() As String

    GetEmptyTemplatesPath = Application.ActiveWorkbook.Path & "\template\"

End Function

Sub FillTemplate(ByRef openFile As Variant, ByRef ValuesToFill As LetterData)
    
    With openFile.Content.Find
      .Execute FindText:="$cityAndDate$", ReplaceWith:=ValuesToFill.GetCityAndDate(), _
        Format:=True, Replace:=wdReplaceAll, Forward:=True
    End With
    With openFile.Content.Find
      .Execute FindText:="$addressee$", ReplaceWith:=ValuesToFill.GetAddressee(), _
        Format:=True, Replace:=wdReplaceAll, Forward:=True
    End With
    With openFile.Content.Find
      .Execute FindText:="$informationOn$", ReplaceWith:=ValuesToFill.GetInformationOnDate(), _
        Format:=True, Replace:=wdReplaceAll, Forward:=True
    End With
    With openFile.Content.Find
      .Execute FindText:="$authorEmail$", ReplaceWith:=ValuesToFill.GetAuthorMail(), _
        Format:=True, Replace:=wdReplaceAll, Forward:=True
    End With
    With openFile.Content.Find
      .Execute FindText:="$authorName$", ReplaceWith:=ValuesToFill.GetAuthorName(), _
        Format:=True, Replace:=wdReplaceAll, Forward:=True
    End With
    With openFile.Content.Find
      .Execute FindText:="$informationOnMinusYear$", ReplaceWith:=ValuesToFill.GetPreviousInformationOnYear(), _
        Format:=True, Replace:=wdReplaceAll, Forward:=True
    End With

End Sub

Function ReadValues(rowIterator As Integer) As String()

    Dim Values(5) As String
    Dim columnIterator As Integer

    For columnIterator = 0 To 5
        Values(columnIterator) = ReadCellValue(rowIterator, columnIterator + 2)
    Next columnIterator
    
    ReadValues = Values

End Function

Function ReadCellValue(row As Integer, column As Integer) As String

    ReadCellValue = ThisWorkbook.ActiveSheet.Cells(row, column).Value

End Function

Sub SaveFile(ByRef openFile As Variant, CompanyName As String, fileName As String)

    openFile.SaveAs2 (GetFilledTemplatesPath() & CompanyName & "_" & fileName & ".docx")

End Sub

Function GetFilledTemplatesPath() As String

    GetFilledTemplatesPath = Application.ActiveWorkbook.Path & "\filled\"

End Function


Sub CloseFile(ByRef fileToClose As Variant)

    fileToClose.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    fileToClose.Quit

End Sub




