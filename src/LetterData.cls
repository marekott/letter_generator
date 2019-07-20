VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LetterData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private city As String
Private letterDate As String
Private addressee As String
Private informationOn As String
Private authorMail As String
Private authorName As String

Public Sub Constructor(ByRef data As Variant)

    city = data(0)
    letterDate = data(1)
    addressee = data(2)
    informationOn = data(3)
    authorMail = data(4)
    authorName = data(5)

End Sub

Public Function GetCity() As String

    GetCity = city

End Function

Public Function GetLetterDate() As String

    GetLetterDate = letterDate

End Function

Public Function GetCityAndDate() As String

    GetCityAndDate = city & ", " & letterDate

End Function

Public Function GetAddressee() As String

    GetAddressee = addressee

End Function

Public Function GetInformationOnDate() As String

    GetInformationOnDate = informationOn

End Function

Public Function GetAuthorName() As String

    GetAuthorName = authorName

End Function

Public Function GetAuthorMail() As String

    GetAuthorMail = authorMail

End Function

Public Function GetPreviousInformationOnYear() As String

    Dim convertedPreviousYear As Date: convertedPreviousYear = year(informationOn) - 1
    Dim convertedMonth As Date: convertedMonth = Month(informationOn)
    Dim convertedDay As Date: convertedDay = Day(informationOn)
    
    GetPreviousInformationOnYear = CStr(Format(DateSerial(convertedPreviousYear, convertedMonth, convertedDay), "yyyy-mm-dd"))

End Function
