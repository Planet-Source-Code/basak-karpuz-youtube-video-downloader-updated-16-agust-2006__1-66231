Attribute VB_Name = "rSplitModule"
Option Explicit
'This module is completely coded by Ramci
'my email is ramci_geliyo@hotmail.com

Public Sub ShowSample()

    Dim Script$(), Cntr%

    Call rSplit("|1|2|3|4|5|", "|", Script)
    Debug.Print Chr(95)
    For Cntr = 0 To UBound(Script) - 1
        Debug.Print Script(Cntr)
    Next Cntr
    Debug.Print Chr(175)

End Sub

Public Sub rSplit(MainString$, SeperatorString$, OutputArray$(), Optional CapsSensetive As Boolean)

    Dim UboundOutputArray%, Counter%

    If Not CapsSensetive Then
        MainString = LCase(MainString)
        SeperatorString = LCase(SeperatorString)
    End If
    If Left(MainString, Len(SeperatorString)) <> SeperatorString Then MainString = SeperatorString + MainString
    If Right(MainString, Len(SeperatorString)) <> SeperatorString Then MainString = MainString + SeperatorString
    UboundOutputArray = rRepetition(MainString, SeperatorString) - 1
    If UboundOutputArray < 1 Then GoTo ErrorOccured
    ReDim Preserve OutputArray(UboundOutputArray - 1)
    For Counter = 0 To UboundOutputArray - 1
        OutputArray(Counter) = rGetString(MainString, SeperatorString, SeperatorString, Counter + 1)
    Next Counter
    Exit Sub

ErrorOccured:
    Debug.Print "Error Occured in rSplit Sub"
    Debug.Print "MainString = " + Chr(34) + MainString + Chr(34)
    Debug.Print "SeperatorString = " + Chr(34) + SeperatorString + Chr(34)
    Debug.Print "UboundOutputArray = "; UboundOutputArray
    Debug.Print "Counter = "; Counter

End Sub

Public Function rJoin$(StringArray$(), SeperatorString$)

    Dim Counter%

    For Counter = 0 To UBound(StringArray) - 1
        rJoin = rJoin + SeperatorString + StringArray(Counter)
    Next Counter
    rJoin = Mid(rJoin, Len(SeperatorString))

End Function

Public Function rRepetition%(MainString$, SubString$, Optional CapsSensetive As Boolean)

    Dim SubStart&

    If Not CapsSensetive Then
        MainString = LCase(MainString)
        SubString = LCase(SubString)
    End If
    SubStart = InStr(1, MainString, SubString) + Len(SubString)
    If SubStart = Len(SubString) Then GoTo ErrorOccured
    While SubStart > Len(SubString)
        SubStart = InStr(SubStart, MainString, SubString) + Len(SubString)
        rRepetition = rRepetition + 1
    Wend
    Exit Function

ErrorOccured:
    Debug.Print "Error Occured in rRepetition Function"
    Debug.Print "MainString = " + Chr(34) + MainString + Chr(34)
    Debug.Print "SubString = " + Chr(34) + SubString + Chr(34)
    Debug.Print "SubStart = "; SubStart
    Debug.Print "rRepetition = "; rRepetition

End Function

Public Function rGetString$(MainString$, StringBefore$, StringAfter$, Optional StringLoop% = 1, Optional CapsSensetive As Boolean)

    Dim StringStart&, StringEnd&, StringLenght&, StringLooper%

    If Not CapsSensetive Then
        MainString = LCase(MainString)
        StringBefore = LCase(StringBefore)
        StringAfter = LCase(StringAfter)
    End If
    If StringLoop < 1 Then GoTo ErrorOccured
    StringStart = InStr(1, MainString, StringBefore) + Len(StringBefore)
    If StringStart = Len(StringBefore) Then GoTo ErrorOccured
    If StringLoop > 1 Then
        For StringLooper = 1 To StringLoop - 1
            StringStart = InStr(StringStart, MainString, StringBefore) + Len(StringBefore)
            If StringStart = Len(StringBefore) Then GoTo ErrorOccured
        Next StringLooper
    End If
    StringEnd = InStr(StringStart, MainString, StringAfter)
    StringLenght = StringEnd - StringStart
    If StringLenght < 1 Then GoTo ErrorOccured
    rGetString = Mid(MainString, StringStart, StringLenght)
    Exit Function

ErrorOccured:
    Debug.Print "Error Occured in rGetString Function"
    Debug.Print "MainString = " + Chr(34) + MainString + Chr(34)
    Debug.Print "StringBefore = " + Chr(34) + StringBefore + Chr(34)
    Debug.Print "StringAfter = " + Chr(34) + StringAfter + Chr(34)
    Debug.Print "StringLoop = "; StringLoop
    Debug.Print "StringStart = "; StringStart
    Debug.Print "StringEnd = "; StringEnd
    Debug.Print "StringLenght = "; StringLenght
    Debug.Print "StringLooper = "; StringLooper
    rGetString = ""

End Function

