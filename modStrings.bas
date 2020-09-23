Attribute VB_Name = "modStrings"
' ================================================================================
' Function    : ExtractCode
' Description : Function to return a sub-string between supplied characters within
'               supplied string
' e.g.        : RhExtractCode("abcd[efgh]ijkl","[","]") will return "efgh"
' Parameters  : lsString - The string to be validated
'               lsOpenCharacter  - The opening delimiter
'               lsCloseCharacter - The closing delimiter
' Returns     : The sub-string
' ================================================================================
Public Function ExtractCode(lsString, lsOpenCharacter, lsCloseCharacter) As String
    Dim liPos1 As Integer
    Dim liPos2 As Integer

    ExtractCode = ""

    liPos1 = InStr(lsString, lsOpenCharacter) + 1
    liPos2 = InStr(lsString, lsCloseCharacter) - 1

    Select Case liPos1
        Case Is = liPos2
            ExtractCode = Mid$(lsString, liPos1, 1)
        Case Is > liPos2
            Exit Function
        Case Else
            ExtractCode = Mid$(lsString, liPos1, liPos2 - liPos1 + 1)
    End Select
End Function

