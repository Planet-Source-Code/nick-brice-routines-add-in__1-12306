VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routine Add In's"
   ClientHeight    =   3195
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comAddIn 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraRoutines 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5775
      Begin VB.CommandButton btnSelectRoutines 
         Caption         =   "&Deselect All"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton btnSelectRoutines 
         Caption         =   "&Select All"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ListBox lstRoutines 
         Height          =   2310
         Left            =   0
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label labLabels 
      AutoSize        =   -1  'True
      Caption         =   "Add In:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

' ================================================================================
' Routine     : btnSelectRoutines_Click
' Description : Select/Deselect all routine names in the list box control
' Parameters  : Index  - 0 = Select / 1 = Deselect
' Returns     : None
' ================================================================================
Private Sub btnSelectRoutines_Click(Index As Integer)
    Dim llCount As Long
    For llCount = 0 To lstRoutines.ListCount - 1
        lstRoutines.Selected(llCount) = Index - 1
    Next llCount
End Sub

' ================================================================================
' Routine     : CancelButton_Click
' Description : Cancel the AddIn
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub CancelButton_Click()
    Unload Me
    Connect.Hide
End Sub

' ================================================================================
' Routine     : comAddIn_Click
' Description : Initialise the selected AddIn type
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub comAddIn_Click()
    Select Case comAddIn.ListIndex
        Case 0  ' Add comments to routines"
            fraRoutines.Visible = True
            fraRoutines.ZOrder
            GetRoutineList
        Case 1  ' Add error handling to routines"
            fraRoutines.Visible = True
            fraRoutines.ZOrder
            GetRoutineList
        Case Else
    End Select
    OKButton.ZOrder
    CancelButton.ZOrder
End Sub

' ================================================================================
' Routine     : Form_Activate
' Description : Read the VB manual !!
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub Form_Activate()
    comAddIn.ListIndex = -1
    lstRoutines.Clear
End Sub

' ================================================================================
' Routine     : Form_Load
' Description : Read the VB manual !!
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub Form_Load()
    fraRoutines.Visible = False
    With comAddIn
        .Clear
        .AddItem "Add comments to routines"
        .AddItem "Add error handling to routines"
        .ListIndex = -1
    End With
    lstRoutines.Clear
End Sub

' ================================================================================
' Routine     : OKButton_Click
' Description : Process the selected AddIn and close
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub OKButton_Click()
    Select Case comAddIn.ListIndex
        Case 0  ' Add comments to routines
            AddCommentsToRoutines
        Case 1  ' Add error handling to routines
            AddErrorHandlingToRoutines
        Case Else
    End Select
    Unload Me
    Connect.Hide
End Sub

' ================================================================================
' Routine     : GetRoutineList
' Description : Get the list of routines in the current code module
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub GetRoutineList()
    Dim lsLine As String
    Dim lsElements() As String
    Dim lsRoutineName As String
    Dim llCount As Long
    Dim llTemp As Long
    lstRoutines.Clear
    With VBInstance
        On Error GoTo GetRoutineListError
        If .ActiveCodePane.Window.Type <> vbext_wt_CodeWindow Then Exit Sub
        On Error GoTo 0
        For llCount = .ActiveCodePane.CodeModule.CountOfDeclarationLines + 1 To .ActiveCodePane.CodeModule.CountOfLines
            lsLine = .ActiveCodePane.CodeModule.Lines(llCount, 1)
            lsElements = Split(lsLine, " ")
            If UBound(lsElements) > 1 Then
                Select Case LCase$(lsElements(0))
                    Case "public", "private"
                        lsRoutineName = Left$(lsElements(2), InStr(lsElements(2), "(") - 1)
                    Case "sub", "function"
                        lsRoutineName = Left$(lsElements(1), InStr(lsElements(1), "(") - 1)
                    Case Else
                        lsRoutineName = ""
                End Select
                If lsRoutineName > "" Then
                    Select Case Me.comAddIn.ListIndex
                        Case 0  ' Add comments to routines
                            lstRoutines.AddItem lsRoutineName
                        Case 1  ' Add error handling to routines
                        ' Check to make sure that error handling isn't already in routine
                            If InStr(LCase$(.ActiveCodePane.CodeModule.Lines(llCount, .ActiveCodePane.CodeModule.ProcCountLines(lsRoutineName, vbext_pk_Proc))), "on error goto ") = 0 Then lstRoutines.AddItem lsRoutineName
                        Case Else
                    End Select
                End If
            End If
        Next llCount
    End With
    Exit Sub
GetRoutineListError:
    MsgBox "This add in required the code window to be open !!", vbExclamation, "Warning"
End Sub

' ================================================================================
' Routine     : AddCommentsToRoutines
' Description : Code for the "Add Comments To Routines" AddIn
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub AddCommentsToRoutines()
    Dim lsRoutineList As String
    Dim lsRoutineName As String
    Dim lsLine As String
    Dim lsElements() As String
    Dim lsParameters() As String
    Dim llCount As Long
    Dim llLine As Long
    Dim lsReturns() As String
    Dim liMaxLen As Integer

    
    If lstRoutines.SelCount = 0 Then
        MsgBox "You have not selected any routines to comment !!", vbExclamation, "Warning"
        Exit Sub
    End If
    
    lsRoutineList = "|"
    For llLine = 0 To lstRoutines.ListCount - 1
        If lstRoutines.Selected(llLine) = True Then lsRoutineList = lsRoutineList & lstRoutines.List(llLine) & "|"
    Next llLine
    
    With VBInstance
        llLine = .ActiveCodePane.CodeModule.CountOfDeclarationLines + 1
        Do While llLine <= .ActiveCodePane.CodeModule.CountOfLines
            lsLine = .ActiveCodePane.CodeModule.Lines(llLine, 1)
            lsElements = Split(lsLine, " ")
            If UBound(lsElements) > 1 Then
                If (InStr(LCase$(lsLine), "sub") > 0 Or InStr(LCase$(lsLine), "function") > 0) And InStr(lsLine, "(") > 0 And InStr(lsLine, ")") > 0 Then
                ' Must be a routine declaration line
                    lsRoutineName = GetRoutineName(lsLine)
                    If InStr(lsRoutineList, "|" & lsRoutineName & "|") > 0 Then
                    ' Routine has been selected by user to be commented
                        lsParameters() = GetParameters(lsLine)
                        liMaxLen = GetMaxParameterLength(lsParameters)
                        lsReturns() = GetReturns(lsLine)
                    ' Output the comment lines
                      ' Separator line (========================================...)
                        .ActiveCodePane.CodeModule.InsertLines llLine, "' " & String$(80, "=")
                        llLine = llLine + 1
                      ' Routine name
                        .ActiveCodePane.CodeModule.InsertLines llLine, "' Routine     : " & lsRoutineName
                        llLine = llLine + 1
                      ' Routine description
                        .ActiveCodePane.CodeModule.InsertLines llLine, "' Description : "
                        llLine = llLine + 1
                      ' Parameter(s)
                        If UBound(lsParameters) >= 0 Then
                            .ActiveCodePane.CodeModule.InsertLines llLine, "' Parameters  : " & Left$(lsParameters(0) & Space$(liMaxLen), liMaxLen) & " - "
                            llLine = llLine + 1
                            For llCount = 1 To UBound(lsParameters)
                                .ActiveCodePane.CodeModule.InsertLines llLine, "'             : " & Left$(lsParameters(llCount) & Space$(liMaxLen), liMaxLen) & " - "
                                llLine = llLine + 1
                            Next llCount
                        Else
                            .ActiveCodePane.CodeModule.InsertLines llLine, "' Parameters  : None"
                            llLine = llLine + 1
                        End If
                      ' Return value(s)
                        If UBound(lsReturns) >= 0 Then
                            If liMaxLen = 0 Then liMaxLen = 5
                            If Trim$(lsReturns(0)) = "" Then
                                .ActiveCodePane.CodeModule.InsertLines llLine, "' Returns     : "
                            ElseIf Trim$(lsReturns(0)) = "None" Then
                                .ActiveCodePane.CodeModule.InsertLines llLine, "' Returns     : " & lsReturns(0)
                            Else
                                .ActiveCodePane.CodeModule.InsertLines llLine, "' Returns     : " & Left$(lsReturns(0) & Space$(liMaxLen), liMaxLen) & " - "
                            End If
                            llLine = llLine + 1
                            For llCount = 1 To UBound(lsReturns)
                                .ActiveCodePane.CodeModule.InsertLines llLine, "'             : " & Left$(lsReturns(llCount) & Space$(liMaxLen), liMaxLen) & " - "
                                llLine = llLine + 1
                            Next llCount
                        Else
                            .ActiveCodePane.CodeModule.InsertLines llLine, "' Returns     : None"
                            llLine = llLine + 1
                        End If
                      ' Separator line (========================================...)
                        .ActiveCodePane.CodeModule.InsertLines llLine, "' " & String$(80, "=")
                        llLine = llLine + 1
                    End If
                End If
            End If
        llLine = llLine + 1
        Loop
    End With
    Exit Sub
End Sub

' ================================================================================
' Routine     : AddErrorHandlingToRoutines
' Description : Code for the "Add Error Handling To Routines" AddIn
' Returns     : None
' Parameters  : None
' ================================================================================
Private Sub AddErrorHandlingToRoutines()
    Dim lsRoutineList As String
    Dim lsRoutineName As String
    Dim lsRoutineType As String
    Dim lsReturnDefault As String
    Dim lsLine As String
    Dim lsElements() As String
    Dim llCount As Long
    Dim llLine As Long

    
    If lstRoutines.SelCount = 0 Then
        MsgBox "You have not selected any routines to comment !!", vbExclamation, "Warning"
        Exit Sub
    End If
    
    lsRoutineList = "|"
    For llLine = 0 To lstRoutines.ListCount - 1
        If lstRoutines.Selected(llLine) = True Then lsRoutineList = lsRoutineList & lstRoutines.List(llLine) & "|"
    Next llLine
    
    With VBInstance
        llLine = .ActiveCodePane.CodeModule.CountOfDeclarationLines + 1
        Do While llLine <= .ActiveCodePane.CodeModule.CountOfLines
            lsLine = .ActiveCodePane.CodeModule.Lines(llLine, 1)
            lsElements = Split(lsLine, " ")
            If UBound(lsElements) > 1 Then
                If (InStr(LCase$(lsLine), "sub") > 0 Or InStr(LCase$(lsLine), "function") > 0) And InStr(lsLine, "(") > 0 And InStr(lsLine, ")") > 0 Then
                ' Must be a routine declaration line
                    lsRoutineName = GetRoutineName(lsLine)
                    lsRoutineType = GetRoutineType(lsLine)
                    lsReturnDefault = GetReturnDefault(lsLine)
                    If InStr(lsRoutineList, "|" & lsRoutineName & "|") > 0 Then
                    ' Routine has been selected by user
                    ' Output the new error handling lines
                      ' Increment the line so that the error handling starts on the first line of the routine
                        llLine = llLine + 1
                      ' Dimension a variable to place the error message into
                        .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & "Dim lsRhError As String"
                        llLine = llLine + 1
                      ' On Error Goto <Error Label>
                        .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & "On Error Goto " & lsRoutineName & "_Error"
                        llLine = .ActiveCodePane.CodeModule.ProcStartLine(lsRoutineName, vbext_pk_Proc) + .ActiveCodePane.CodeModule.ProcCountLines(lsRoutineName, vbext_pk_Proc) - 1
                      ' Make sure this line is in the current procedure
                        Do Until .ActiveCodePane.CodeModule.Lines(llLine, 1) Like "End*"
                            llLine = llLine - 1
                        Loop
                      ' On Error Goto 0
                        .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & "On Error Goto 0"
                        llLine = llLine + 1
                      ' Exit Sub/Function
                        .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & "Exit " & lsRoutineType
                        llLine = llLine + 1
                      ' <Error Label>:
                        .ActiveCodePane.CodeModule.InsertLines llLine, lsRoutineName & "_Error:"
                        llLine = llLine + 1
                      ' Set up the error message
                        .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & "lsRhError = " & Chr$(34) & "(" & Chr$(34) & " & Err.Number & " & Chr$(34) & ") " & Chr$(34) & " & Err.Description"
                        llLine = llLine + 1
'                      ' Output the error message to the log file
'                        .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & "RhLog gcoptions.sLogFile, " & Chr$(34) & lsRoutineName & "_Error:" & Chr$(34) & " & lsRhError"
'                        llLine = llLine + 1
                      ' Display the error message in a message box
                        .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & "MsgBox " & Chr$(34) & "lsRhError" & Chr$(34) & ", vbCritical, " & Chr$(34) & "Error encountered" & Chr$(34)
                        llLine = llLine + 1
                      ' If the routine is a function, then return the default
                        If LCase$(lsRoutineType) = "function" Then
                            .ActiveCodePane.CodeModule.InsertLines llLine, vbTab & lsRoutineName & " = " & lsReturnDefault
                            llLine = llLine + 1
                        End If
                    End If
                End If
            End If
        llLine = llLine + 1
        Loop
    End With
    Exit Sub
End Sub

' ================================================================================
' Routine     : GetRoutineName
' Description : Get the routine name from the routine declaration
' Parameters  : psLine  - The line of code containing the routine declaration
' Returns     : The routine name
' ================================================================================
Private Function GetRoutineName(psLine As String) As String
    Dim lsElements() As String
    Dim liCount As Integer
    lsElements = Split(Left$(psLine, InStr(psLine, "(") - 1), " ")
    GetRoutineName = lsElements(UBound(lsElements))
End Function

' ================================================================================
' Routine     : GetRoutineType
' Description : Get the routine type from the routine declaration
' Parameters  : psLine  - The line of code containing the routine declaration
' Returns     : The routine type : Sub/Function
' ================================================================================
Private Function GetRoutineType(psLine As String) As String
    Dim lsElements() As String
    Dim liCount As Integer
    lsElements = Split(Left$(psLine, InStr(psLine, "(") - 1), " ")
    GetRoutineType = lsElements(UBound(lsElements) - 1)
End Function


' ================================================================================
' Routine     : GetParameters
' Description : Get the list of parameters in the routine declaration
' Parameters  : psLine  - The line of code containing the routine declaration
' Returns     : An array containing the parameter names
' ================================================================================
Private Function GetParameters(psLine As String) As Variant
    Dim lsParameters() As String
    Dim lsElements() As String
    Dim liParameter As Integer
    Dim liElement As Integer
    
    lsParameters = Split(ExtractCode(psLine, "(", ")"), ",")
    For liParameter = 0 To UBound(lsParameters)
        lsParameters(liParameter) = Trim$(lsParameters(liParameter))
        If InStr(LCase$(lsParameters(liParameter)), "as") > 0 Then
            lsElements() = Split(lsParameters(liParameter), " ")
            For liElement = 0 To UBound(lsElements)
                Select Case LCase$(lsElements(liElement))
                    Case "as"
                        lsParameters(liParameter) = lsElements(liElement - 1)
                        Exit For
                    Case Else
                End Select
            Next liElement
        End If
    Next liParameter
    GetParameters = lsParameters
End Function

' ================================================================================
' Routine     : GetMaxParameterLength
' Description : Get the maximum length of the parameter names for the routine
'               declaration
' Parameters  : psParameters() - The array of parameter names
' Returns     : The maximum length
' ================================================================================
Public Function GetMaxParameterLength(psParameters() As String) As Integer
    Dim liCount As Integer
    Dim liMaxLen As Integer
    liMaxLen = 0
    For liCount = 0 To UBound(psParameters)
        If Len(psParameters(liCount)) > liMaxLen Then liMaxLen = Len(psParameters(liCount))
    Next liCount
    GetMaxParameterLength = liMaxLen
End Function

' ================================================================================
' Routine     : GetReturns
' Description : Get the return value names/types for the routine declaration
' Parameters  : psLine  - The line of code containing the routine declaration
' Returns     : An array containing the return names/types
' ================================================================================
Private Function GetReturns(psLine As String) As Variant
    Dim lsElements() As String
    Dim liCount As Integer
    If InStr(LCase$(psLine), "sub ") > 0 Then
        ReDim lsElements(0)
        lsElements(0) = "None"
    Else
        lsElements = Split(psLine, " ")
        If LCase$(lsElements(UBound(lsElements))) = "boolean" Then
            ReDim lsElements(1)
            lsElements(0) = "True"
            lsElements(1) = "False"
        Else
            ReDim lsElements(0)
            lsElements(0) = ""
        End If
    End If
    GetReturns = lsElements
End Function

' ================================================================================
' Routine     : GetReturnDefault
' Description : Get the return types for the routine declaration
' Parameters  : psLine  - The line of code containing the routine declaration
' Returns     : The VB return type
' ================================================================================
Private Function GetReturnDefault(psLine As String) As String
    Dim lsElements() As String
    Dim liCount As Integer
    lsElements = Split(psLine, " ")
    Select Case LCase$(lsElements(UBound(lsElements)))
        Case "boolean"
            GetReturnDefault = False
        Case "integer", "double", "long"
            GetReturnDefault = "0"
        Case Else
            GetReturnDefault = Chr$(34) & Chr$(34)
    End Select
End Function

