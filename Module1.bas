Attribute VB_Name = "Module1"
' Simple Example which will
'Find & Replace a String within Any file(s)
'Binaries  included ...
'
'by Pbryan^(2K And 2)
'
'Option Compare Database
'sorry, I wrote this in VBA, and copied to VB
' so this will work in VBA as Well ....
'
'

Option Explicit
Public NumFound As Integer
Const ChunkSize = 128 ' Size of Chunks for Prompting
Public Sub MakeChanges(Fname$, IDString$, NString$)
    Dim Resp
    Dim PosString, WhereString
    Dim FileNumber, a$, b$, NewString$
    Dim AString As String * ChunkSize
    Dim IsChanged As Boolean
    Dim BlockIsChanged As Boolean
    Dim tempstring
    IsChanged = False
    BlockIsChanged = False
    On Error GoTo Problems
    FileNumber = FreeFile
    PosString = 1
    WhereString = 0
    AString = Space$(ChunkSize)
    'Make sure strings have same size

    If Len(IDString$) > Len(NString$) Then
        NewString$ = NString$ + Space$(Len(IDString$) - Len(NString$))
    Else
        NewString$ = Left$(NString$, Len(IDString$))
    End If

    Open Fname$ For Binary As FileNumber
    NumFound = 0

    If LOF(FileNumber) < ChunkSize Then
        a$ = Space$(LOF(FileNumber))
        Get #FileNumber, 1, a$
        WhereString = LocateInStr(1, a$, IDString$)
    Else
        a$ = Space$(ChunkSize)
        Get #FileNumber, 1, a$
        b$ = a$
        WhereString = LocateInStr(1, a$, IDString$)
    End If

    Do
        While WhereString <> 0
            tempstring = Left$(a$, WhereString - 1) & NewString$ & Mid$(a$, WhereString + Len(NewString$))
            a$ = tempstring
            IsChanged = True
            BlockIsChanged = True
            WhereString = LocateInStr(WhereString + 1, a$, IDString$)
        Wend

        If BlockIsChanged Then
            Resp = MsgBox(b$, vbOKCancel, "Replace '" & IDString$ & "' with '" & NString$ & "' Here?")
                If Resp = 1 Then
                    Put #FileNumber, PosString, a$
                    NumFound = NumFound + 1
                    BlockIsChanged = False
                Else
                    BlockIsChanged = False
                End If
        End If
        PosString = ChunkSize + PosString - Len(IDString$)
        ' If we're finished, exit the loop

        If EOF(FileNumber) Or PosString > LOF(FileNumber) Then
            Exit Do
        End If

        ' Get the next chunk to scan

        If PosString + ChunkSize > LOF(FileNumber) Then
            a$ = Space$(LOF(FileNumber) - PosString + 1)
            Get #FileNumber, PosString, a$
            WhereString = LocateInStr(1, a$, IDString$)
            b$ = a$
        Else
            a$ = Space$(ChunkSize)
            Get #FileNumber, PosString, a$
            WhereString = LocateInStr(1, a$, IDString$)
            b$ = a$
        End If

    Loop Until EOF(FileNumber) Or PosString > LOF(FileNumber)

   
    Close
        Exit Sub
Problems:
    Close
    MsgBox "Error in MakeChanges." & vbCrLf & Err.Description, _
    vbExclamation, "Modify SQL: Error #" & Err.Number
    
    End
End Sub

Private Function LocateInStr(StartPos As Integer, StrToSearch As String, _
    StrToFind As String) As Integer
    
        LocateInStr = InStr(StartPos, UCase(StrToSearch), UCase(StrToFind))
End Function




