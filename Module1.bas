Attribute VB_Name = "Module1"

Public comstate As Integer
Public selectedPort As Integer
'Public row As Integer
Public mystr As String
'Public eofcnt As Integer
'Public chararray(20001) As Variant
Public charcnt As Integer
Public nd As Integer
'Public tokidx As Integer
'Public token(20) As String
Public cmdstr As String
'Public paramschk As Integer
'Public setparams As Boolean
'Public memflg As Boolean
'Public clearflg As Boolean
'Public defacomm As Integer
Public curlab As Integer
Public comm(3) As Boolean
Public ackflag As Integer
Public Runflag As Integer

Public SpanVal As Integer
Public ZeroVal As Integer

Public PVal As Integer
Public IVal As Integer
Public DVal As Integer

Public Teeth As Integer
Public Band As Integer
Public Dir As Integer
Public MaxAngle As Integer
Public Freq As Integer

Sub main()
    Form3.Show
    ackflag = 0
    Runflag = 0
    Form4.Timer1.Enabled = False
    
End Sub

Sub CommCheck()
selectedPort = 0
    On Error Resume Next
    For i = 1 To 8 '30
        Err = 0
        Form4.MSComm1.CommPort = i
        Form4.MSComm1.PortOpen = True
        Form4.MSComm1.PortOpen = False
        
        If Err = 0 Then
            Form3.optcomm(i - 1).Enabled = True
            selectedPort = i
        Else
            Form3.optcomm(i - 1).Enabled = False
        End If
    Next
    
    If selectedPort = 0 Then
       If MsgBox("Serial Port not Connected", vbOKCancel, "Error") = vbOK Then End
    End If
End Sub
