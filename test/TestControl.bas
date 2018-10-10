Option Explicit

Dim g_cur As String

' [SYN-G-GEN OUTPUT START] indent(0) $/^E_/$
'  psggConverterLib.dll converted from TestControl.xlsx. 
Dim num As Integer

' [SYN-G-GEN OUTPUT END]


' 
Dim g_bYesNo As Boolean
Function br_YES(st As String)
    If g_cur = "" Then
        If g_bYesNo=True Then
            g_cur = st
        End If
    End If
End Function

Function br_NO(st As String)
    If g_cur = "" Then
        If g_bYesNo=False Then
            g_cur = st
        End If
    End If
End Function

'
Sub TestControl()
    Dim lp As Integer
    g_cur = "S_START"
    For lp = 1 To 10000
        If g_cur = "S_END" Then
            Exit For
        
        ' [SYN-G-GEN OUTPUT START] indent(12) $/^S_/$
'  psggConverterLib.dll converted from TestControl.xlsx. 
            ElseIf g_cur = "S_END" Then 'S_END : END
                g_cur = ""
                '
            ElseIf g_cur = "S_NUM_1" Then 'S_NUM_1 :
                g_cur = ""
                '
                MsgBox "ANS. 1"
                If g_cur = "" Then
                    g_cur = "S_END"
                EndIf
            ElseIf g_cur = "S_NUM_2" Then 'S_NUM_2 :
                g_cur = ""
                '
                MsgBox "ANS. 2"
                If g_cur = "" Then
                    g_cur = "S_END"
                EndIf
            ElseIf g_cur = "S_NUM_3" Then 'S_NUM_3 :
                g_cur = ""
                '
                MsgBox "ANS. UNKNOWN"
                If g_cur = "" Then
                    g_cur = "S_END"
                EndIf
            ElseIf g_cur = "S_SELECT" Then 'S_SELECT :
                g_cur = ""
                '
                num = 3
                If num = 1 Then
                    g_cur = "S_NUM_1"
                ElseIf num = 2 Then
                    g_cur = "S_NUM_2"
                Else
                    g_cur = "S_NUM_3"
                End If
            ElseIf g_cur = "S_START" Then 'S_START : START
                g_cur = ""
                '
                If g_cur = "" Then
                    g_cur = "S_SELECT"
                EndIf


        ' [SYN-G-GEN OUTPUT END]
        Else
            'nothing to do
        End If
    Next

End Sub

