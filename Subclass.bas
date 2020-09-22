Attribute VB_Name = "Subclass"
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Private lpPrevWndProc As Long
Public mysock As Long

Public Function Hook(ByVal hWnd As Long)
    'ok, we are going to catch ALL msg's sent
    'to the handle we are subclassing (form1)
    lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub UnHook(ByVal hWnd As Long)
    'if we dont un-subclass before we shutdown
    'the program, we get an illigal procedure error.
    'fun.
    Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim x As Long, a As String
Dim ReadBuffer(1000) As Byte
Debug.Print uMsg, wParam, lParam
    Select Case uMsg
        Case 1025:
            Select Case lParam
                Case FD_READ: 'lets check for data
                    x = recv(mysock, ReadBuffer(0), 1000, 0) 'try to get some
                    If x > 0 Then 'was there any?
                        a = StrConv(ReadBuffer, vbUnicode) 'yep, lets change it to stuff we can understand
                        Form1.txtStatus.Text = Form1.txtStatus.Text & a & vbCrLf 'add it to the text box on form1
                    End If

                Case FD_CONNECT: 'did we connect?
                    mysock = wParam 'yep, we did! yayay
                    Data = "NICK fred_bob_test" & Chr$(10) & Chr$(13) 'string to send
                    Form1.txtStatus.Text = Form1.txtStatus.Text & Data 'add to txtstatus on form1
                    Call SendData(mysock, Data) 'send the data
                    Data = "USER fred_bob_test asfd asfd asdf:asdf" & Chr$(10) & Chr$(13) 'string to send
                    Form1.txtStatus.Text = Form1.txtStatus.Text & Data 'add to txtstatus on form1
                    Call SendData(mysock, Data) 'send the data

                Case FD_CLOSE: 'uh oh. they closed the connection
                    Call closesocket(wp) 'so we need to close
            End Select
    End Select
    'let the msg get through to the form
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
