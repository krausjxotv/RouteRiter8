Attribute VB_Name = "Module5"
Option Explicit

Public Const WM_QUERYENDSESSION = &H11
Public Const WM_ENDSESSION = &H16

Public lpPrevWndProc As Long
Public gHW As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = -4

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg

        Case WM_QUERYENDSESSION
            ' MUST return non-zero or Windows thinks you are blocking shutdown
            WindowProc = 1
            Exit Function

        Case WM_ENDSESSION
            If wParam <> 0 Then
                ' Shutdown is happening — save settings, close files, etc.
                Call frmUtils.SaveSettingsBeforeExit
                Unload frmUtils
            End If
            WindowProc = 0
            Exit Function

    End Select

    WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)

End Function

