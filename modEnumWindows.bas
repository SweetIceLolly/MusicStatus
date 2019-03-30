Attribute VB_Name = "modEnumWindows"
Option Explicit

'Purpose:   The callback procedure of EnumWindows() function
'Args:      hWnd: Current window handle
'           lParam: Additional information.
'Return:    A BOOL type value, TRUE means keep enumerate windows, FALSE means exit the enumeration
Public Function EnumProc(ByVal hWnd As Long, lParam As Long) As BOOL
    Dim WindowNameBuffer    As String * 255
    
    GetWindowTextW hWnd, WindowNameBuffer, 255
    If InStr(WindowNameBuffer, " - YouTube - ±ù¹÷µÄä¯ÀÀÆ÷") Then                            'YouTube music found
        frmMain.SongName = Split(WindowNameBuffer, " - YouTube - ±ù¹÷µÄä¯ÀÀÆ÷")(0) & vbNullChar
        EnumProc = 0
        Exit Function
    End If
    EnumProc = 1                                                                        'Search for next window
End Function
