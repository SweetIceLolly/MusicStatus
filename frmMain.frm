VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "Tray.ocx"
Begin VB.Form frmMain 
   Caption         =   "MusicStatus"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheckFreeSocket 
      Interval        =   5000
      Left            =   6000
      Top             =   3720
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox chkAutoScroll 
      Caption         =   "Auto Scroll"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin SysTrayCtl.cSysTray Tray 
      Left            =   6600
      Top             =   3600
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmMain.frx":0000
      TrayTip         =   "MusicStatus"
   End
   Begin MSWinsockLib.Winsock wsMain 
      Index           =   0
      Left            =   7200
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox edLog 
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FreeSocket()        As Boolean                  'Free socket index
Public SongName         As String                   'Current song name returned by EnumProc()
Public PublicIP         As String                   'IP of icelolly.ddns.net

'Purpose:   To check if there is a socket listening for connection. If no, start one
Private Sub CheckListeningSocket()
    On Error Resume Next
    Dim i       As Integer
    
    For i = 0 To Me.wsMain.UBound                   'Check all socket status, exit procedure if there is a socket listening
        If Me.wsMain(i).state = sckListening Then
            If Err.Number = 0 Then
                Exit Sub
            Else
                Err.Clear
            End If
        End If
    Next i
    
    For i = 0 To UBound(FreeSocket)                 'Check if there are any free sockets
        If FreeSocket(i) = True Then
            Me.wsMain(i).Close                          'Start the free socket
            Me.wsMain(i).Bind 80
            Me.wsMain(i).Listen
            FreeSocket(i) = False                       'Mark the socket as unfree
            Exit Sub
        End If
    Next i
    
    i = Me.wsMain.UBound + 1                        'Index of the new socket
    Load Me.wsMain(i)                               'If there aren't any free socket, create a new socket
    ReDim Preserve FreeSocket(i)                    'Change capacity of free socket index list
    Me.wsMain(i).Close                              'Start the new socket
    Me.wsMain(i).Bind 80
    Me.wsMain(i).Listen
End Sub

'Purpose:   To add the specified string into the Log textbox
'Args:      strLog: The log string to be added
Private Sub AddLog(strLog As String)
    Me.edLog.Text = Me.edLog.Text & Time & " " & strLog & vbCrLf
    If Me.chkAutoScroll.Value = 1 Then
        Me.edLog.SelStart = Len(Me.edLog.Text)
    End If
End Sub

Private Sub cmdClear_Click()
    Me.edLog.Text = ""
End Sub

Private Sub cmdHide_Click()
    Me.Hide
    Me.Tray.InTray = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    'Check command line
    If LCase(Command) = "/hide" Then
        Me.Hide
    End If
    
    'Init. socket status array
    ReDim FreeSocket(0)
    
    'Start server
    Me.wsMain(0).Bind 80
    Me.wsMain(0).Listen
    
    'Retrieve IP of icelolly.ddns.net
    Load Me.wsMain(1)
    Me.wsMain(1).LocalPort = 0
    Me.wsMain(1).Connect "icelolly.ddns.net", 80
    Do
        DoEvents
        Sleep 10
    Loop Until Me.wsMain(1).RemoteHostIP <> ""
    PublicIP = Me.wsMain(1).RemoteHostIP
    AddLog "Resolved icelolly.ddns.net: IP=" & PublicIP
    Me.wsMain(1).Close
    Unload Me.wsMain(1)
    
    Me.wsMain(0).Close
    Me.wsMain(0).Bind 80
    Me.wsMain(0).Listen
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub tmrCheckFreeSocket_Timer()
    CheckListeningSocket
End Sub

Private Sub Tray_MouseDown(Button As Integer, Id As Long)
    If Button = vbLeftButton Then
        Me.Show
    ElseIf Button = vbRightButton Then
        PopupMenu Me.mnuPopup
    End If
End Sub

Private Sub wsMain_Close(Index As Integer)
    'When connection closes, mark as free
    Me.wsMain(Index).Close
    FreeSocket(Index) = True
    Call CheckListeningSocket
End Sub

Private Sub wsMain_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    
    'Refuse connection from the same IP
    Dim i   As Integer
    For i = 0 To Me.wsMain.UBound
        If i <> Index And Me.wsMain(i).RemoteHostIP = Me.wsMain(Index).RemoteHostIP And Me.wsMain(Index).state = sckConnected Then
            If Err.Number = 0 Then
                Me.wsMain(Index).Close
                Me.wsMain(Index).Bind 80
                Me.wsMain(Index).Listen
                AddLog "Rejected connection from " & Me.wsMain(Index).RemoteHostIP
                Exit Sub
            Else
                Err.Clear
            End If
        End If
    Next i
    
    Me.wsMain(Index).Close
    Me.wsMain(Index).Accept requestID
    AddLog Me.wsMain(Index).RemoteHostIP & " connected"
    FreeSocket(Index) = False
    
    Call CheckListeningSocket
End Sub

Private Sub wsMain_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim rtnString       As String
    Dim rtnData()       As Byte
    
    SongName = ""
    EnumWindows AddressOf EnumProc, 0
    
    rtnString = "Your IP: " & Me.wsMain(Index).RemoteHostIP & "<br>"
    rtnString = rtnString & "IP of icelolly.ddns.net is " & PublicIP & "<br>"
    
    If FindWindowW("#32770", "SoundWire Server") <> 0 Then
        rtnString = rtnString & "Music streaming server status: Started" & "<br>"
    Else
        rtnString = rtnString & "Music streaming server status: Not Started" & "<br>"
    End If
    
    If SongName <> "" Then
        rtnString = rtnString & "Current playing: " & SongName
    Else
        rtnString = rtnString & "No music is playing!"
    End If
    
    rtnData = StrConv(rtnString, vbFromUnicode)
    Me.wsMain(Index).SendData "HTTP/1.1 200 OK" & vbCrLf & _
                              "Date: Sun, 1, Jan 1950 00:00:00 GMT" & vbCrLf & _
                              "Content-Type: text/html" & vbCrLf & _
                              "Content-length: " & UBound(rtnData) & vbCrLf & vbCrLf & rtnString
End Sub

Private Sub wsMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'When socket error occured, mark as free
    Me.wsMain(Index).Close
    FreeSocket(Index) = True
    Call CheckListeningSocket
End Sub

Private Sub wsMain_SendComplete(Index As Integer)
    On Error Resume Next
    
    'When data is sent, mark as free
    AddLog Me.wsMain(Index).RemoteHostIP & " echo completed"
    Me.wsMain(Index).Close
    FreeSocket(Index) = True
    Call CheckListeningSocket
End Sub
