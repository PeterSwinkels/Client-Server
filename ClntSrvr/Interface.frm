VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form InterfaceWindow 
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4935
   ClipControls    =   0   'False
   Icon            =   "Interface.frx":0000
   ScaleHeight     =   13.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   41.125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ServerOutputBox 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox ClientOutputBox 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox DataOutBox 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton SendButton 
      Caption         =   "&Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Monitor 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu ClientModeMenu 
         Caption         =   "&Client Mode"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu ServerModeMenu 
         Caption         =   "&Server Mode"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu ProgramSeparator1Menu 
         Caption         =   "-"
      End
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^J
      End
      Begin VB.Menu ProgramSeparator2Menu 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu ClientMainMenu 
      Caption         =   "&Client"
      Begin VB.Menu ClientCloseMenu 
         Caption         =   "C&lose"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ClientConnectMenu 
         Caption         =   "C&onnect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu ClientRemoteHostAndPortMenu 
         Caption         =   "&Remote Host and Port"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu MonitorMainMenu 
      Caption         =   "&Monitor"
      Begin VB.Menu MonitorCloseMenu 
         Caption         =   "&Close"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MonitorListenMenu 
         Caption         =   "&Listen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MonitorLocalHostAndPortMenu 
         Caption         =   "&Local Host and Port"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu ServerMainMenu 
      Caption         =   "&Server"
      Begin VB.Menu ServerCloseMenu 
         Caption         =   "&Close"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu OptionsMainMenu 
      Caption         =   "&Options"
      Begin VB.Menu DataSuffixMenu 
         Caption         =   "&Data Suffix"
         Shortcut        =   ^S
      End
      Begin VB.Menu EchoInputMenu 
         Caption         =   "&Echo Input"
         Shortcut        =   ^E
      End
      Begin VB.Menu RemoteLineBreakMenu 
         Caption         =   "&Remote Line Break"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu OtherMainMenu 
      Caption         =   "O&ther"
      Begin VB.Menu ClearMenu 
         Caption         =   "&Clear"
         Shortcut        =   ^C
      End
      Begin VB.Menu DisplayInformationMenu 
         Caption         =   "Display &Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu RepeatInputMenu 
         Caption         =   "&Repeat Input"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface window.
Option Explicit



'This procedure gives the client/monitor/server the command to perform the specified action.
Private Sub DoAction(Action As ActionsE, Optional ByVal DataOut As String, Optional BytesReceived As Long = 0, Optional Request As Long = 0)
On Error GoTo ErrorTrap
Dim Data As String
Dim NewHost As String
Dim NewPort As String

   Select Case Action
      Case ClientGetData, ClientSendData
         If Not Client.State = sckConnected Then
            If Client.State = sckConnected Then Display ClientOutputBox, "Closing connection." & vbCrLf
      
            Client.Close
            DoClientConnect
         End If
      Case ClientConnect, ClientSetRemoteHostPort
         If Client.State = sckConnected Then Display ClientOutputBox, "Closing connection." & vbCrLf
         
         Client.Close
      Case MonitorListen, MonitorSetLocalHostPort
         If Monitor.State = sckListening Then Display ServerOutputBox, "Closing the connection monitor." & vbCrLf
         Monitor.Close
         Server.Close
      Case ServerAccept
         Display ServerOutputBox, "Connection id: " & CStr(Request) & vbCrLf
         If Server.State = sckConnected Then Display ServerOutputBox, "Closing connection." & vbCrLf
         Monitor.Close
         Server.Close
   End Select
    
   Select Case Action
      Case ClientClose
         Display ClientOutputBox, "Closing connection." & vbCrLf
         Client.Close
      Case ClientConnect
         DoClientConnect
      Case ClientDisplayInformation
         Display ClientOutputBox, "Local: [" & Client.LocalIP & "] " & Client.LocalHostName & ":" & CStr(Client.LocalPort) & vbCrLf
         Display ClientOutputBox, "Remote: [" & Client.RemoteHostIP & "] " & Client.RemoteHost & ":" & CStr(Client.RemotePort) & vbCrLf
      Case ClientGetData
         Client.GetData Data, vbString, BytesReceived
         Display ClientOutputBox, Data
      Case ClientSendData
         DoSendData DataOut, Client, ClientOutputBox
      Case ClientSetRemoteHostPort
         RequestHostAndPort "Remote host and port. (host:port)", NewHost, NewPort, Client.RemoteHost & ":" & CStr(Client.RemotePort)
      
         If Not (NewHost = vbNullString And NewPort = vbNullString) Then
            Client.RemoteHost = NewHost
            Client.RemotePort = Val(NewPort)
         ElseIf Not NewHost = vbNullString Then
            Client.RemoteHost = NewHost
         End If
      Case MonitorClose
         Display ServerOutputBox, "Closing the connection monitor." & vbCrLf
         Monitor.Close
      Case MonitorListen
         Display ServerOutputBox, "Listening at [" & Monitor.LocalIP & "] " & Monitor.LocalHostName & ":" & CStr(Monitor.LocalPort) & "." & vbCrLf
         Monitor.Listen
      Case MonitorSetLocalHostPort
         RequestHostAndPort "Local host and port. (host:port)", NewHost, NewPort, Monitor.LocalHostName & ":" & CStr(Monitor.LocalPort)
         
         If Not (NewHost = vbNullString And NewPort = vbNullString) Then
            Server.Bind Monitor.LocalPort, Monitor.LocalHostName
            Monitor.Bind Val(NewPort), NewHost
         ElseIf Not NewHost = vbNullString Then
            Monitor.Bind , NewHost
            Server.Bind , NewHost
         End If
      Case ServerAccept
         Server.Accept Request
         Display ServerOutputBox, "Connection from [" & Server.RemoteHostIP & "] " & Server.RemoteHost & ":" & CStr(Server.RemotePort) & " accepted." & vbCrLf
      Case ServerClose
         Display ServerOutputBox, "Closing connection." & vbCrLf
         Server.Close
      Case ServerDisplayInformation
         Display ServerOutputBox, "Local: [" & Server.LocalIP & "] " & Server.LocalHostName & ":" & CStr(Server.LocalPort) & vbCrLf
         Display ServerOutputBox, "Remote: [" & Server.RemoteHostIP & "] " & Server.RemoteHost & ":" & CStr(Server.RemotePort) & vbCrLf
         Display ServerOutputBox, "Monitor: [" & Monitor.LocalIP & "] " & Monitor.LocalHostName & ":" & CStr(Monitor.LocalPort) & vbCrLf
      Case ServerGetData
         Server.GetData Data, vbString, BytesReceived
         Display ServerOutputBox, Data
      Case ServerSendData
         DoSendData DataOut, Server, ServerOutputBox
   End Select
   
EndRoutine:
   Me.Caption = GetState()
   Exit Sub

ErrorTrap:
   HandleActionError
   Resume EndRoutine
End Sub

'This procedure attempts to connect the client to a server.
Private Sub DoClientConnect()
On Error GoTo ErrorTrap
   Display ClientOutputBox, "Connecting." & vbCrLf
   Client.LocalPort = 0
   Client.Connect

   Do Until Client.State = sckConnected Or Client.State = sckError Or DoEvents() = 0
      Me.Caption = GetState()
   Loop
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleActionError
   Resume EndRoutine
End Sub


Private Sub DoSendData(DataOut As String, Sender As Winsock, OutputBox As TextBox)
On Error GoTo ErrorTrap
Dim ErrorAt As Long

   DataOut = Unescape(DataOut, , ErrorAt) & Unescape(Suffix())

   If Not EscapeSequenceError(ErrorAt) Then
      If Echo() Then Display OutputBox, DataOut
      Sender.SendData DataOut
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleActionError
   Resume EndRoutine
End Sub

'This procedure manages and returns the user input echo option.
Private Function Echo(Optional Toggle As Boolean = False) As Boolean
On Error GoTo ErrorTrap
   Static CurrentEcho As Boolean
   
   If Toggle Then
      CurrentEcho = Not CurrentEcho
      EchoInputMenu.Checked = CurrentEcho
   End If
   
EndRoutine:
   Echo = CurrentEcho
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the current state for the client, server and monitor.
Private Function GetState()
On Error GoTo ErrorTrap
Dim CurrentState  As String
Static PreviousState As String

   Select Case Mode()
      Case ClientMode
         CurrentState = "Client - Client: " & StateDescription(Client.State)
      Case ServerMode
         CurrentState = "Server - Monitor: " & StateDescription(Monitor.State) & " --- Server: " & StateDescription(Server.State)
   End Select
   
   If Not CurrentState = PreviousState Then
      CurrentState = ProgramInformation() & " --- " & CurrentState
      PreviousState = CurrentState
   End If
EndRoutine:
   GetState = CurrentState
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure handles any client/monitor/server action errors that occur.
Private Sub HandleActionError()
Dim Description As String
Dim ErrorCode As Long
   
   Description = Err.Description
   ErrorCode = Err.Number
   
   On Error GoTo ErrorTrap
   Description = vbCrLf & "Error code: " & CStr(ErrorCode) & " - " & Description & vbCrLf
   Select Case Mode()
      Case ClientMode
         Display ClientOutputBox, Description
      Case ServerMode
         Display ServerOutputBox, Description
   End Select
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure manages and returns the mode in which this program is being executed.
Private Function Mode(Optional NewMode As ModesE = NoMode) As ModesE
On Error GoTo ErrorTrap
Static CurrentMode As ModesE

   If Not NewMode = NoMode Then
      CurrentMode = NewMode
   
      Select Case CurrentMode
         Case ClientMode
            ClientModeMenu.Checked = True
            ServerModeMenu.Checked = False
      
            ClientMainMenu.Visible = True
            MonitorMainMenu.Visible = False
            ServerMainMenu.Visible = False
      
            ClientOutputBox.Visible = True
            ServerOutputBox.Visible = False
         Case ServerMode
            ClientModeMenu.Checked = False
            ServerModeMenu.Checked = True
      
            ClientMainMenu.Visible = False
            MonitorMainMenu.Visible = True
            ServerMainMenu.Visible = True
      
            ClientOutputBox.Visible = False
            ServerOutputBox.Visible = True
      End Select
   
      Me.Caption = GetState()
   End If
   
EndRoutine:
   Mode = CurrentMode
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure sends the specified user input.
Private Sub SendData(Optional ByVal DataOut As String = vbNullString, Optional Repeat As Boolean = False)
On Error GoTo ErrorTrap
Static LastDataOut As String

   If Repeat Then
      DataOut = LastDataOut
   ElseIf Not Repeat Then
      LastDataOut = DataOut
   End If
   
   Select Case Mode()
      Case ClientMode
         DoAction ClientSendData, DataOut
      Case ServerMode
         DoAction ServerSendData, DataOut
   End Select
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure clears the display.
Private Sub ClearMenu_Click()
On Error GoTo ErrorTrap

   Select Case Mode()
      Case ClientMode
         ClientOutputBox.Text = vbNullString
      Case ServerMode
         ServerOutputBox.Text = vbNullString
   End Select
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the client the command to close any active connection.
Private Sub ClientCloseMenu_Click()
On Error GoTo ErrorTrap
   DoAction ClientClose
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the client the command to connect with a server.
Private Sub ClientConnectMenu_Click()
On Error GoTo ErrorTrap
   DoAction ClientConnect
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to activate the client mode.
Private Sub ClientModeMenu_Click()
On Error GoTo ErrorTrap
   Mode NewMode:=ClientMode
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the client the command to request the user to specify the remote host and port.
Private Sub ClientRemoteHostAndPortMenu_Click()
On Error GoTo ErrorTrap
   DoAction ClientSetRemoteHostPort
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure informs the user when the client's connection has been closed.
Private Sub Client_Close()
On Error GoTo ErrorTrap
   Display ClientOutputBox, "Connection closed." & vbCrLf
   Me.Caption = GetState()
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub




'This procedure informs the user when the client has made a connection.
Private Sub Client_Connect()
On Error GoTo ErrorTrap
   Display ClientOutputBox, "Connected to [" & Client.RemoteHostIP & "] " & Client.RemoteHost & ":" & CStr(Client.RemotePort) & "." & vbCrLf
   Me.Caption = GetState()
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the client the command to retrieve data that has been received.
Private Sub Client_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrorTrap
   DoAction ClientGetData, , bytesTotal
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure informs the user when the client has encountered an error.
Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo ErrorTrap
   Display ClientOutputBox, vbCrLf & "Client error: " & CStr(Number) & " - " & Description & vbCrLf
   Me.Caption = GetState()
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure requests the user to specify a new data suffix.
Private Sub DataSuffixMenu_Click()
On Error GoTo ErrorTrap
Dim ErrorAt As Long
Dim NewSuffix As String

   NewSuffix = InputBox$("New suffix for data sent:", , Suffix())
   If Not StrPtr(NewSuffix) = 0 Then
      NewSuffix = Escape(Unescape(NewSuffix, , ErrorAt), , EscapeLineBreaks:=True)
      If Not EscapeSequenceError(ErrorAt) Then
         Suffix NewSuffix:=NewSuffix, SetSuffix:=True
      End If
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to display client/server hosts/ips/ports.
Private Sub DisplayInformationMenu_Click()
On Error GoTo ErrorTrap
   DoAction ClientDisplayInformation
   DoAction ServerDisplayInformation
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to toggle the echo user input option on/off.
Private Sub EchoInputMenu_Click()
On Error GoTo ErrorTrap
   Echo Toggle:=True
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Width = Screen.Width / 1.1
   Me.Height = Screen.Height / 1.1
   
   DataOutBox.ToolTipText = "Specify a character code using the foreward slash escape character (""/"") followed by a two digit hexadecimal value. Two escape characters are interpreted as a single foreward slash."
   If Not Echo() Then Echo Toggle:=True
   Mode NewMode:=ClientMode
   RemoteLineBreak NewRemoteLineBreak:=vbNullString, SetRemoteLineBreak:=True
   Suffix NewSuffix:="/0D/0A", SetSuffix:=True
   
   DoAction ClientDisplayInformation
   DoAction ServerDisplayInformation
   Me.Caption = GetState()
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure adjusts this window's objects to the window's new size.
Private Sub Form_Resize()
On Error Resume Next
   ClientOutputBox.Height = Me.ScaleHeight - DataOutBox.Height
   ClientOutputBox.Left = 0
   ClientOutputBox.Top = 0
   ClientOutputBox.Width = Me.ScaleWidth
   
   DataOutBox.Left = 0
   DataOutBox.Width = Me.ScaleWidth - SendButton.Width - 2
   DataOutBox.Top = Me.ScaleHeight - DataOutBox.Height
   
   SendButton.Left = DataOutBox.Width + 1
   SendButton.Top = (Me.ScaleHeight - (DataOutBox.Height / 2)) - (SendButton.Height / 2)
   
   ServerOutputBox.Height = Me.ScaleHeight - DataOutBox.Height
   ServerOutputBox.Left = 0
   ServerOutputBox.Top = 0
   ServerOutputBox.Width = Me.ScaleWidth
End Sub

'This procedure gives the command to close any active connections when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   DoAction ClientClose
   DoAction MonitorClose
   DoAction ServerClose
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the information for this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   MsgBox App.Comments, vbInformation
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the monitor the command to stop listening for connection requests.
Private Sub MonitorCloseMenu_Click()
On Error GoTo ErrorTrap
   DoAction MonitorClose
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the monitor the command to start listening for connection requests.
Private Sub MonitorListenMenu_Click()
On Error GoTo ErrorTrap
   DoAction MonitorListen
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the monitor the command to request the user to specify the local host and port.
Private Sub MonitorLocalHostAndPortMenu_Click()
On Error GoTo ErrorTrap
   DoAction MonitorSetLocalHostPort
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure informs the user when a connection is requested and gives the command to accept the request.
Private Sub Monitor_ConnectionRequest(ByVal requestID As Long)
On Error GoTo ErrorTrap
   DoAction ServerAccept, , , requestID
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure informs the user when the monitor has encountered an error.
Private Sub Monitor_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo ErrorTrap
   Display ServerOutputBox, vbCrLf & "Monitor error: " & CStr(Number) & " - " & Description & vbCrLf
   Me.Caption = GetState()
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure requests the user to specify a new remote line break.
Private Sub RemoteLineBreakMenu_Click()
On Error GoTo ErrorTrap
Dim ErrorAt As Long
Dim NewRemoteLineBreak As String

   NewRemoteLineBreak = InputBox$("New remote line break:", , RemoteLineBreak())
   If Not StrPtr(NewRemoteLineBreak) = 0 Then
      NewRemoteLineBreak = Escape(Unescape(NewRemoteLineBreak, , ErrorAt), , EscapeLineBreaks:=True)
      If Not EscapeSequenceError(ErrorAt) Then
         RemoteLineBreak NewRemoteLineBreak:=NewRemoteLineBreak, SetRemoteLineBreak:=True
      End If
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure repeats the previous user input.
Private Sub RepeatInputMenu_Click()
On Error GoTo ErrorTrap
   SendData , Repeat:=True
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to send the user's input.
Private Sub SendButton_Click()
On Error GoTo ErrorTrap
   SendData DataOutBox.Text
   DataOutBox.Text = vbNullString
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the server the command to close any active connections.
Private Sub ServerCloseMenu_Click()
On Error GoTo ErrorTrap
   DoAction ServerClose
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to activate the client mode.
Private Sub ServerModeMenu_Click()
On Error GoTo ErrorTrap
   Mode NewMode:=ServerMode
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure informs the user when the server's connection has been closed.
Private Sub Server_Close()
On Error GoTo ErrorTrap
   Display ServerOutputBox, "Server: Connection closed." & vbCrLf
   Me.Caption = GetState()
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the server the command to retrieve data that has been received.
Private Sub Server_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrorTrap
   DoAction ServerGetData, , bytesTotal
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure informs the user when the server has encountered an error.
Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo ErrorTrap
   Display ServerOutputBox, vbCrLf & "Server error: " & CStr(Number) & " - " & Description & vbCrLf
   Me.Caption = GetState()
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

