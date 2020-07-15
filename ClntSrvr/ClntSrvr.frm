VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ClientServerWindow 
   ClientHeight    =   3252
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   4932
   ClipControls    =   0   'False
   Icon            =   "ClntSrvr.frx":0000
   ScaleHeight     =   13.55
   ScaleMode       =   4  'Character
   ScaleWidth      =   41.1
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
         Size            =   7.8
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
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   120
      Top             =   1080
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   120
      Top             =   120
      _ExtentX        =   593
      _ExtentY        =   593
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
Attribute VB_Name = "ClientServerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface window.
Option Explicit

'The actions that can be performed by the client, monitor, and server.
Private Enum ActionsE
   ClientClose = 1                'Closes the client side connection.
   ClientConnect                  'Connects the client to the server.
   ClientDisplayInformation       'Displays the client's connection information.
   ClientGetData                  'Gets the data received by the client.
   ClientSendData                 'Sends data from the client to the server.
   ClientSetRemoteHostPort        'Sets the server and port to connect with.
   MonitorClose                   'Closes the monitor.
   MonitorListen                  'Starts the monitor.
   MonitorSetLocalHostPort        'Sets the host and port to connect with by the client.
   ServerAccept                   'Accepts the connection requested by the client.
   ServerClose                    'Closes the server side connection.
   ServerDisplayInformation       'Displays the server's connection information.
   ServerGetData                  'Gets the data received by the server.
   ServerSendData                 'Sends data from the server to the client.
End Enum

'The modes in which this program can be executed.
Private Enum ModesE
   NoMode = 1                     'Indicates neither client nor server mode.
   ClientMode                     'Indicates client mode.
   ServerMode                     'Indicates server mode.
End Enum

'This procedure updates the specified output display with the specified text.
Private Sub Display(OutputBox As TextBox, ByVal NewText As String)
On Error Resume Next
Dim Position As Long
Dim Text As String

   If Not NewText = Empty Then
      NewText = Escape(NewText)
      
      With OutputBox
         For Position = 1 To Len(NewText) Step .MaxLength
            If Len(Mid$(NewText, Position)) < .MaxLength Then
               Text = Mid$(NewText, Position)
            Else
               Text = Mid$(NewText, Position, .MaxLength)
            End If
      
            If Len(.Text & NewText) > .MaxLength Then .Text = Mid$(.Text, Len(NewText))
                        
            .SelLength = 0
            .SelStart = Len(.Text)
            .SelText = .SelText & NewText
         Next Position
            
         If Not RemoteLineBreak() = Empty Then .Text = Escape(Replace(Unescape(.Text), Unescape(RemoteLineBreak()), vbCrLf))
      End With
   End If
End Sub

'This procedure displays the client's/monitor's/server's current states.
Private Sub DisplayState()
On Error Resume Next
Dim CurrentState  As String
Static PreviousState As String

   Select Case Mode()
      Case ClientMode
         CurrentState = "Client - Client: " & StateDescription(Client.State)
      Case ServerMode
         CurrentState = "Server - Monitor: " & StateDescription(Monitor.State) & " --- Server: " & StateDescription(Server.State)
   End Select
   
   If Not CurrentState = PreviousState Then
      With App
         Me.Caption = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName & " --- " & CurrentState
      End With
      
      PreviousState = CurrentState
   End If
End Sub

'This procedure gives the client/monitor/server the command to perform the specified action.
Private Sub DoAction(Action As ActionsE, Optional ByVal DataOut As String, Optional BytesReceived As Long = 0, Optional Request As Long = 0)
On Error GoTo ErrorTrap
Dim Data As String
Dim Description As String
Dim ErrorAt As Long
Dim ErrorCode As Long
Dim NewLocalHostPort As String
Dim NewRemoteHostPort As String
Dim Position As Long
   
   Select Case Action
      Case ClientGetData, ClientSendData
         If Not Client.State = sckConnected Then
            If Client.State = sckConnected Then Display ClientOutputBox, "Closing connection." & vbCrLf
      
            Client.Close
            Display ClientOutputBox, "Connecting." & vbCrLf
            Client.LocalPort = 0
            Client.Connect
      
            Do Until Client.State = sckConnected Or Client.State = sckError Or DoEvents() = 0
               DisplayState
            Loop
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
         Display ClientOutputBox, "Connecting." & vbCrLf
         Client.LocalPort = 0
         Client.Connect
      
         Do Until Client.State = sckConnected Or Client.State = sckError Or DoEvents() = 0
            DisplayState
         Loop
      Case ClientDisplayInformation
         Display ClientOutputBox, "Local: [" & Client.LocalIP & "] " & Client.LocalHostName & ":" & CStr(Client.LocalPort) & vbCrLf
         Display ClientOutputBox, "Remote: [" & Client.RemoteHostIP & "] " & Client.RemoteHost & ":" & CStr(Client.RemotePort) & vbCrLf
      Case ClientGetData
         Client.GetData Data, vbString, BytesReceived
         Display ClientOutputBox, Data
      Case ClientSendData
         DataOut = Unescape(DataOut, , ErrorAt) & Unescape(Suffix())
   
         If Not EscapeSequenceError(ErrorAt) Then
            If Echo() Then Display ClientOutputBox, DataOut
            Client.SendData DataOut
         End If
      Case ClientSetRemoteHostPort
         NewRemoteHostPort = InputBox$("Remote host and port. (host:port)", , Client.RemoteHost & ":" & CStr(Client.RemotePort))
      
         If Not NewRemoteHostPort = Empty Then
            Position = InStr(NewRemoteHostPort, ":")
            If Position > 0 Then
               Client.RemoteHost = Left$(NewRemoteHostPort, Position - 1)
               Client.RemotePort = Val(Mid$(NewRemoteHostPort, Position + 1))
            Else
               Client.RemoteHost = NewRemoteHostPort
            End If
         End If
      Case MonitorClose
         Display ServerOutputBox, "Closing the connection monitor." & vbCrLf
         Monitor.Close
      Case MonitorListen
         Display ServerOutputBox, "Listening at [" & Monitor.LocalIP & "] " & Monitor.LocalHostName & ":" & CStr(Monitor.LocalPort) & "." & vbCrLf
         Monitor.Listen
      Case MonitorSetLocalHostPort
         NewLocalHostPort = InputBox$("Local host and port. (host:port)", , Monitor.LocalHostName & ":" & CStr(Monitor.LocalPort))
         If Not NewLocalHostPort = Empty Then
            Position = InStr(NewLocalHostPort, ":")
            If Position > 0 Then
               Server.Bind Monitor.LocalPort, Monitor.LocalHostName
               Monitor.Bind Val(Mid$(NewLocalHostPort, Position + 1)), Left$(NewLocalHostPort, Position - 1)
            Else
               Monitor.Bind , NewLocalHostPort
               Server.Bind , Monitor.LocalHostName
            End If
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
         DataOut = Unescape(DataOut, , ErrorAt) & Unescape(Suffix())
   
         If Not EscapeSequenceError(ErrorAt) Then
            If Echo() Then Display ServerOutputBox, DataOut
            Server.SendData DataOut
         End If
   End Select
   
   DisplayState
   Exit Sub

ErrorTrap:
   Description = Err.Description
   ErrorCode = Err.Number
   Err.Clear
   On Error Resume Next
   Description = vbCrLf & "Error: " & CStr(ErrorCode) & " - " & Description & vbCrLf
   Select Case Mode()
      Case ClientMode
         Display ClientOutputBox, Description
      Case ServerMode
         Display ServerOutputBox, Description
   End Select
   Resume Next
End Sub


'This procedure manages the user input echo option.
Private Function Echo(Optional Toggle As Boolean = False) As Boolean
On Error Resume Next
   Static CurrentEcho As Boolean
   
   If Toggle Then
      CurrentEcho = Not CurrentEcho
      EchoInputMenu.Checked = CurrentEcho
   End If
   
   Echo = CurrentEcho
End Function

'This procedure converts non-displayable characters in the specified text to escape sequences.
Private Function Escape(Text As String, Optional EscapeCharacter As String = "/", Optional EscapeLineBreaks As Boolean = False) As String
On Error Resume Next
Dim Character As String
Dim Escaped As String
Dim Index As Long
Dim NextCharacter As String

   Escaped = Empty
   Index = 1
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = EscapeCharacter Then
         Escaped = Escaped & String$(2, EscapeCharacter)
      ElseIf Character = vbTab Or Character >= " " Then
         Escaped = Escaped & Character
      ElseIf Character & NextCharacter = vbCrLf And Not EscapeLineBreaks Then
         Escaped = Escaped & vbCrLf
         Index = Index + 1
      Else
         Escaped = Escaped & EscapeCharacter & String$(2 - Len(Hex$(Asc(Character))), "0") & Hex$(Asc(Character))
      End If
      Index = Index + 1
   Loop
   
   Escape = Escaped
End Function

'This procedure checks whether the return value for escape sequence procedures indicates an error.
Private Function EscapeSequenceError(ErrorAt As Long) As Boolean
On Error Resume Next
Dim EscapeError As Boolean
   
   EscapeError = (ErrorAt > 0)
   If EscapeError Then MsgBox "Bad escape sequence at character #" & CStr(ErrorAt) & ".", vbExclamation
   
   EscapeSequenceError = EscapeError
End Function


'This procedure returns/sets the mode in which this program is being executed.
Private Function Mode(Optional NewMode As ModesE = NoMode) As ModesE
On Error Resume Next
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
   
      DisplayState
   End If
   
   Mode = CurrentMode
End Function
'This procedure returns/sets the line break character(s) used by the remote client/server.
Private Function RemoteLineBreak(Optional NewRemoteLineBreak As String = Empty, Optional SetRemoteLineBreak As Boolean = False) As String
On Error Resume Next
Static CurrentRemoteLineBreak As String
   
   If SetRemoteLineBreak Then CurrentRemoteLineBreak = NewRemoteLineBreak
   
   RemoteLineBreak = CurrentRemoteLineBreak
End Function


'This procedure sends the user's input.
Private Sub SendData(Optional ByVal DataOut As String = Empty, Optional Repeat As Boolean = False)
On Error Resume Next
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
End Sub


'This procedure returns the description for the specified state.
Private Function StateDescription(State As Long) As String
On Error Resume Next
Dim States As Variant

   If States = Empty Then States = Array("Closed", "Open", "Listening", "Connection pending", "Resolving Host", "Host resolved", "Connecting", "Connected", "Peer is closing the connection", "Error")
   
   StateDescription = CStr(States(State))
End Function

'This procedure returns/sets a suffix for data that is sent by the user.
Private Function Suffix(Optional NewSuffix As String = Empty, Optional SetSuffix As Boolean = False) As String
On Error Resume Next
Static CurrentSuffix As String
   
   If SetSuffix Then CurrentSuffix = NewSuffix
   
   Suffix = CurrentSuffix
End Function

'This procedure converts any escape sequences in the specified text to characters.
Private Function Unescape(Text As String, Optional EscapeCharacter As String = "/", Optional ErrorAt As Long = 0) As String
On Error Resume Next
Dim Character As String
Dim Hexadecimals As String
Dim Index As Long
Dim NextCharacter As String
Dim Unescaped As String

   ErrorAt = 0
   Index = 1
   Unescaped = Empty
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = EscapeCharacter Then
         If NextCharacter = EscapeCharacter Then
            Unescaped = Unescaped & Character
            Index = Index + 1
         Else
            Hexadecimals = UCase$(Mid$(Text, Index + 1, 2))
            If Len(Hexadecimals) = 2 Then
               If Left$(Hexadecimals, 1) = "0" Then Hexadecimals = Right$(Hexadecimals, 1)
      
               If UCase$(Hex$(CLng(Val("&H" & Hexadecimals & "&")))) = Hexadecimals Then
                  Unescaped = Unescaped & Chr$(CLng(Val("&H" & Hexadecimals & "&")))
                  Index = Index + 2
               Else
                  ErrorAt = Index
                  Exit Do
               End If
            Else
               ErrorAt = Index
               Exit Do
            End If
         End If
      Else
         Unescaped = Unescaped & Character
      End If
      Index = Index + 1
   Loop
   
   Unescape = Unescaped
End Function

'This procedure clears the display.
Private Sub ClearMenu_Click()
On Error Resume Next

   Select Case Mode()
      Case ClientMode
         ClientOutputBox.Text = Empty
      Case ServerMode
         ServerOutputBox.Text = Empty
   End Select
End Sub

'This procedure gives the client the command to close any active connection.
Private Sub ClientCloseMenu_Click()
On Error Resume Next
   DoAction ClientClose
End Sub


'This procedure gives the client the command to connect with a server.
Private Sub ClientConnectMenu_Click()
On Error Resume Next
   DoAction ClientConnect
End Sub

'This procedure gives the command to activate the client mode.
Private Sub ClientModeMenu_Click()
On Error Resume Next
   Mode NewMode:=ClientMode
End Sub

'This procedure gives the client the command to request the user to specify the remote host and port.
Private Sub ClientRemoteHostAndPortMenu_Click()
On Error Resume Next
   DoAction ClientSetRemoteHostPort
End Sub

'This procedure informs the user when the client's connection has been closed.
Private Sub Client_Close()
On Error Resume Next
   Display ClientOutputBox, "Connection closed." & vbCrLf
   DisplayState
End Sub




'This procedure informs the user when the client has made a connection.
Private Sub Client_Connect()
On Error Resume Next
   Display ClientOutputBox, "Connected to [" & Client.RemoteHostIP & "] " & Client.RemoteHost & ":" & CStr(Client.RemotePort) & "." & vbCrLf
   DisplayState
End Sub

'This procedure gives the client the command to retrieve data that has been received.
Private Sub Client_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
   DoAction ClientGetData, , bytesTotal
End Sub

'This procedure informs the user when the client has encountered an error.
Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
   Display ClientOutputBox, vbCrLf & "Client error: " & CStr(Number) & " - " & Description & vbCrLf
   DisplayState
End Sub

'This procedure requests the user to specify a new data suffix.
Private Sub DataSuffixMenu_Click()
On Error Resume Next
Dim ErrorAt As Long
Dim NewSuffix As String

   NewSuffix = InputBox$("New suffix for data sent:", , Suffix())
   If StrPtr(NewSuffix) = 0 Then Exit Sub
   NewSuffix = Escape(Unescape(NewSuffix, , ErrorAt), , EscapeLineBreaks:=True)
   If EscapeSequenceError(ErrorAt) Then Exit Sub
   
   Suffix NewSuffix:=NewSuffix, SetSuffix:=True
End Sub

'This procedure gives the command to display client/server hosts/ips/ports.
Private Sub DisplayInformationMenu_Click()
On Error Resume Next
   DoAction ClientDisplayInformation
   DoAction ServerDisplayInformation
End Sub

'This procedure gives the command to toggle the echo user input option on/off.
Private Sub EchoInputMenu_Click()
On Error Resume Next
   Echo Toggle:=True
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error Resume Next
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   Me.Width = Screen.Width / 1.1
   Me.Height = Screen.Height / 1.1
   
   DataOutBox.ToolTipText = "Specify a character code using the foreward slash escape character (""/"") followed by a two digit hexadecimal value. Two escape characters are interpreted as a single foreward slash."
   If Not Echo() Then Echo Toggle:=True
   Mode NewMode:=ClientMode
   RemoteLineBreak NewRemoteLineBreak:=Empty, SetRemoteLineBreak:=True
   Suffix NewSuffix:="/0D/0A", SetSuffix:=True
   
   DoAction ClientDisplayInformation
   DoAction ServerDisplayInformation
   DisplayState
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
On Error Resume Next
   DoAction ClientClose
   DoAction MonitorClose
   DoAction ServerClose
   End
End Sub

'This procedure displays the information for this program.
Private Sub InformationMenu_Click()
On Error Resume Next
   MsgBox App.Comments, vbInformation
End Sub

'This procedure gives the monitor the command to stop listening for connection requests.
Private Sub MonitorCloseMenu_Click()
On Error Resume Next
   DoAction MonitorClose
End Sub


'This procedure gives the monitor the command to start listening for connection requests.
Private Sub MonitorListenMenu_Click()
On Error Resume Next
   DoAction MonitorListen
End Sub

'This procedure gives the monitor the command to request the user to specify the local host and port.
Private Sub MonitorLocalHostAndPortMenu_Click()
On Error Resume Next
   DoAction MonitorSetLocalHostPort
End Sub

'This procedure informs the user when a connection is requested and gives the command to accept the request.
Private Sub Monitor_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
   DoAction ServerAccept, , , requestID
End Sub


'This procedure informs the user when the monitor has encountered an error.
Private Sub Monitor_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
   Display ServerOutputBox, vbCrLf & "Monitor error: " & CStr(Number) & " - " & Description & vbCrLf
   DisplayState
End Sub


'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error Resume Next
   Unload Me
End Sub

'This procedure requests the user to specify a new remote line break.
Private Sub RemoteLineBreakMenu_Click()
On Error Resume Next
Dim ErrorAt As Long
Dim NewRemoteLineBreak As String

   NewRemoteLineBreak = InputBox$("New remote line break:", , RemoteLineBreak())
   If StrPtr(NewRemoteLineBreak) = 0 Then Exit Sub
   NewRemoteLineBreak = Escape(Unescape(NewRemoteLineBreak, , ErrorAt), , EscapeLineBreaks:=True)
   If EscapeSequenceError(ErrorAt) Then Exit Sub
   
   RemoteLineBreak NewRemoteLineBreak:=NewRemoteLineBreak, SetRemoteLineBreak:=True
End Sub

'This procedure repeats the previous user input.
Private Sub RepeatInputMenu_Click()
On Error Resume Next
   SendData , Repeat:=True
End Sub

'This procedure gives the command to send the user's input.
Private Sub SendButton_Click()
On Error Resume Next
   SendData DataOutBox.Text
   DataOutBox.Text = Empty
End Sub


'This procedure gives the server the command to close any active connections.
Private Sub ServerCloseMenu_Click()
On Error Resume Next
   DoAction ServerClose
End Sub


'This procedure gives the command to activate the client mode.
Private Sub ServerModeMenu_Click()
On Error Resume Next
   Mode NewMode:=ServerMode
End Sub


'This procedure informs the user when the server's connection has been closed.
Private Sub Server_Close()
On Error Resume Next
   Display ServerOutputBox, "Server: Connection closed." & vbCrLf
   DisplayState
End Sub

'This procedure gives the server the command to retrieve data that has been received.
Private Sub Server_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
   DoAction ServerGetData, , bytesTotal
End Sub

'This procedure informs the user when the server has encountered an error.
Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
   Display ServerOutputBox, vbCrLf & "Server error: " & CStr(Number) & " - " & Description & vbCrLf
   DisplayState
End Sub

