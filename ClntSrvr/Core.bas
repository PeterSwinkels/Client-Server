Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'This enumeration lists actions that can be performed by the client, monitor, and server.
Public Enum ActionsE
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

'This enumeration lists the modes in which this program can be executed.
Public Enum ModesE
   NoMode = 1                     'Indicates neither client nor server mode.
   ClientMode                     'Indicates client mode.
   ServerMode                     'Indicates server mode.
End Enum

'This procedure updates the specified output display with the specified text.
Public Sub Display(OutputBox As TextBox, ByVal NewText As String)
On Error GoTo ErrorTrap
Dim Position As Long
Dim Text As String

   If Not NewText = vbNullString Then
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
            
         If Not RemoteLineBreak() = vbNullString Then .Text = Escape(Replace(Unescape(.Text), Unescape(RemoteLineBreak()), vbCrLf))
      End With
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure converts non-displayable characters in the specified text to escape sequences and returns the result.
Public Function Escape(Text As String, Optional EscapeCharacter As String = "/", Optional EscapeLineBreaks As Boolean = False) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Escaped As String
Dim Index As Long
Dim NextCharacter As String

   Escaped = vbNullString
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
   
EndRoutine:
   Escape = Escaped
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure checks whether the specified escape sequence procedure value indicates an error and return this value.
Public Function EscapeSequenceError(ErrorAt As Long) As Boolean
On Error GoTo ErrorTrap
Dim EscapeError As Boolean
   
   EscapeError = (ErrorAt > 0)
   If EscapeError Then MsgBox "Bad escape sequence at character #" & CStr(ErrorAt) & ".", vbExclamation
   
EndRoutine:
   EscapeSequenceError = EscapeError
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   
   On Error GoTo ErrorTrap
   If MsgBox("Error code: " & CStr(ErrorCode) & vbCr & Description, vbOKCancel Or vbExclamation) = vbCancel Then End
EndRoutine:
   Exit Sub

ErrorTrap:
   Resume Terminate
   
Terminate:
   End
End Sub

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   InterfaceWindow.Show
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns this program's information.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String
   
   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName & ", ***2022*** "
   End With
   
EndRoutine:
   ProgramInformation = Information
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages and returns the line break character(s) used by the remote client/server.
Public Function RemoteLineBreak(Optional NewRemoteLineBreak As String = vbNullString, Optional SetRemoteLineBreak As Boolean = False) As String
On Error GoTo ErrorTrap
Static CurrentRemoteLineBreak As String
   
   If SetRemoteLineBreak Then CurrentRemoteLineBreak = NewRemoteLineBreak
   
EndRoutine:
   RemoteLineBreak = CurrentRemoteLineBreak
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure requests the user to specify a host and port.
Public Sub RequestHostAndPort(Prompt As String, ByRef NewHost As String, ByRef NewPort As String, DefaultHostAndPort As String)
On Error GoTo ErrorTrap
Dim NewHostAndPort As String
Dim Position As Long

   NewHostAndPort = InputBox$(Prompt, , DefaultHostAndPort)
   Position = InStr(NewHostAndPort, ":")
   If Position > 0 Then
      NewHost = Left$(NewHostAndPort, Position - 1)
      NewPort = Mid$(NewHostAndPort, Position + 1)
   Else
      NewHost = NewHostAndPort
      NewPort = vbNullString
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the description for the specified state.
Public Function StateDescription(State As Long) As String
On Error GoTo ErrorTrap
Dim States As Variant

   If States = vbNullString Then States = Array("Closed", "Open", "Listening", "Connection pending", "Resolving Host", "Host resolved", "Connecting", "Connected", "Peer is closing the connection", "Error")
   
EndRoutine:
   StateDescription = CStr(States(State))
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages and returns the suffix for data sent by the user.
Public Function Suffix(Optional NewSuffix As String = vbNullString, Optional SetSuffix As Boolean = False) As String
On Error GoTo ErrorTrap
Static CurrentSuffix As String
   
   If SetSuffix Then CurrentSuffix = NewSuffix
   
EndRoutine:
   Suffix = CurrentSuffix
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure converts any escape sequences in the specified text to characters and returns the result.
Public Function Unescape(Text As String, Optional EscapeCharacter As String = "/", Optional ErrorAt As Long = 0) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Hexadecimals As String
Dim Index As Long
Dim NextCharacter As String
Dim Unescaped As String

   ErrorAt = 0
   Index = 1
   Unescaped = vbNullString
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
   
EndRoutine:
   Unescape = Unescaped
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

