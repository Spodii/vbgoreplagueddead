VERSION 5.00
Object = "{00C99381-8913-471F-9EED-4A517B2EB0F9}#1.0#0"; "GOREsockServer.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Account Server"
   ClientHeight    =   1005
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ShutTmr 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer KeepAlive 
      Interval        =   60000
      Left            =   720
      Top             =   120
   End
   Begin GOREsock.GOREsockServer GOREsock 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnu 
      Caption         =   "menu"
      Begin VB.Menu mnushutdown 
         Caption         =   "&Shut Down"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ConBuf As DataBuffer
Private LocalSocketID As Long

Private Type ConnectionList
    LastPacket As Long
End Type
Private LastConnection As Long
Private ConnectionList() As ConnectionList

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Form_Load()
Dim s() As String
    
    'Add to the tray
    TrayAdd Me, "Plagued Dead Account Server", MouseMove
    Me.Hide
    Me.Visible = False
    Me.Refresh
    DoEvents

    'Load the file paths
    InitFilePaths
    
    'Create the conversion buffer
    Set ConBuf = New DataBuffer
    
    'Generate the encryption keys
    GenerateEncryptionKeys s()
    GOREsock.SetEncryption PacketEncTypeServerIn, PacketEncTypeServerOut, s()

    'Create the listen socket
    GOREsock.ClearPicture
    LocalSocketID = GOREsock.Listen("127.0.0.1", 16077)
    If GOREsock.Address(LocalSocketID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbNewLine & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbNewLine & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly
    
    'Connect to the database
    MySQL_Init

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Show the pop-up menu
    If X = 7740 Then
        Me.PopupMenu mnu, 0
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If ShutTmr.Enabled = False Then Cancel = 1

End Sub

Private Sub GOREsock_OnConnection(inSox As Long)

    'Resize the array to fit the new connection
    If inSox > LastConnection Then
        LastConnection = LastConnection + 10
        ReDim Preserve ConnectionList(1 To LastConnection)
    End If
    
    'Clear the last packet time
    ConnectionList(inSox).LastPacket = timeGetTime

End Sub

Private Sub GOREsock_OnDataArrival(inSox As Long, inData() As Byte)
Dim PacketID As Byte
Dim rBuf As DataBuffer
Dim BufUBound As Long

    'Check the last packet time
    If ConnectionList(inSox).LastPacket + 5000 < timeGetTime Then Exit Sub
    ConnectionList(inSox).LastPacket = timeGetTime
    
    'Create the databuffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    
    'Store the buffer ubound
    BufUBound = UBound(inData)

    Do
    
        'Get the packet ID
        PacketID = rBuf.Get_Byte
    
        'Go through the packet IDs
        Select Case PacketID
            Case 0
                rBuf.Overflow
            
            Case 1: Data_GetAccountData rBuf, inSox
            
        End Select
        
        'Exit when the buffer runs out
        If rBuf.Get_ReadPos > BufUBound Then Exit Do
        
    Loop
    
    Set rBuf = Nothing

End Sub

Private Sub Data_GetAccountData(ByRef rBuf As DataBuffer, ByVal inSox As Long)
Dim AccountName As String
Dim Password As String
    
    'Get the information
    AccountName = rBuf.Get_String
    Password = rBuf.Get_String
    AccountName = Trim$(AccountName)
    Password = Trim$(Password)
    
    'Check for valid strings
    If AccountName = vbNullString Then Exit Sub
    If Password = vbNullString Then Exit Sub
    If Len(AccountName) > 10 Then Exit Sub
    If Len(AccountName) < 3 Then Exit Sub
    If InStr(1, AccountName, ";") Then Exit Sub
    If InStr(1, Password, ";") Then Exit Sub
    
    'Request the information from the database
    DB_RS.Open "SELECT password,user1,user2,user3,user4,user5 FROM accounts WHERE `name`='" & AccountName & "'", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Check if the account exists
    If DB_RS.EOF Then
        ConBuf.PreAllocate 1
        ConBuf.Put_Byte 1
        GOREsock.SendData inSox, ConBuf.Get_Buffer()
        DB_RS.Close
        Exit Sub
    End If
    
    'MD5 the password
    Password = MD5_String(Password)
    
    'Check if the password is correct
    If DB_RS!Password <> Password Then
        ConBuf.PreAllocate 1
        ConBuf.Put_Byte 1
        GOREsock.SendData inSox, ConBuf.Get_Buffer()
        DB_RS.Close
        Exit Sub
    End If
    
    'Send the list of characters
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte 3
    ConBuf.Put_String DB_RS!user1
    ConBuf.Put_String DB_RS!user2
    ConBuf.Put_String DB_RS!user3
    ConBuf.Put_String DB_RS!user4
    ConBuf.Put_String DB_RS!user5
    GOREsock.SendData inSox, ConBuf.Get_Buffer()
    
    'Close the recordset
    DB_RS.Close

End Sub

Private Sub KeepAlive_Timer()
Dim i As Long

    'Send the KeepAlive query to the database
    DB_RS.Open "SELECT * FROM mail_lastid WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    
    'Remove the sockets that have been open too long (>10 seconds)
    For i = 1 To LastConnection
        If ConnectionList(i).LastPacket <> 0 Then
            If timeGetTime - ConnectionList(i).LastPacket > 10000 Then
                GOREsock.Shut i
            End If
        End If
    Next i

End Sub

Private Sub mnushutdown_Click()

    ShutTmr.Enabled = True

End Sub

Private Sub ShutTmr_Timer()
Static FailedUnloads As Long

    If FailedUnloads > 5 Or frmMain.GOREsock.ShutDown <> soxERROR Then
        GOREsock.UnHook
        TrayDelete
        On Error Resume Next
        DB_Conn.Close
        On Error GoTo 0
        Set ConBuf = Nothing
        Unload Me

    Else
        
        'If the socket is making an error on the shutdown sequence for more than a second, just unload anyways
        FailedUnloads = FailedUnloads + 1
        
    End If
    
End Sub
