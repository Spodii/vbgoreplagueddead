VERSION 5.00
Object = "{EA7042E9-9A74-4968-8251-4C5826CAF760}#1.0#0"; "GOREsockClient.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Client"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin GOREsock.GOREsockClient GOREsock 
      Left            =   600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer ShutdownTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements DirectXEvent8

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
Dim DevData(1 To 50) As DIDEVICEOBJECTDATA
Dim NumEvents As Long
Dim LoopC As Long
Dim Moved As Byte
Dim OldMousePos As POINTAPI

    On Error GoTo ErrOut

    'Check if message is for us
    If EventID <> MouseEvent Then Exit Sub
    If GetActiveWindow = 0 Then Exit Sub

    'Retrieve data
    NumEvents = DIDevice.GetDeviceData(DevData, DIGDD_DEFAULT)

    'Loop through data
    For LoopC = 1 To NumEvents
        Select Case DevData(LoopC).lOfs

        'Move on X axis
        Case DIMOFS_X
            If Windowed Then
                OldMousePos = MousePos
                GetCursorPos MousePos
                MousePos.X = MousePos.X - (Me.Left \ Screen.TwipsPerPixelX)
                MousePos.Y = MousePos.Y - (Me.Top \ Screen.TwipsPerPixelY)
                MousePosAdd.X = -(OldMousePos.X - MousePos.X)
                MousePosAdd.Y = -(OldMousePos.Y - MousePos.Y)
            Else
                MousePosAdd.X = (DevData(LoopC).lData * MouseSpeed)
                MousePos.X = MousePos.X + MousePosAdd.X
                If MousePos.X < 0 Then MousePos.X = 0
                If MousePos.X > frmMain.ScaleWidth Then MousePos.X = frmMain.ScaleWidth
            End If
            Moved = 1

        'Move on Y axis
        Case DIMOFS_Y
            If Windowed Then
                OldMousePos = MousePos
                GetCursorPos MousePos
                MousePos.X = MousePos.X - (Me.Left \ Screen.TwipsPerPixelX)
                MousePos.Y = MousePos.Y - (Me.Top \ Screen.TwipsPerPixelY)
                MousePosAdd.X = -(OldMousePos.X - MousePos.X)
                MousePosAdd.Y = -(OldMousePos.Y - MousePos.Y)
            Else
                MousePosAdd.Y = (DevData(LoopC).lData * MouseSpeed)
                MousePos.Y = MousePos.Y + MousePosAdd.Y
                If MousePos.Y < 0 Then MousePos.Y = 0
                If MousePos.Y > ScreenHeight Then MousePos.Y = ScreenHeight
            End If
            Moved = 1
            
        'Mouse wheel is scrolled
        Case DIMOFS_Z
            
            'Scroll the chat buffer if the cursor is over the chat buffer window
            If ShowGameWindow(ChatWindow) Then
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, GameWindow.ChatWindow.Screen.X, GameWindow.ChatWindow.Screen.Y, GameWindow.ChatWindow.Screen.Width, GameWindow.ChatWindow.Screen.Height) Then
                    If DevData(LoopC).lData > 0 Then
                        ChatBufferChunk = ChatBufferChunk + 0.25
                    ElseIf DevData(LoopC).lData < 0 Then
                        ChatBufferChunk = ChatBufferChunk - 0.25
                    End If
                    Engine_UpdateChatArray
                    GoTo NextLoopC
                End If
            End If
            
            'Scroll the quest log if over the quest window
            If ShowGameWindow(QuestLogWindow) Then
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, GameWindow.QuestLog.Screen.X, GameWindow.QuestLog.Screen.Y, GameWindow.QuestLog.Screen.Width, GameWindow.QuestLog.Screen.Height) Then
                    If DevData(LoopC).lData > 0 Then
                        GameWindow.QuestLog.ListStart = GameWindow.QuestLog.ListStart - 1
                        If GameWindow.QuestLog.ListStart < 1 Then GameWindow.QuestLog.ListStart = 1
                    ElseIf DevData(LoopC).lData < 0 Then
                        GameWindow.QuestLog.ListStart = GameWindow.QuestLog.ListStart + 1
                        If GameWindow.QuestLog.ListStart + GameWindow.QuestLog.ListSize > QuestInfoUBound + 1 Then GameWindow.QuestLog.ListStart = QuestInfoUBound - GameWindow.QuestLog.ListSize + 1
                    End If
                    GoTo NextLoopC
                End If
            End If
            
            'Scroll the zoom if the buffer didn't scroll
            If DevData(LoopC).lData > 0 Then
                ZoomLevel = ZoomLevel + (ElapsedTime * 0.001)
                If ZoomLevel > MaxZoomLevel Then ZoomLevel = MaxZoomLevel
            ElseIf DevData(LoopC).lData < 0 Then
                ZoomLevel = ZoomLevel - (ElapsedTime * 0.001)
                If ZoomLevel < 0 Then ZoomLevel = 0
            End If

        'Left button pressed
        Case DIMOFS_BUTTON0
            If DevData(LoopC).lData = 0 Then
                MouseLeftDown = 0
                SelGameWindow = 0
            Else
                If MouseLeftDown = 0 Then   'Clicked down
                    MouseLeftDown = 1
                    Input_Mouse_LeftClick
                End If
            End If

        'Right button pressed
        Case DIMOFS_BUTTON1
            If DevData(LoopC).lData = 0 Then
                MouseRightDown = 0
                Input_Mouse_RightRelease
            Else
                If MouseRightDown = 0 Then  'Clicked down
                    MouseRightDown = 1
                    Input_Mouse_RightClick
                End If
            End If

        End Select

        'Update movement
        If Moved Then
            Input_Mouse_Move

            'Reset move variables
            Moved = 0
            MousePosAdd.X = 0
            MousePosAdd.Y = 0
        End If
        
NextLoopC:

    Next LoopC

ErrOut:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Input_Keys_Down KeyCode
    KeyCode = 0
    Shift = 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Input_Keys_Press KeyAscii
    KeyAscii = 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    KeyCode = 0
    Shift = 0

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Regain focus to Direct Input mouse
    On Error Resume Next
        DIDevice.Acquire
        MousePos.X = X
        MousePos.Y = Y
    On Error GoTo 0
    
End Sub

Private Sub Form_Resize()

    'Regain focus to Direct Input mouse
    On Error Resume Next
        If Not DIDevice Is Nothing Then
            If Not Windowed Then DIDevice.Acquire
        End If
    On Error GoTo 0
    
End Sub

Private Sub GOREsock_Click()

End Sub

Private Sub ShutdownTimer_Timer()
Static FailedUnloads As Long

    On Error Resume Next    'Who cares about an error if we are closing down

    'Quit the client - we must user a timer since DoEvents wont work (since we're not multithreaded)
    
    'Close down the socket
    If FailedUnloads > 5 Or frmMain.GOREsock.ShutDown <> soxERROR Then
        frmMain.GOREsock.UnHook

        'Unload the engine
        Engine_Init_UnloadTileEngine
        
        'Unload the forms
        Engine_UnloadAllForms
        
        'Unload everything else
        End

    Else
        
        'If the socket is making an error on the shutdown sequence for more than a second, just unload anyways
        FailedUnloads = FailedUnloads + 1
        
    End If
            

End Sub

Private Sub GOREsock_OnDataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Retrieve the CommandIDs and send to corresponding data handler
'*********************************************
Dim rBuf As DataBuffer
Dim CommandID As Byte
Dim BufUBound As Long
Dim s As String
Dim i As Byte

    'Packet arrived!
    LastServerPacketTime = timeGetTime

    'Set up the data buffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    BufUBound = UBound(inData)

    'Uncomment this to see packets going into the client
    'Dim i As Long
    'Dim s As String
    'For i = LBound(inData) To UBound(inData)
    '    If inData(i) >= 100 Then
    '        s = s & inData(i) & " "
    '    ElseIf inData(i) >= 10 Then
    '        s = s & "0" & inData(i) & " "
    '    Else
    '        s = s & "00" & inData(i) & " "
    '    End If
    'Next i
    'Debug.Print s
    
    Do
        'Get the Command ID
        CommandID = rBuf.Get_Byte

        '*** LOGIN SERVER ***
        If GettingAccount Then
            
            Select Case CommandID
            Case 0
            Case 1
                MsgBox "Invalid account name.", vbOKOnly
            Case 2
                MsgBox "Invalid account password.", vbOKOnly
            Case 3
                frmConnect.CharLst.Clear
                For i = 1 To 5
                    s = rBuf.Get_String
                    If s <> vbNullString Then frmConnect.CharLst.AddItem s
                Next i
                frmConnect.CharLst.ListIndex = 0
            End Select
        
        '*** GAME SERVER ***
        Else
    
            'Make the appropriate call based on the Command ID
            With DataCode
                Select Case CommandID
    
                Case 0 'This often means there was an offset problem in the packet, adding too many empty values
    
                Case .Comm_Talk: Data_Comm_Talk rBuf
    
                Case .Map_LoadMap: Data_Map_LoadMap rBuf
                Case .Map_SendName:  Data_Map_SendName rBuf
    
                Case .Server_ChangeChar: Data_Server_ChangeChar rBuf
                Case .Server_ChangeCharType: Data_Server_ChangeCharType rBuf
                Case .Server_CharHP: Data_Server_CharHP rBuf
                Case .Server_CharEP: Data_Server_CharEP rBuf
                Case .Server_Connect: Data_Server_Connect
                Case .Server_Disconnect: Data_Server_Disconnect
                Case .Server_EraseChar: Data_Server_EraseChar rBuf
                Case .Server_EraseObject: Data_Server_EraseObject rBuf
                Case .Server_MailBox: Data_Server_Mailbox rBuf
                Case .Server_MailItemRemove: Data_Server_MailItemRemove rBuf
                Case .Server_MailMessage: Data_Server_MailMessage rBuf
                Case .Server_MailObjUpdate: Data_Server_MailObjUpdate rBuf
                Case .Server_MakeChar: Data_Server_MakeChar rBuf
                Case .Server_MakeCharCached: Data_Server_MakeCharCached rBuf
                Case .Server_MakeEffect: Data_Server_MakeEffect rBuf
                Case .Server_MakeSlash: Data_Server_MakeSlash rBuf
                Case .Server_MakeObject: Data_Server_MakeObject rBuf
                Case .Server_MakeProjectile: Data_Server_MakeProjectile rBuf
                Case .Server_Message: Data_Server_Message rBuf
                Case .Server_MoveChar: Data_Server_MoveChar rBuf
                Case .Server_PlaySound: Data_Server_PlaySound rBuf
                Case .Server_PlaySound3D: Data_Server_PlaySound3D rBuf
                Case .Server_RemoveIcon: Data_Server_RemoveIcon rBuf
                Case .Server_SendQuestInfo: Data_Server_SendQuestInfo rBuf
                Case .Server_SetCharDamage: Data_Server_SetCharDamage rBuf
                Case .Server_SetCharSpeed: Data_Server_SetCharSpeed rBuf
                Case .Server_SetIcon: Data_Server_SetIcon rBuf
                Case .Server_SetUserPosition: Data_Server_SetUserPosition rBuf
                Case .Server_UserCharIndex: Data_Server_UserCharIndex rBuf
                Case .Server_WarpChar: Data_Server_WarpChar rBuf
    
                Case .User_Attack: Data_User_Attack rBuf
                Case .User_Bank_Open: Data_User_Bank_Open rBuf
                Case .User_Bank_UpdateSlot: Data_User_Bank_UpdateSlot rBuf
                Case .User_BaseStat: Data_User_BaseStat rBuf
                Case .User_Blink: Data_User_Blink rBuf
                Case .User_CastSkill: Data_User_CastSkill rBuf
                Case .User_ChangeClass: Data_User_ChangeClass rBuf
                Case .User_ChangeServer: Data_User_ChangeServer rBuf
                Case .User_Emote: Data_User_Emote rBuf
                Case .User_Hide: Data_User_Hide rBuf
                Case .User_KnownSkills: Data_User_KnownSkills rBuf
                Case .User_LookLeft: Data_User_LookLeft rBuf
                Case .User_LookRight: Data_User_LookLeft rBuf
                Case .User_ModStat: Data_User_ModStat rBuf
                Case .User_Profile: Data_User_Profile rBuf
                Case .User_Rotate: Data_User_Rotate rBuf
                Case .User_SendClass: Data_User_SendClass rBuf
                Case .User_SetInventorySlot: Data_User_SetInventorySlot rBuf
                Case .User_SetRage: Data_User_SetRage rBuf
                Case .User_SetSkillDelay: Data_User_SetSkillDelay rBuf
                Case .User_SetWeaponRange: Data_User_SetWeaponRange rBuf
                Case .User_Target: Data_User_Target rBuf
                Case .User_Trade_Accept: Data_User_Trade_Accept rBuf
                Case .User_Trade_Cancel: Data_User_Trade_Cancel
                Case .User_Trade_StartNPCTrade: Data_User_Trade_StartNPCTrade rBuf
                Case .User_Trade_Trade: Data_User_Trade_Trade rBuf
                Case .User_Trade_UpdateTrade: Data_User_Trade_UpdateTrade rBuf
    
                Case .Combo_ProjectileSoundRotateDamage: Data_Combo_ProjectileSoundRotateDamage rBuf
                Case .Combo_SlashSoundRotateDamage: Data_Combo_SlashSoundRotateDamage rBuf
                Case .Combo_SoundRotateDamage: Data_Combo_SoundRotateDamage rBuf
    
                Case Else
                    rBuf.Overflow  'Something went wrong or we hit the end, either way, RUN!!!!
    
                End Select
            End With

        End If

        'Exit when the buffer runs out
        If rBuf.Get_ReadPos > BufUBound Then Exit Do

    Loop
    
    'If we got data from the account server, close since we don't need anything more than a single packet
    If GettingAccount Then
        frmMain.GOREsock.Shut SoxID
        frmMain.GOREsock.ShutDown
        GettingAccount = False
    End If
    
    Set rBuf = Nothing

End Sub

Private Sub GOREsock_OnConnecting(inSox As Long)

    If SocketOpen = 0 Then
        
        If GettingAccount Then
        
            Sleep 50
            DoEvents
            
            'Build the packet
            sndBuf.Clear
            sndBuf.Put_Byte 1
            sndBuf.Put_String Trim$(frmConnect.NameTxt.Text)
            sndBuf.Put_String Trim$(frmConnect.PasswordTxt.Text)
            
            'Send the data
            Data_Send
            DoEvents
        
        Else
            
            Sleep 50
            DoEvents
            
            'Send the packet
            sndBuf.Put_Byte DataCode.User_Login
            sndBuf.Put_String Trim$(frmConnect.NameTxt.Text)
            sndBuf.Put_String UserName
            sndBuf.Put_String UserPassword

            'Save Game.ini
            If Not SavePass Then UserPassword = vbNullString
            Var_Write DataPath & "Game.ini", "INIT", "Name", UserName
            Var_Write DataPath & "Game.ini", "INIT", "Password", UserPassword
    
            'Send the data
            Data_Send
            DoEvents
        
        End If
    
    End If
    
End Sub
