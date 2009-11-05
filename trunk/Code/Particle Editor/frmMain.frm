VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Particle Editor"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DispEffect As Byte  'Index of the displayed effect

Private ResetX As Single
Private ResetY As Single
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
Dim i As Long
        
    Randomize
    MapInfo.Width = 25500
    MapInfo.Height = 25500
    
    'Init particle engine
    Me.Show
    InitFilePaths
    Engine_Init_TileEngine

    'Set initial reset position (center screen)
    ResetX = frmMain.ScaleWidth * 0.5
    ResetY = frmMain.ScaleHeight * 0.5

    'Create initial effect
    ResetEffect

    'Main loop
    EngineRun = True

    Do While EngineRun

        If DispEffect > 0 Then
            'If Not Effect(DispEffect).Used Then ResetEffect
        End If
        
        'Draw
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, -10197916, 1#, 0
        D3DDevice.BeginScene
        Effect_UpdateAll
        D3DDevice.EndScene
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

        'FPS
        ElapsedTime = Engine_ElapsedTime()
        If ElapsedTime < 16 Then
            Sleep 16 - ElapsedTime
        End If
        If FPS_Last_Check + 1000 < timeGetTime Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            FPS_Last_Check = timeGetTime
            frmMain.Caption = "Particle Editor: FPS " & FPS
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If

        DoEvents

    Loop
    
    'Clear arrays
    Erase CharList()
    Erase Effect()

    'Unload engine
    Engine_Init_UnloadTileEngine
    Engine_UnloadAllForms

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Select Case Button
        Case vbLeftButton: ResetEffect
        Case vbRightButton: Effect_Kill 0, True
    End Select
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Stop the engine
    EngineRun = False

End Sub

Private Sub ResetEffect()
    
    'Resets the effect - use this sub to change the effect displayed
    'DispEffect = Effect_ChangeClass_Begin(Me.ScaleWidth * 0.5, Me.ScaleHeight * 0.5, 7, 100, 23)
    DispEffect = Effect_BloodSpray_Begin(Me.ScaleWidth * 0.5, Me.ScaleHeight * 0.5, 50, Rnd * 360)
    'DispEffect = Effect_BloodSplatter_Begin(Me.ScaleWidth * 0.5, Me.ScaleHeight * 0.5, 100)

End Sub
