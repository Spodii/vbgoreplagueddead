Attribute VB_Name = "General"
Option Explicit

'Dummy stuff
Public ServerID As Long
Public Enum LogType
    General = 0
    CodeTracker = 1
    PacketIn = 2
    PacketOut = 3
    CriticalError = 4
    InvalidPacketData = 5
End Enum
#If False Then
Private General, CodeTracker, PacketIn, PacketOut, CriticalError, InvalidPacketData
#End If

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Public Sub Log(s As String, b As LogType)
    'Dummy routine
End Sub

Public Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

    Var_Get = Space$(1000)
    GetPrivateProfileString Main, Var, vbNullString, Var_Get, 1000, File
    Var_Get = RTrim$(Var_Get)
    If LenB(Var_Get) <> 0 Then Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function

Public Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

    WritePrivateProfileString Main, Var, Value, File

End Sub

Public Sub Server_Unload()

    Unload frmMain

End Sub
