VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Music"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'ファイルの読み込み
Public Function OpenSound(ByVal MusicPath As String) As Integer
    Dim cmd As String
    cmd = "Open " + Chr(34) + MusicPath & Chr(34) & " alias " & aliasName
    OpenSound = mciSendString(cmd, "", 0, 0)
End Function

'再生
Public Function Play() As Integer
    Play = mciSendString("Play " & aliasName, "", 0, 0)
    'Play = mciSendString("Play " & Chr(34) & MusicPath & Chr(34), "", 0, 0)
End Function

'一時停止
Public Function Pause() As Integer
    Call mciSendString("Pause " & aliasName, "", 0, 0)
    'Call mciSendString("Pause " & Chr(34) & MusicPath & Chr(34), "", 0, 0)
End Function

'再開
Public Function PlayResume() As Integer
    Call mciSendString("Resume " & aliasName, "", 0, 0)
End Function

'停止
Public Function CloseAll() As Integer
    CloseAll = mciSendString("Close All", "", 0, 0)
    Call TimerStop
End Function

'指定位置再生
Public Sub PlayPosition(ByVal PositionSec As Long)
    Call mciSendString("Seek " & aliasName & " to " & PositionSec * 1000, "", 0, 0)
    Call mciSendString("Play " & aliasName, "", 0, 0)
End Sub
