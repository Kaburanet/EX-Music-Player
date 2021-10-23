Attribute VB_Name = "MainModule"
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare PtrSafe Function SetTimer Lib "USER32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As Long
Public Declare PtrSafe Function KillTimer Lib "USER32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public status As String
Public Repeat As String
Public CalledFrom As String '呼び出し元のリストボックス名
Public CurrentMusicPath As String
Public CurrentSec As Long
Public artist As String
Public Title As String

Public Const aliasName As String = "EXMusicPlayer"

Type MusicInfomation
    Lyrics As String        '歌詞
    Title As String         '曲名
    Singer As String        '歌手
    SongWriter As String    '作詞者
    Composer As String      '作曲者
    FirstLine As String     '歌いだし
    TieupInfo As String     'タイアップ情報
End Type



Public mTimerID As Long

Sub TimerProc()
    If mTimerID = 0 Then End '終了できない時の対策
    On Error Resume Next 'デバッグ出すとExcelが固まるので
    
    Dim CInfo As New Info
    Dim Length As Double: Length = CInfo.GetPosition
    MainForm.CurrentTimeLabel.Caption = CInfo.Convert_Sec_To_Min(Length)
    MainForm.MusicScrollBar.Value = Length
    CurrentSec = Int(Length)
End Sub

Sub TimerStart()
    If mTimerID <> 0 Then
        MsgBox "起動済です。"
        Exit Sub
    End If
    mTimerID = SetTimer(0&, 1&, 1000&, AddressOf TimerProc)
End Sub

Sub TimerStop()
    Call KillTimer(0&, mTimerID)
    mTimerID = 0
End Sub



Sub Sound_Close()
    Call mciSendString("Close All", "", 0, 0)
End Sub



Sub StartMusicPlayer()
    MainForm.Show vbModeless
End Sub
