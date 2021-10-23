Attribute VB_Name = "MainModule"
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare PtrSafe Function SetTimer Lib "USER32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As Long
Public Declare PtrSafe Function KillTimer Lib "USER32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public status As String
Public Repeat As String
Public CalledFrom As String '�Ăяo�����̃��X�g�{�b�N�X��
Public CurrentMusicPath As String
Public CurrentSec As Long
Public artist As String
Public Title As String

Public Const aliasName As String = "EXMusicPlayer"

Type MusicInfomation
    Lyrics As String        '�̎�
    Title As String         '�Ȗ�
    Singer As String        '�̎�
    SongWriter As String    '�쎌��
    Composer As String      '��Ȏ�
    FirstLine As String     '�̂�����
    TieupInfo As String     '�^�C�A�b�v���
End Type



Public mTimerID As Long

Sub TimerProc()
    If mTimerID = 0 Then End '�I���ł��Ȃ����̑΍�
    On Error Resume Next '�f�o�b�O�o����Excel���ł܂�̂�
    
    Dim CInfo As New Info
    Dim Length As Double: Length = CInfo.GetPosition
    MainForm.CurrentTimeLabel.Caption = CInfo.Convert_Sec_To_Min(Length)
    MainForm.MusicScrollBar.Value = Length
    CurrentSec = Int(Length)
End Sub

Sub TimerStart()
    If mTimerID <> 0 Then
        MsgBox "�N���ςł��B"
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
