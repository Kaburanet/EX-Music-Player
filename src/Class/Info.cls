VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


'曲の長さをファイルから読み取る
Public Function GetLength() As Double
    
    Dim RetBuffer As String * 20
    Dim MCICommandString As String
    
    MCICommandString = "status " & aliasName & " Length"
    
    Call mciSendString(MCICommandString, RetBuffer, Len(RetBuffer), 0)
        
    GetLength = Val(RetBuffer) / 1000
    
End Function

'現在位置
Public Function GetPosition() As Integer
    Dim RetBuffer As String * 20
    Dim MCICommandString As String
    
    MCICommandString = "status " & aliasName & " Position"
    
    Call mciSendString(MCICommandString, RetBuffer, Len(RetBuffer), 0)
        
    GetPosition = Val(RetBuffer) / 1000
    
End Function


Public Function Convert_Sec_To_Min(ByVal Length As Double) As String
    Dim ret As String
    
    Dim Min As Integer
    Dim Sec As Integer
    Min = Int(Length / 60)
    Sec = Int(Length - Min * 60)
    
    ret = CStr(Min) + ":" + Format(CStr(Sec), "00")
    
    Convert_Sec_To_Min = ret
    
End Function


Public Function GetArtist(ByVal FilePath As String) As String
    Dim ret As String   '戻り値

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    Dim shFolder As Object, shFile As Object
    
    Set shFolder = sh.Namespace(fso.GetParentFolderName(FilePath))
    
    Set shFile = shFolder.ParseName(fso.GetFileName(FilePath))
    ret = shFolder.GetDetailsOf(shFile, 13) '参加アーティストのプロパティは13
    
    If ret = "" Then ret = "不明なアーティスト"
    
    artist = ret
    
    GetArtist = ret
    
End Function

Public Function GetTitle(ByVal FilePath As String) As String
    Dim ret As String   '戻り値

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    Dim shFolder As Object, shFile As Object
    
    Set shFolder = sh.Namespace(fso.GetParentFolderName(FilePath))
    
    Set shFile = shFolder.ParseName(fso.GetFileName(FilePath))
    ret = shFolder.GetDetailsOf(shFile, 21) '曲名のプロパティは13
    
    If ret = "" Then ret = "不明なタイトル"
    
    Title = ret
    
    GetTitle = ret
End Function

Public Function GetAlbum(ByVal FilePath As String) As String
    Dim ret As String   '戻り値

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    Dim shFolder As Object, shFile As Object
    
    Set shFolder = sh.Namespace(fso.GetParentFolderName(FilePath))
    
    Set shFile = shFolder.ParseName(fso.GetFileName(FilePath))
    ret = shFolder.GetDetailsOf(shFile, 14) 'アルバムのプロパティは14
    
    If ret = "" Then ret = "不明なアルバム"
    
    Title = ret
    
    GetAlbum = ret
End Function
