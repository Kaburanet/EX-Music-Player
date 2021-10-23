VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'特定のフォルダ以下にあるmp3ファイルの名前をすべて取得する
Public Function Get_Files(ByVal DirPath As String) As String()
    Dim ret() As String
    Dim buf As String, Counter As Long
    Counter = 0
    
    buf = Dir(DirPath + "\*.mp3")
    Do While buf <> ""
        ReDim Preserve ret(Counter) As String
        ret(Counter) = buf
        buf = Dir()
        Counter = Counter + 1
    Loop
    
    Get_Files = ret
End Function

'特定のフォルダ以下にあるmp3ファイルの絶対パスを取得する(戻り値はDictionary型)
Public Function Get_FilePaths(ByVal DirPath As String) As Boolean

    Range("A:F").Clear
    
    Call SearchAllFiles(DirPath)
    
    Get_FilePaths = True
    
End Function

Public Sub SearchAllFiles(ByVal DirPath As String)
    
    Dim CInfo As New Info
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Folder As Object, SubFolder As Object
    Set Folder = fso.GetFolder(DirPath)
    
    For Each SubFolder In Folder.SubFolders
         Call SearchAllFiles(SubFolder.Path)
    Next
    
    Dim f As Object
    Dim LastCellRow As Long: LastCellRow = 1
    For Each f In Folder.Files
        If fso.GetExtensionName(f.Path) = "wav" Or fso.GetExtensionName(f.Path) = "mp3" Then
            Cells(LastCellRow, 1).Value = f.Path
            Cells(LastCellRow, 2).Value = CInfo.GetTitle(f.Path)
            Cells(LastCellRow, 3).Value = CInfo.GetArtist(f.Path)
            Cells(LastCellRow, 4).Value = CInfo.GetAlbum(f.Path)
            LastCellRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
        End If
        
    Next
    
End Sub
