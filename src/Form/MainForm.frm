VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "EX Music Player"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10212
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AlbumBackCommandButton_Click()
    If AlbumListBox.Tag = 2 Then
        AlbumListBox.ColumnWidths = 0
        AlbumListBox.Clear
        'アルバムリストを追加
        Dim i As Long
        Dim album() As String: album = Distinct("Album")
        For i = 0 To UBound(album)
            AlbumListBox.AddItem album(i)
            Call setListBoxColumnWidths(album(i), AlbumListBox)
        Next i
        
        AlbumListBox.Tag = 1
    End If
    
End Sub

Private Sub AlbumCommandButton_Click()
    MainMultiPage.Value = 1
End Sub

Private Sub AlbumListBox_Click()
    If AlbumListBox.Tag = 1 Then
        AlbumListBox.ColumnWidths = 0
        Dim SelectedAlbum As String: SelectedAlbum = AlbumListBox.List(AlbumListBox.ListIndex, 0)
        AlbumListBox.Clear
        
        Dim i As Long
        For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
            If Cells(i, 4).Value = SelectedAlbum Then
                AlbumListBox.AddItem Cells(i, 1).Value
                Call setListBoxColumnWidths(Cells(i, 1).Value, AlbumListBox)
            End If
        Next i
        
        AlbumListBox.Tag = 2
    End If

End Sub

Private Sub AlbumListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If AlbumListBox.Tag = 2 Then
    
        CurrentMusicPath = AlbumListBox.List(AlbumListBox.ListIndex, 0)
                  
        Call StartPlay
                  
        status = "play"
        
        CalledFrom = "Album"
    End If
    
End Sub

Private Sub AllSongsCommandButton_Click()
    MainMultiPage.Value = 0
End Sub

Private Sub AllSongsListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
     
    CurrentMusicPath = AllSongsListBox.List(AllSongsListBox.ListIndex, 0)
              
    Call StartPlay
              
    status = "play"
    
    CalledFrom = "All"
    
End Sub

Private Sub ArtistBackCommandButton_Click()
    If ArtistListBox.Tag = 2 Then
        ArtistListBox.ColumnWidths = 0
        ArtistListBox.Clear
        'アルバムリストを追加
        Dim i As Long
        Dim artist() As String: artist = Distinct("Artist")
        For i = 0 To UBound(artist)
            ArtistListBox.AddItem artist(i)
            Call setListBoxColumnWidths(artist(i), ArtistListBox)
        Next i
        
        ArtistListBox.Tag = 1
    End If
    
End Sub

Private Sub ArtistCommandButton_Click()
    MainMultiPage.Value = 2
End Sub

Private Sub ArtistListBox_Click()
    If ArtistListBox.Tag = 1 Then
        ArtistListBox.ColumnWidths = 0
        Dim SelectedArtist As String: SelectedArtist = ArtistListBox.List(ArtistListBox.ListIndex, 0)
        ArtistListBox.Clear
        
        Dim i As Long
        For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
            If Cells(i, 3).Value = SelectedArtist Then
                ArtistListBox.AddItem Cells(i, 1).Value
                Call setListBoxColumnWidths(Cells(i, 1).Value, ArtistListBox)
            End If
        Next i
        
        ArtistListBox.Tag = 2
    End If
End Sub

Private Sub ArtistListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
     If ArtistListBox.Tag = 2 Then
        CurrentMusicPath = ArtistListBox.List(ArtistListBox.ListIndex, 0)
                  
        Call StartPlay
                  
        status = "play"
        
        CalledFrom = "Artist"
        
    End If
End Sub

Private Sub LyricsCommandButton_Click()
    MainMultiPage.Value = 4
End Sub

Private Sub LyricsSearchCommandButton_Click()
    Dim CLyrics As New Lyrics
    Dim tmp As MusicInfomation
    'tmp = CLyrics.FindLyrics(Title, artist)
    tmp = CLyrics.FindLyrics(Title, "")
    LyricsTextBox.Text = tmp.Lyrics
    LyricsTextBox.SelStart = 0
    MainMultiPage.Value = 4
End Sub

Private Sub MusicScrollBar_Change()
    
    '再生中の位置より3以上ずれていたら手動で動かしたとみなす
    If Abs(MusicScrollBar.Value - CurrentSec) >= 3 Then
        Dim CMusic As Music
        Set CMusic = New Music
        Call CMusic.PlayPosition(MusicScrollBar.Value)
    End If
    
    '終わりまで達したら再生を終了
    If MusicScrollBar.Value = MusicScrollBar.Max Then
        If Repeat = "R1" Then
            Call StartPlay
        ElseIf Repeat = "RA" Then
            Call GoNext
        Else
            Call StopPlay
        End If
    End If
    
End Sub

Private Sub PlayListCommandButton_Click()
    MainMultiPage.Value = 3
End Sub

Private Sub PlayListListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CurrentMusicPath = PlayListListBox.List(PlayListListBox.ListIndex, 0)
              
    Call StartPlay
              
    status = "play"
    
    CalledFrom = "PlayList"
End Sub

Private Sub RepeatCommandButton_Click()
    If RepeatCommandButton.Caption = "R(X)" Then
        RepeatCommandButton.Caption = "R(1)"
        Repeat = "R1"
    ElseIf RepeatCommandButton.Caption = "R(1)" Then
        RepeatCommandButton.Caption = "R(A)"
        Repeat = "RA"
    ElseIf RepeatCommandButton.Caption = "R(A)" Then
        RepeatCommandButton.Caption = "R(X)"
        Repeat = "RX"
    End If
End Sub


Private Sub SendToPlayListCommandButton1_Click()
    If AlbumListBox.Tag = "2" Then
        If AlbumListBox.ListIndex <> -1 Then
            PlayListListBox.AddItem AlbumListBox.List(AlbumListBox.ListIndex, 0)
            Call setListBoxColumnWidths(AlbumListBox.List(AlbumListBox.ListIndex, 0), PlayListListBox)
        End If
    End If
End Sub

Private Sub SendToPlayListCommandButton2_Click()
    If ArtistListBox.Tag = "2" Then
        If ArtistListBox.ListIndex <> -1 Then
            PlayListListBox.AddItem ArtistListBox.List(ArtistListBox.ListIndex, 0)
            Call setListBoxColumnWidths(ArtistListBox.List(ArtistListBox.ListIndex, 0), PlayListListBox)
        End If
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim i As Long
    
    '何も追加されていない場合は全部のリストを追加
    If WorksheetFunction.CountA(Range("A:A")) = 0 Then
        Dim CFile As New File
        Dim ls As Boolean: ls = CFile.Get_FilePaths(Cells(2,10).Value)
    End If
    
    
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        AllSongsListBox.AddItem (Cells(i, 1).Value)
        Call setListBoxColumnWidths(Cells(i, 1).Value, AllSongsListBox)
    Next i
    
    'アルバムリストを追加
    Dim album() As String: album = Distinct("Album")
    For i = 0 To UBound(album)
        AlbumListBox.AddItem album(i)
        Call setListBoxColumnWidths(album(i), AlbumListBox)
    Next i
    
    'アーティストリストを追加
    Dim artist() As String: artist = Distinct("Artist")
     For i = 0 To UBound(artist)
        ArtistListBox.AddItem artist(i)
        Call setListBoxColumnWidths(artist(i), ArtistListBox)
    Next i
    
    
    status = "stop"
    CurrentMusicPath = ""
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call StopPlay
End Sub

Private Sub BackCommandButton_Click()
    
    Dim Ctrl As String
    If CalledFrom = "" Then
        Exit Sub
    ElseIf CalledFrom = "All" Then
       Ctrl = "AllSongsListBox"
    ElseIf CalledFrom = "Album" Then
        If AlbumListBox.Tag = 1 Then
            Exit Sub
        Else
            Ctrl = "AlbumListBox"
        End If
    ElseIf CalledFrom = "Artist" Then
        If ArtistListBox.Tag = 1 Then
            Exit Sub
        Else
            Ctrl = "ArtistListBox"
        End If
    ElseIf CalledFrom = "PlayList" Then
        Ctrl = "PlayListListBox"
    End If
    
    Dim i As Long
    'リストボックス内を検索してひとつ前のインデックスを取得
    For i = 0 To Controls(Ctrl).ListCount - 1
        If Controls(Ctrl).List(i) = CurrentMusicPath Then
            Dim index As Long
            If i > 0 Then
                Controls(Ctrl).ListIndex = i - 1
                index = i - 1
            Else
                Controls(Ctrl).ListIndex = Controls(Ctrl).ListCount - 1
                index = Controls(Ctrl).ListCount - 1
            End If
            
            Call StopPlay
            CurrentMusicPath = Controls(Ctrl).List(index)
            Call StartPlay
            
            Exit For
        End If
    Next i
    
End Sub

Private Sub PlayStopCommandButton_Click()
    Dim CMusic As New Music
    Dim CInfo As New Info
    
    If CurrentMusicPath = "" Then Exit Sub

    If PlayStopCommandButton.Caption = "|>" Then
        PlayStopCommandButton.Caption = "| |"
        
        '一時停止中なら再開、それ以外は普通に再生
        If status = "stop" Then

            Call StartPlay

        ElseIf status = "pause" Then
            Call CMusic.PlayResume
            
            status = "play"
        End If
        
    ElseIf PlayStopCommandButton.Caption = "| |" Then
        PlayStopCommandButton.Caption = "|>"
        
        Call CMusic.Pause
        
        status = "pause"
    End If
    
End Sub

Private Sub NextCommandButton_Click()
   Call GoNext
End Sub

Private Sub StartPlay()
    Dim CMusic As New Music
    Dim CInfo As New Info
    Dim MusicLength As Integer
    
    Call CMusic.CloseAll
    
    Call CMusic.OpenSound(CurrentMusicPath)
    
    MusicLength = Int(CInfo.GetLength)
    TimeLabel.Caption = CInfo.Convert_Sec_To_Min(MusicLength)
    CurrentTimeLabel.Caption = CInfo.Convert_Sec_To_Min(0)
    MusicScrollBar.Max = MusicLength
    InfoTextBox.Text = CInfo.GetTitle(CurrentMusicPath) & vbCrLf & CInfo.GetArtist(CurrentMusicPath)
    
    Call CMusic.Play
    
    Call TimerStart
    
    
    PlayStopCommandButton.Caption = "| |"
    status = "play"
    
End Sub

Private Sub StopPlay()
    Dim CMusic As New Music
    Dim CInfo As New Info
        
    Call CMusic.CloseAll
    
    PlayStopCommandButton.Caption = "|>"
    MusicScrollBar.Value = 0
    MusicScrollBar.Max = 100
    TimeLabel.Caption = "00:00"
    CurrentTimeLabel.Caption = "00:00"
    
    status = "stop"
    
End Sub

Private Function Distinct(ByVal FindType As String) As String()
    Dim ret() As String
    
    Dim myCollection As New Collection
    
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        If FindType = "Artist" Then
            myCollection.Add Cells(i, 3).Value, Cells(i, 3).Value
        ElseIf FindType = "Album" Then
            myCollection.Add Cells(i, 4).Value, Cells(i, 4).Value
        End If
    Next i
    On Error GoTo 0
    
    For i = 1 To myCollection.Count
        ReDim Preserve ret(i - 1) As String
        ret(i - 1) = myCollection(i)
    Next i
    
    Distinct = ret
    
    
End Function

Private Sub GoNext()
    Dim Ctrl As String
    If CalledFrom = "" Then
        Exit Sub
    ElseIf CalledFrom = "All" Then
       Ctrl = "AllSongsListBox"
    ElseIf CalledFrom = "Album" Then
        If AlbumListBox.Tag = 1 Then
            Exit Sub
        Else
            Ctrl = "AlbumListBox"
        End If
    ElseIf CalledFrom = "Artist" Then
        If ArtistListBox.Tag = 1 Then
            Exit Sub
        Else
            Ctrl = "ArtistListBox"
        End If
    ElseIf CalledFrom = "PlayList" Then
        Ctrl = "PlayListListBox"
    End If
    
    Dim i As Long
    'リストボックス内を検索してひとつ後のインデックスを取得
    For i = 0 To Controls(Ctrl).ListCount - 1
        If Controls(Ctrl).List(i) = CurrentMusicPath Then
            Dim index As Long
            If i < Controls(Ctrl).ListCount - 1 Then
                Controls(Ctrl).ListIndex = i + 1
                index = i + 1
            Else
                Controls(Ctrl).ListIndex = 0
                index = 0
            End If
            
            Call StopPlay
            CurrentMusicPath = Controls(Ctrl).List(index)
            Call StartPlay
            
            Exit For
        End If
    Next i
End Sub


Private Function setListBoxColumnWidths(ByVal Str As String, ByVal lsb As Object)
    Dim Length As Integer: Length = LenB(StrConv(Str, vbFromUnicode)) * 4.7
    
    If (Val(lsb.ColumnWidths) < Length) Then
        lsb.ColumnWidths = Length
    End If
End Function
