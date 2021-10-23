VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lyrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'歌詞や人気曲、ランキングなどをインターネットから取得する
'その他検索結果も取得してくる

Dim MusicInfo() As MusicInfomation

'Type MusicInfomation
'    Lyrics As String        '歌詞
'    Title As String         '曲名
'    Singer As String        '歌手
'    SongWriter As String    '作詞者
'    Composer As String      '作曲者
'    FirstLine As String     '歌いだし
'    TieupInfo As String     'タイアップ情報
'End Type

Public Function FindLyrics(ByVal Title As String, OtherSearchWords As String) As MusicInfomation
    Dim ret As MusicInfomation  '結果
    
    '順に歌詞、曲名、歌手、作詞、作曲、歌いだし、タイアップ情報
    Dim MI() As MusicInfomation: MI = Get_Info(Title)   '一旦情報を取得して入れておく
    Dim flag As Boolean
    Dim Is_Found As Boolean: Is_Found = False
    Dim i As Long
    
    If Not Not MI Then  '配列が空じゃないとき
    
        For i = 0 To UBound(MI) - 1
            flag = False
            
            If OtherSearchWords = "" Then
                flag = True
            End If
            
            If InStr(MI(i).Singer, OtherSearchWords) <> 0 Then
                flag = True
            ElseIf InStr(MI(i).SongWriter, OtherSearchWords) <> 0 Then
                flag = True
            ElseIf InStr(MI(i).Composer, OtherSearchWords) <> 0 Then
                flag = True
            ElseIf InStr(MI(i).FirstLine, OtherSearchWords) <> 0 Then
                flag = True
            ElseIf InStr(MI(i).TieupInfo, OtherSearchWords) <> 0 Then
                flag = True
            End If
            
            If flag = True Then
                ret.Lyrics = MI(i).Lyrics
                ret.Title = MI(i).Title
                ret.Singer = MI(i).Singer
                ret.SongWriter = MI(i).SongWriter
                ret.Composer = MI(i).Composer
                ret.FirstLine = MI(i).FirstLine
                ret.TieupInfo = MI(i).TieupInfo
                Is_Found = True
            End If
        Next
        
        If Is_Found = False Then
            ret.Lyrics = "NotFound"
        End If
        
    Else
        ret.Lyrics = "NotFound"
    End If
    
    FindLyrics = ret
    
End Function

'HTMLの取得
Public Function Get_HTML(ByVal Path As String) As String
    On Error Resume Next
    Dim html As String
    Dim http As Object
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", Path, False
    http.Send
    
    ' ダウンロード待ち
    Do While http.readyState <> 4
        DoEvents
    Loop
    
    ' 結果出力
    html = http.responseText
    
    Get_HTML = html
    
    
    '解放
    Set http = Nothing
    
    
End Function

'Shift-JISにしてエンコード
Public Function URLEnc_SHIFT_JIS(ByVal txt As String) As String
    Dim bin() As Byte
    Dim i As Integer
    Dim res As String
    bin = StrConv(txt, vbFromUnicode) 'Shift-Jisにしてバイナリ配列へ
    'バイナリ配列を文字列に変更
    For i = 0 To UBound(bin)
        Select Case bin(i)
            Case &H31 To &H39, &H41 To &H5A, &H61 To &H7A '0-9,A-Z,a-z なら
            res = res & Chr(bin(i)) 'そのままの文字
        Case Else 'それ以外のエンコード対象文字なら
            res = res & "%" & Hex(bin(i)) 'エンコード
            'res = res & "%" & LCase(Hex(bin(i))) '小文字にする場合
        End Select
    Next
    
    URLEnc_SHIFT_JIS = res
End Function


Public Function EncodeURL_MusicTitle(ByVal MusicTitle As String) As String
    Dim EncodedTitle As String  'エンコードされた曲名
    EncodedTitle = URLEnc_SHIFT_JIS(MusicTitle)
    
    Dim SearchURL As String '歌詞サイトの検索結果のURL
    SearchURL = "https://www.utamap.com/searchkasi.php?searchname=title&word=" & EncodedTitle & "&act=search&search_by_keyword=%8C%9F%26%23160%3B%26%23160%3B%26%23160%3B%8D%F5&sortname=1&pattern=4"
    
    EncodeURL_MusicTitle = SearchURL
    
End Function





'*********************************************************************************歌詞********************************************************************************************

'インデックスから最終的な歌詞を取得
Public Function Get_Lyrics_FromIndex(ByVal LyricsIndex As String) As String
    On Error Resume Next
    Dim Path As String: Path = ("https://www.utamap.com/phpflash/flashfalsephp.php?unum=" & LyricsIndex)
    Dim html As String
    Dim http As Object
    
    Dim Enc As Object
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", Path, False
    http.Send
    
    ' ダウンロード待ち
    Do While http.readyState <> 4
        DoEvents
    Loop
    
    Set Enc = CreateObject("ADODB.Stream")
    Enc.Open
    Enc.Position = 0
    Enc.Type = 1 'adTypeBinary
    Enc.write http.responseBody
    
    Enc.Position = 0
    Enc.Type = 2 'adTypeText
    Enc.Charset = "UTF-8" 'UTF-8
    
    html = Enc.ReadText '変換済みのテキストを結果とする
    Enc.Close
    
    Dim index As Integer
    index = InStr(html, "=") + 1
    index = InStr(index, html, "=") + 1
    html = Mid(html, index, Len(html) - index + 1)
    
    ' 結果出力
    Get_Lyrics_FromIndex = html
    
    
    '解放
    Set http = Nothing
    Set Enc = Nothing
    
End Function

'*****************************************************************歌手・作詞・作曲・歌いだし・タイアップ情報**********************************************************************
'曲情報を一括取得
Public Function Get_Info(ByVal MusicTitle As String) As MusicInfomation()
    
    'HTMLを取得
    Dim utamapHTML As String: utamapHTML = Get_HTML(EncodeURL_MusicTitle(MusicTitle))
    
    Dim Info() As MusicInfomation           '順に歌詞、曲名、歌手、作詞、作曲、歌いだし、タイアップ情報
    Dim Counter As Long: Counter = 0        'カウンタ
    Dim HitNum As Long                      'ヒットした曲数の総数
    
    Dim sIndex As Long  '始まり
    Dim eIndex As Long  '終わり
    
    
    'まず初めにヒット件数を取得する
    sIndex = InStr(1, utamapHTML, "検索キーワード:") + Len("検索キーワード:") + Len(MusicTitle) + Len(" &nbsp; 全 ")
    eIndex = InStr(sIndex, utamapHTML, "件") - 1
    HitNum = CLng(Mid(utamapHTML, sIndex, eIndex - sIndex))
    If HitNum = 0 Then HitNum = -1
    
    sIndex = 1
    eIndex = 1
    
    Do While (Counter <= HitNum)
    
        ReDim Preserve Info(Counter) As MusicInfomation
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct160>") + Len("<TD class=ct160>")
        
        '歌詞を取得する
        '曲のIDを取得
        Dim MusicID As String
        sIndex = InStr(sIndex, utamapHTML, "./showkasi.php?surl=") + Len("./showkasi.php?surl=")
        eIndex = InStr(sIndex, utamapHTML, ">") - 1
        MusicID = Mid(utamapHTML, sIndex, eIndex - sIndex)
        Info(Counter).Lyrics = Get_Lyrics_FromIndex(MusicID)
        
        
        'タイトルを取得する
        sIndex = InStr(sIndex, utamapHTML, ">") + Len(">")
        eIndex = InStr(sIndex, utamapHTML, "</A></td>")
        Info(Counter).Title = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '開始点の更新
        sIndex = eIndex - 1
        
        '明らかに文字数が多すぎるときはデータが取得できていないとしてエラーとする(ここをチェックしておけば最初で判断できる)
        If Len(Info(Counter).Title) > 50 Then
            Info(Counter).Lyrics = "_ERROR"
            Info(Counter).Title = "_ERROR"
            Info(Counter).Singer = "_ERROR"
            Info(Counter).SongWriter = "_ERROR"
            Info(Counter).Composer = "_ERROR"
            Info(Counter).FirstLine = "_ERROR"
            Info(Counter).TieupInfo = "_ERROR"
        End If
        
        
        '歌手を取得する
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct120>") + Len("<TD class=ct120>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).Singer = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '開始点の更新
        sIndex = eIndex - 1
        
        
        '作詞者を取得する
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct120>") + Len("<TD class=ct120>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).SongWriter = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '開始点の更新
        sIndex = eIndex - 1
        
        
        '作曲者を取得する
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct120>") + Len("<TD class=ct120>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).Composer = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '開始点の更新
        sIndex = eIndex - 1
        
        '歌いだしを取得する
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct160>") + Len("<TD class=ct160>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).FirstLine = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '開始点の更新
        sIndex = eIndex - 1
        
        'タイアップ情報を取得する
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct>") + Len("<TD class=ct>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).TieupInfo = Mid(utamapHTML, sIndex, eIndex - sIndex)
        If Info(Counter).TieupInfo = "" Or Info(Counter).TieupInfo = "-" Then Info(Counter).TieupInfo = "なし"
        If Len(Info(Counter).TieupInfo) > 50 Then Info(Counter).TieupInfo = "なし"
        sIndex = eIndex - 1
        
        Counter = Counter + 1
    
    Loop
    
    Get_Info = Info
    
End Function

'*********************************************************************************************************************************************************************************




