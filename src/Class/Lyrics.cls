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

'�̎���l�C�ȁA�����L���O�Ȃǂ��C���^�[�l�b�g����擾����
'���̑��������ʂ��擾���Ă���

Dim MusicInfo() As MusicInfomation

'Type MusicInfomation
'    Lyrics As String        '�̎�
'    Title As String         '�Ȗ�
'    Singer As String        '�̎�
'    SongWriter As String    '�쎌��
'    Composer As String      '��Ȏ�
'    FirstLine As String     '�̂�����
'    TieupInfo As String     '�^�C�A�b�v���
'End Type

Public Function FindLyrics(ByVal Title As String, OtherSearchWords As String) As MusicInfomation
    Dim ret As MusicInfomation  '����
    
    '���ɉ̎��A�Ȗ��A�̎�A�쎌�A��ȁA�̂������A�^�C�A�b�v���
    Dim MI() As MusicInfomation: MI = Get_Info(Title)   '��U�����擾���ē���Ă���
    Dim flag As Boolean
    Dim Is_Found As Boolean: Is_Found = False
    Dim i As Long
    
    If Not Not MI Then  '�z�񂪋󂶂�Ȃ��Ƃ�
    
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

'HTML�̎擾
Public Function Get_HTML(ByVal Path As String) As String
    On Error Resume Next
    Dim html As String
    Dim http As Object
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", Path, False
    http.Send
    
    ' �_�E�����[�h�҂�
    Do While http.readyState <> 4
        DoEvents
    Loop
    
    ' ���ʏo��
    html = http.responseText
    
    Get_HTML = html
    
    
    '���
    Set http = Nothing
    
    
End Function

'Shift-JIS�ɂ��ăG���R�[�h
Public Function URLEnc_SHIFT_JIS(ByVal txt As String) As String
    Dim bin() As Byte
    Dim i As Integer
    Dim res As String
    bin = StrConv(txt, vbFromUnicode) 'Shift-Jis�ɂ��ăo�C�i���z���
    '�o�C�i���z��𕶎���ɕύX
    For i = 0 To UBound(bin)
        Select Case bin(i)
            Case &H31 To &H39, &H41 To &H5A, &H61 To &H7A '0-9,A-Z,a-z �Ȃ�
            res = res & Chr(bin(i)) '���̂܂܂̕���
        Case Else '����ȊO�̃G���R�[�h�Ώە����Ȃ�
            res = res & "%" & Hex(bin(i)) '�G���R�[�h
            'res = res & "%" & LCase(Hex(bin(i))) '�������ɂ���ꍇ
        End Select
    Next
    
    URLEnc_SHIFT_JIS = res
End Function


Public Function EncodeURL_MusicTitle(ByVal MusicTitle As String) As String
    Dim EncodedTitle As String  '�G���R�[�h���ꂽ�Ȗ�
    EncodedTitle = URLEnc_SHIFT_JIS(MusicTitle)
    
    Dim SearchURL As String '�̎��T�C�g�̌������ʂ�URL
    SearchURL = "https://www.utamap.com/searchkasi.php?searchname=title&word=" & EncodedTitle & "&act=search&search_by_keyword=%8C%9F%26%23160%3B%26%23160%3B%26%23160%3B%8D%F5&sortname=1&pattern=4"
    
    EncodeURL_MusicTitle = SearchURL
    
End Function





'*********************************************************************************�̎�********************************************************************************************

'�C���f�b�N�X����ŏI�I�ȉ̎����擾
Public Function Get_Lyrics_FromIndex(ByVal LyricsIndex As String) As String
    On Error Resume Next
    Dim Path As String: Path = ("https://www.utamap.com/phpflash/flashfalsephp.php?unum=" & LyricsIndex)
    Dim html As String
    Dim http As Object
    
    Dim Enc As Object
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", Path, False
    http.Send
    
    ' �_�E�����[�h�҂�
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
    
    html = Enc.ReadText '�ϊ��ς݂̃e�L�X�g�����ʂƂ���
    Enc.Close
    
    Dim index As Integer
    index = InStr(html, "=") + 1
    index = InStr(index, html, "=") + 1
    html = Mid(html, index, Len(html) - index + 1)
    
    ' ���ʏo��
    Get_Lyrics_FromIndex = html
    
    
    '���
    Set http = Nothing
    Set Enc = Nothing
    
End Function

'*****************************************************************�̎�E�쎌�E��ȁE�̂������E�^�C�A�b�v���**********************************************************************
'�ȏ����ꊇ�擾
Public Function Get_Info(ByVal MusicTitle As String) As MusicInfomation()
    
    'HTML���擾
    Dim utamapHTML As String: utamapHTML = Get_HTML(EncodeURL_MusicTitle(MusicTitle))
    
    Dim Info() As MusicInfomation           '���ɉ̎��A�Ȗ��A�̎�A�쎌�A��ȁA�̂������A�^�C�A�b�v���
    Dim Counter As Long: Counter = 0        '�J�E���^
    Dim HitNum As Long                      '�q�b�g�����Ȑ��̑���
    
    Dim sIndex As Long  '�n�܂�
    Dim eIndex As Long  '�I���
    
    
    '�܂����߂Ƀq�b�g�������擾����
    sIndex = InStr(1, utamapHTML, "�����L�[���[�h:") + Len("�����L�[���[�h:") + Len(MusicTitle) + Len(" &nbsp; �S ")
    eIndex = InStr(sIndex, utamapHTML, "��") - 1
    HitNum = CLng(Mid(utamapHTML, sIndex, eIndex - sIndex))
    If HitNum = 0 Then HitNum = -1
    
    sIndex = 1
    eIndex = 1
    
    Do While (Counter <= HitNum)
    
        ReDim Preserve Info(Counter) As MusicInfomation
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct160>") + Len("<TD class=ct160>")
        
        '�̎����擾����
        '�Ȃ�ID���擾
        Dim MusicID As String
        sIndex = InStr(sIndex, utamapHTML, "./showkasi.php?surl=") + Len("./showkasi.php?surl=")
        eIndex = InStr(sIndex, utamapHTML, ">") - 1
        MusicID = Mid(utamapHTML, sIndex, eIndex - sIndex)
        Info(Counter).Lyrics = Get_Lyrics_FromIndex(MusicID)
        
        
        '�^�C�g�����擾����
        sIndex = InStr(sIndex, utamapHTML, ">") + Len(">")
        eIndex = InStr(sIndex, utamapHTML, "</A></td>")
        Info(Counter).Title = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '�J�n�_�̍X�V
        sIndex = eIndex - 1
        
        '���炩�ɕ���������������Ƃ��̓f�[�^���擾�ł��Ă��Ȃ��Ƃ��ăG���[�Ƃ���(�������`�F�b�N���Ă����΍ŏ��Ŕ��f�ł���)
        If Len(Info(Counter).Title) > 50 Then
            Info(Counter).Lyrics = "_ERROR"
            Info(Counter).Title = "_ERROR"
            Info(Counter).Singer = "_ERROR"
            Info(Counter).SongWriter = "_ERROR"
            Info(Counter).Composer = "_ERROR"
            Info(Counter).FirstLine = "_ERROR"
            Info(Counter).TieupInfo = "_ERROR"
        End If
        
        
        '�̎���擾����
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct120>") + Len("<TD class=ct120>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).Singer = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '�J�n�_�̍X�V
        sIndex = eIndex - 1
        
        
        '�쎌�҂��擾����
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct120>") + Len("<TD class=ct120>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).SongWriter = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '�J�n�_�̍X�V
        sIndex = eIndex - 1
        
        
        '��Ȏ҂��擾����
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct120>") + Len("<TD class=ct120>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).Composer = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '�J�n�_�̍X�V
        sIndex = eIndex - 1
        
        '�̂��������擾����
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct160>") + Len("<TD class=ct160>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).FirstLine = Mid(utamapHTML, sIndex, eIndex - sIndex)
        '�J�n�_�̍X�V
        sIndex = eIndex - 1
        
        '�^�C�A�b�v�����擾����
        sIndex = InStr(sIndex, utamapHTML, "<TD class=ct>") + Len("<TD class=ct>")
        eIndex = InStr(sIndex, utamapHTML, "</td>")
        Info(Counter).TieupInfo = Mid(utamapHTML, sIndex, eIndex - sIndex)
        If Info(Counter).TieupInfo = "" Or Info(Counter).TieupInfo = "-" Then Info(Counter).TieupInfo = "�Ȃ�"
        If Len(Info(Counter).TieupInfo) > 50 Then Info(Counter).TieupInfo = "�Ȃ�"
        sIndex = eIndex - 1
        
        Counter = Counter + 1
    
    Loop
    
    Get_Info = Info
    
End Function

'*********************************************************************************************************************************************************************************



