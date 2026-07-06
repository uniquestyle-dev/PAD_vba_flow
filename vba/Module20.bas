Attribute VB_Name = "Module20"
Option Explicit

Public Sub STORES�ϊ�()
    Dim inputFile As String
    Dim outputFolder As String
    Dim wbInput As Workbook
    Dim ws As Worksheet
    Dim colItem As Long, colZip As Long
    Dim colSei As Long, colMei As Long
    Dim colPref As Long, colAddr As Long
    Dim colOrderNumber As Long, colOrderDate As Long
    Dim lastRow As Long
    Dim dict As Object
    Dim cell As Range
    Dim item As Variant
    Dim fso As Object, ts As Object, tsSummary As Object
    Dim fields As Variant
    Dim lineParts As Variant
    Dim i As Long
    Dim headerVal As String
    Dim col As Long
    Dim category As String
    Dim itemName As String

    '������ ����CSV�t�@�C���������擾�i�쐬�������ŐV��20*.csv�j������
    Dim dlFolder As Object, dlFile As Object, latestDate As Date
    Dim dlPath As String
    dlPath = "C:\Users\lenovo\Desktop\�_�E�����[�h\"
    latestDate = 0
    inputFile = ""

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dlFolder = fso.GetFolder(dlPath)
    For Each dlFile In dlFolder.Files
        If LCase(dlFile.Name) Like "20*.csv" Then
            If dlFile.DateCreated > latestDate Then
                latestDate = dlFile.DateCreated
                inputFile = dlFile.Path
            End If
        End If
    Next dlFile

    If inputFile = "" Then
        Exit Sub
    End If

    Debug.Print "�t�@�C��: " & inputFile

    ' �o�͐�t�H���_
    outputFolder = "C:\Users\lenovo\Desktop\�_�E�����[�h\"

    ' CSV��UTF-8�œǂݍ���
    Set wbInput = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbInput.Sheets(1)
    
    Dim qt As QueryTable
    Set qt = ws.QueryTables.Add("TEXT;" & inputFile, ws.Range("A1"))
    With qt
        .TextFilePlatform = 932
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        
        Dim tArr(0 To 255) As Variant
        For i = 0 To 255: tArr(i) = xlTextFormat: Next i
        .TextFileColumnDataTypes = tArr
        
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    ' �w�b�_�s��������������
    colItem = 0: colZip = 0: colSei = 0: colMei = 0
    colPref = 0: colAddr = 0: colOrderNumber = 0: colOrderDate = 0
    
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headerVal = ws.Cells(1, col).Value
        Select Case headerVal
            Case "�A�C�e����":                      colItem = col
            Case "�X�֔ԍ�(�w����)", "�X�֔ԍ�(�z����)":  colZip = col
            Case "��(�w����)", "��(�z����)":           colSei = col
            Case "��(�w����)", "��(�z����)":           colMei = col
            Case "�s���{��(�w����)", "�s���{��(�z����)":  colPref = col
            Case "�Z��(�w����)", "�Z��(�z����)":        colAddr = col
            Case "�I�[�_�[�ԍ�":                      colOrderNumber = col
            Case "�I�[�_�[����", "�I�[�_�[��":          colOrderDate = col
        End Select
    Next col

    ' �f�[�^�ŏI�s
    lastRow = ws.Cells(ws.Rows.Count, colItem).End(xlUp).row

    If lastRow <= 1 Then
        wbInput.Close SaveChanges:=False
        Debug.Print "STORES ��M�f�[�^0�� - �X�L�b�v"
        Exit Sub
    End If

    '���� 1.csv�i�T�}���j���o�� ����
    Set tsSummary = fso.CreateTextFile(outputFolder & "\1.csv", True, False)
    tsSummary.WriteLine Join(Array( _
        "�I�[�_�[�ԍ�", "�I�[�_�[��", _
        "��(�z����)", "��(�z����)", _
        "�z�����@", "�����\�����", "�₢���킹�ԍ�", "���l", _
        "��������" _
    ), ",")
    For i = 2 To lastRow
        tsSummary.WriteLine _
            ws.Cells(i, colOrderNumber).Value & "," & _
            ws.Cells(i, colOrderDate).Value & "," & _
            ws.Cells(i, colSei).Value & "," & _
            ws.Cells(i, colMei).Value & "," & _
            "���{�X��" & "," & _
            "" & "," & _
            "" & "," & _
            "" & "," & _
            "1"
    Next i
    tsSummary.Close

    '���� �J�e�S�����ƂɎ��W ����
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        itemName = Trim(ws.Cells(i, colItem).Value)
        category = GetStoresCategory(itemName)
        If Not dict.Exists(category) Then dict.Add category, Nothing
    Next i

    ' �o�͗p�w�b�_�z��
    fields = Array( _
        "���͂���X�֔ԍ�", "���͂��掁��", "���͂���h��", _
        "���͂���Z��1�s��", "���͂���Z��2�s��", _
        "���͂���Z��3�s��", "���͂���Z��4�s��", "���e�i" _
    )

    ' �J�e�S�����ƂɃt�@�C���o��
    For Each item In dict.Keys
        Set ts = fso.CreateTextFile(outputFolder & "\" & item & ".csv", True, False)
        ts.WriteLine Join(fields, vbTab)
        
        For i = 2 To lastRow
            itemName = Trim(ws.Cells(i, colItem).Value)
            category = GetStoresCategory(itemName)
            
            If category = item Then
                lineParts = Array( _
                    Format(ws.Cells(i, colZip).Value, "0000000"), _
                    SanitizeClickPost(ws.Cells(i, colSei).Value & ws.Cells(i, colMei).Value), _
                    "�l", _
                    SanitizeClickPost(ws.Cells(i, colPref).Value), _
                    SanitizeClickPost(ws.Cells(i, colAddr).Value), _
                    "", "", "����" _
                )
                ts.WriteLine Join(lineParts, vbTab)
            End If
        Next i
        
        ts.Close
        Debug.Print "�o��: " & item & ".csv"
    Next item

    wbInput.Close SaveChanges:=False
   
    ' 1.csv ���J��
    Workbooks.Open fileName:=outputFolder & "\1.csv"
    
    Debug.Print "===== ���� ====="
End Sub


'----------------------------------------------
' �A�C�e���� �� �J�e�S�����}�b�s���O
'----------------------------------------------
Private Function GetStoresCategory(ByVal itemName As String) As String
    Select Case itemName
        Case "�d�q���Ј���T�[�r�X_c"
            GetStoresCategory = "carboff_stores"
        Case "�d�q���Ј���T�[�r�X_s", "�d�q���Ј���T�[�r�X_sr"
            GetStoresCategory = "salasiru_stores"
        Case "�d�q���Ј���T�[�r�X_m", "�d�q���Ј���T�[�r�X_mr"
            GetStoresCategory = "Medico_stores"
        Case "�d�q���Ј���T�[�r�X_i"
            GetStoresCategory = "ice_stores"
        Case "�d�q���Ј���T�[�r�X_sw"
            GetStoresCategory = "sweets_stores"
        Case "�d�q���Ј���T�[�r�X_set"
            GetStoresCategory = "set_stores"
        Case Else
            GetStoresCategory = "other_stores_" & itemName
    End Select
End Function


'----------------------------------------------
' ClickPost�捞�G���[�h�~: , �� " ������
'----------------------------------------------
Private Function SanitizeClickPost(ByVal txt As String) As String
    SanitizeClickPost = Replace(Replace(txt, ",", ""), """", "")
End Function


