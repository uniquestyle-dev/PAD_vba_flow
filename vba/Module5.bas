Attribute VB_Name = "Module5"
Option Explicit

Public Sub STORES�ϊ�()
    Dim inputFile As String
    Dim outputFolder As String
    Dim wbInput As Workbook
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim colItem As Long, colZip As Long
    Dim colSei As Long, colMei As Long
    Dim colPref As Long, colAddr As Long
    Dim colOrderNumber As Long, colOrderDate As Long
    Dim lastRow As Long
    Dim dict As Object
    Dim cell As Range
    Dim item As Variant
    Dim idx As Long
    Dim fso As Object, ts As Object, tsSummary As Object
    Dim fields As Variant
    Dim lineParts As Variant
    Dim i As Long

    ' ����CSV�t�@�C����I��
    inputFile = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "����CSV�t�@�C����I��")
    If inputFile = "False" Then Exit Sub

    ' �o�͐�t�H���_���f�X�N�g�b�v�ɌŒ�
    outputFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    ' CSV��Shift-JIS�œǂݍ���
    Workbooks.OpenText _
        fileName:=inputFile, _
        Origin:=932, _
        DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        Comma:=True

    Set wbInput = ActiveWorkbook
    Set ws = wbInput.Sheets(1)

    ' �w�b�_�s
    Set headerRow = ws.Rows(1)
    colItem = Application.Match("�A�C�e����", headerRow, 0)
    colZip = Application.Match("�X�֔ԍ�(�z����)", headerRow, 0)
    colSei = Application.Match("��(�z����)", headerRow, 0)
    colMei = Application.Match("��(�z����)", headerRow, 0)
    colPref = Application.Match("�s���{��(�z����)", headerRow, 0)
    colAddr = Application.Match("�Z��(�z����)", headerRow, 0)
    colOrderNumber = Application.Match("�I�[�_�[�ԍ�", headerRow, 0)
    colOrderDate = Application.Match("�I�[�_�[��", headerRow, 0)

    ' �f�[�^�ŏI�s
    lastRow = ws.Cells(ws.Rows.Count, colItem).End(xlUp).row

    If lastRow <= 1 Then
        wbInput.Close SaveChanges:=False
        Debug.Print "STORES ��M�f�[�^0�� - �X�L�b�v"
        Exit Sub
    End If

    ' FileSystemObject ����
    Set fso = CreateObject("Scripting.FileSystemObject")

    '���� 1.csv�i�T�}���j���o�� ����
    Set tsSummary = fso.CreateTextFile(outputFolder & "\1.csv", True, False)
    ' �w�b�_�s�i�J���}��؂�j
    tsSummary.WriteLine Join(Array( _
        "�I�[�_�[�ԍ�", "�I�[�_�[��", _
        "��(�z����)", "��(�z����)", _
        "�z�����@", "�����\�����", "�₢���킹�ԍ�", "���l", _
        "��������" _
    ), ",")
    ' �f�[�^�s�i�J���}��؂�j
    For i = 2 To lastRow
        tsSummary.WriteLine _
            ws.Cells(i, colOrderNumber).Value & "," & _
            Format(ws.Cells(i, colOrderDate).Value, "yyyy/m/d h:mm") & "," & _
            ws.Cells(i, colSei).Value & "," & _
            ws.Cells(i, colMei).Value & "," & _
            "���{�X��" & "," & _
            "" & "," & _
            "" & "," & _
            "" & "," & _
            "1"
    Next i
    tsSummary.Close
    '������������������������������������������������������������������������

    ' �A�C�e�����Ń\�[�g�i�C�Ӂj
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(2, colItem), ws.Cells(lastRow, colItem)), _
        SortOn:=xlSortOnValues, Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ws.Columns.Count).End(xlToLeft))
        .Header = xlYes
        .Apply
    End With

    ' �A�C�e�������f�B�N�V���i����
    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In ws.Range(ws.Cells(2, colItem), ws.Cells(lastRow, colItem))
        If Not dict.Exists(cell.Value) Then dict.Add cell.Value, Empty
    Next cell

    ' �o�͗p�w�b�_�z��
    fields = Array( _
        "���͂���X�֔ԍ�", "���͂��掁��", "���͂���h��", _
        "���͂���Z��1�s��", "���͂���Z��2�s��", _
        "���͂���Z��3�s��", "���͂���Z��4�s��", "���e�i" _
    )

    idx = 1
    ' �e�A�C�e�����ƂɃe�L�X�g�t�@�C���𐶐�
    For Each item In dict.Keys
        Set ts = fso.CreateTextFile(outputFolder & "\Book" & idx & ".csv", True, False)
        ' �w�b�_��������
        ts.WriteLine Join(fields, vbTab)
        ' �f�[�^�s��������
        For Each cell In ws.Range(ws.Cells(2, colItem), ws.Cells(lastRow, colItem))
            If cell.Value = item Then
                lineParts = Array( _
                    ws.Cells(cell.row, colZip).Value, _
                    ws.Cells(cell.row, colSei).Value & ws.Cells(cell.row, colMei).Value, _
                    "�l", _
                    ws.Cells(cell.row, colPref).Value, _
                    ws.Cells(cell.row, colAddr).Value, _
                    "", "", "����" _
                )
                ts.WriteLine Join(lineParts, vbTab)
            End If
        Next cell
        ts.Close
        idx = idx + 1
    Next item

    wbInput.Close SaveChanges:=False
   
    ' 1.csv ���J��
    Workbooks.Open fileName:=outputFolder & "\1.csv"
End Sub


