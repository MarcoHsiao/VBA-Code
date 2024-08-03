Attribute VB_Name = "Module2"
Option Explicit

Sub CalculateAndSortWithErrorsAtBottom()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim uValue As Variant, vValue As Variant, wValue As Variant, xValue As Variant, yValue As Variant
    Dim oValue As Double, sValue As Double, kValue As Double
    
    ' �]�w�u�@��
    Set ws = ThisWorkbook.Sheets("��Ӫ�")
    
    ' ��� U �檺�̫�@�ӫD�Ŧ�
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
    
    ' �q�� 2 ��}�l�p��
    For i = 2 To lastRow
        On Error Resume Next ' ��e�榳���~���~��B�z�U�@��
        
        ' ��������ƭ�
        oValue = ws.Cells(i, "O").Value
        sValue = ws.Cells(i, "S").Value
        kValue = ws.Cells(i, "K").Value
        
        ' �p�� U �檺�ƭȨî榡��
        If sValue <> 0 Then
            uValue = (oValue - sValue) / sValue
            ws.Cells(i, "U").Value = Format(uValue, "0.00")
        Else
            ws.Cells(i, "U").Value = "�p����~"
        End If
        
        ' �p�� V �檺�ƭȨî榡��
        If oValue <> 0 Then
            vValue = (kValue - oValue) / oValue
            ws.Cells(i, "V").Value = Format(vValue, "0.00")
        Else
            ws.Cells(i, "V").Value = "�p����~"
        End If
        
        ' �p�� W �檺�ƭȨî榡��
        If IsNumeric(ws.Cells(i, "U").Value) And IsNumeric(ws.Cells(i, "V").Value) Then
            wValue = CDbl(ws.Cells(i, "U").Value) + CDbl(ws.Cells(i, "V").Value)
            ws.Cells(i, "W").Value = Format(wValue, "0.00")
        Else
            ws.Cells(i, "W").Value = "�p����~"
        End If
        
        ' �p�� X �檺�ƭȨî榡��
        If IsNumeric(ws.Cells(i, "U").Value) And IsNumeric(ws.Cells(i, "V").Value) Then
            xValue = CDbl(ws.Cells(i, "V").Value) - CDbl(ws.Cells(i, "U").Value)
            ws.Cells(i, "X").Value = Format(xValue, "0.00")
        Else
            ws.Cells(i, "X").Value = "�p����~"
        End If
        
        ' �p�� Y �檺�ƭȨî榡��
        If IsNumeric(ws.Cells(i, "U").Value) And IsNumeric(ws.Cells(i, "V").Value) And IsNumeric(ws.Cells(i, "W").Value) And IsNumeric(ws.Cells(i, "X").Value) Then
            yValue = CDbl(ws.Cells(i, "U").Value) + CDbl(ws.Cells(i, "V").Value) + CDbl(ws.Cells(i, "W").Value) + CDbl(ws.Cells(i, "X").Value)
            ws.Cells(i, "Y").Value = Format(yValue, "0.00")
        Else
            ws.Cells(i, "Y").Value = "�p����~"
        End If
        
        On Error GoTo 0 ' ���m���~�B�z
    Next i
    
    ' �N�p����~����ƨ�̫�
    ws.Range("A1:Y" & lastRow).Sort Key1:=ws.Range("Y1"), Order1:=xlAscending, Header:=xlYes
    
    ' �A���Ƨ� Y ��]�Ѥj��p�^
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("Y2:Y" & lastRow), Order:=xlDescending
        .SetRange ws.Range("A1:Y" & lastRow) ' ���] A �榳���D�A�ھڹ�ڽd��վ�
        .Header = xlYes ' �����D
        .Apply
    End With
    
Sheets("��ڸs").Cells(3, "G").Value = Sheets("��Ӫ�").Cells(2, "A").Value
Sheets("��ڸs").Cells(4, "G").Value = Sheets("��Ӫ�").Cells(3, "A").Value
Sheets("��ڸs").Cells(5, "G").Value = Sheets("��Ӫ�").Cells(4, "A").Value
Sheets("��ڸs").Cells(6, "G").Value = Sheets("��Ӫ�").Cells(5, "A").Value
Sheets("��ڸs").Cells(7, "G").Value = Sheets("��Ӫ�").Cells(6, "A").Value
Sheets("��ڸs").Cells(8, "G").Value = Sheets("��Ӫ�").Cells(7, "A").Value
Sheets("��ڸs").Cells(9, "G").Value = Sheets("��Ӫ�").Cells(8, "A").Value
Sheets("��ڸs").Cells(10, "G").Value = Sheets("��Ӫ�").Cells(9, "A").Value
Sheets("��ڸs").Cells(11, "G").Value = Sheets("��Ӫ�").Cells(10, "A").Value
Sheets("��ڸs").Cells(12, "G").Value = Sheets("��Ӫ�").Cells(11, "A").Value

End Sub

