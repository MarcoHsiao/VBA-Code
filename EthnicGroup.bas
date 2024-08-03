Attribute VB_Name = "Module1"
Option Explicit
Sub CalculatePercentagesForMultipleSheets()
Attribute CalculatePercentagesForMultipleSheets.VB_Description = "��X�ڸs"
Attribute CalculatePercentagesForMultipleSheets.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim ws As Worksheet
    Dim col As Integer
    Dim firstValue As Double
    Dim lastValue As Double
    Dim lastRow As Long
    Dim percentageDifference As Double
    Dim result As String
    Dim sheetsArray As Variant
    Dim sheetName As Variant

    ' �u�@��W�٦C��
    sheetsArray = Array("��1", "��2", "��3")
    
    ' �`���M���C�Ӥu�@��
    For Each sheetName In sheetsArray
        ' �]�w��e�u�@��
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' �`���M�� B �� K ��
        For col = 2 To 11 ' B �C�O 2�AK �C�O 11
            ' ����e�C���̫�@�ӫD�Ŧ�
            lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
            
            ' �T�O��e�C���ƭ�
            If lastRow >= 3 Then ' �ܤ֦��� 3 �檺�ƭ�
                ' �����e�C���Ĥ@�Ӽƭȡ]�� 3 ��^
                firstValue = ws.Cells(3, col).Value
                
                ' �����e�C���̫�@�Ӽƭ�
                lastValue = ws.Cells(lastRow, col).Value
                
                ' �p��ʤ���t��
                If firstValue <> 0 Then
                    percentageDifference = ((lastValue - firstValue) / firstValue) * 100
                    result = Format(percentageDifference, "0.00") & "%"
                Else
                    result = "�L�k�p��ʤ���" ' �p�G�Ĥ@�ӼƭȬ� 0
                End If
                
                ' ��ܵ��G�b�� 27 ��
                ws.Cells(27, col).Value = result
            Else
                ' �p�G��e�C�S���������ƭ�
                ws.Cells(27, col).Value = "�ƾڤ���"
            End If
        Next col
    Next sheetName

Sheets("��ڸs").Cells(3, "B").Value = Sheets("��1").Cells(27, "B").Value
Sheets("��ڸs").Cells(3, "C").Value = Sheets("��2").Cells(27, "B").Value
Sheets("��ڸs").Cells(3, "D").Value = Sheets("��3").Cells(27, "B").Value
Sheets("��ڸs").Cells(3, "E").Value = Range("B3") + Range("C3") + Range("D3")
Sheets("��ڸs").Cells(4, "B").Value = Sheets("��1").Cells(27, "C").Value
Sheets("��ڸs").Cells(4, "C").Value = Sheets("��2").Cells(27, "C").Value
Sheets("��ڸs").Cells(4, "D").Value = Sheets("��3").Cells(27, "C").Value
Sheets("��ڸs").Cells(4, "E").Value = Range("B4") + Range("C4") + Range("D4")
Sheets("��ڸs").Cells(5, "B").Value = Sheets("��1").Cells(27, "D").Value
Sheets("��ڸs").Cells(5, "C").Value = Sheets("��2").Cells(27, "D").Value
Sheets("��ڸs").Cells(5, "D").Value = Sheets("��3").Cells(27, "D").Value
Sheets("��ڸs").Cells(5, "E").Value = Range("B5") + Range("C5") + Range("D5")
Sheets("��ڸs").Cells(6, "B").Value = Sheets("��1").Cells(27, "E").Value
Sheets("��ڸs").Cells(6, "C").Value = Sheets("��2").Cells(27, "E").Value
Sheets("��ڸs").Cells(6, "D").Value = Sheets("��3").Cells(27, "E").Value
Sheets("��ڸs").Cells(6, "E").Value = Range("B6") + Range("C6") + Range("D6")
Sheets("��ڸs").Cells(7, "B").Value = Sheets("��1").Cells(27, "F").Value
Sheets("��ڸs").Cells(7, "C").Value = Sheets("��2").Cells(27, "F").Value
Sheets("��ڸs").Cells(7, "D").Value = Sheets("��3").Cells(27, "F").Value
Sheets("��ڸs").Cells(7, "E").Value = Range("B7") + Range("C7") + Range("D7")
Sheets("��ڸs").Cells(8, "B").Value = Sheets("��1").Cells(27, "G").Value
Sheets("��ڸs").Cells(8, "C").Value = Sheets("��2").Cells(27, "G").Value
Sheets("��ڸs").Cells(8, "D").Value = Sheets("��3").Cells(27, "G").Value
Sheets("��ڸs").Cells(8, "E").Value = Range("B8") + Range("C8") + Range("D8")
Sheets("��ڸs").Cells(9, "B").Value = Sheets("��1").Cells(27, "H").Value
Sheets("��ڸs").Cells(9, "C").Value = Sheets("��2").Cells(27, "H").Value
Sheets("��ڸs").Cells(9, "D").Value = Sheets("��3").Cells(27, "H").Value
Sheets("��ڸs").Cells(9, "E").Value = Range("B9") + Range("C9") + Range("D9")
Sheets("��ڸs").Cells(10, "B").Value = Sheets("��1").Cells(27, "I").Value
Sheets("��ڸs").Cells(10, "C").Value = Sheets("��2").Cells(27, "I").Value
Sheets("��ڸs").Cells(10, "D").Value = Sheets("��3").Cells(27, "I").Value
Sheets("��ڸs").Cells(10, "E").Value = Range("B10") + Range("C10") + Range("D10")
Sheets("��ڸs").Cells(11, "B").Value = Sheets("��1").Cells(27, "J").Value
Sheets("��ڸs").Cells(11, "C").Value = Sheets("��2").Cells(27, "J").Value
Sheets("��ڸs").Cells(11, "D").Value = Sheets("��3").Cells(27, "J").Value
Sheets("��ڸs").Cells(11, "E").Value = Range("B11") + Range("C11") + Range("D11")
Sheets("��ڸs").Cells(12, "B").Value = Sheets("��1").Cells(27, "K").Value
Sheets("��ڸs").Cells(12, "C").Value = Sheets("��2").Cells(27, "K").Value
Sheets("��ڸs").Cells(12, "D").Value = Sheets("��3").Cells(27, "K").Value
Sheets("��ڸs").Cells(12, "E").Value = Range("B12") + Range("C12") + Range("D12")



End Sub

