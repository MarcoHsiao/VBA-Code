Attribute VB_Name = "Module1"
Option Explicit
Sub CalculatePercentagesForMultipleSheets()
Attribute CalculatePercentagesForMultipleSheets.VB_Description = "選出族群"
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

    ' 工作表名稱列表
    sheetsArray = Array("月1", "月2", "月3")
    
    ' 循環遍歷每個工作表
    For Each sheetName In sheetsArray
        ' 設定當前工作表
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' 循環遍歷 B 到 K 欄
        For col = 2 To 11 ' B 列是 2，K 列是 11
            ' 找到當前列的最後一個非空行
            lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
            
            ' 確保當前列有數值
            If lastRow >= 3 Then ' 至少有第 3 行的數值
                ' 獲取當前列的第一個數值（第 3 行）
                firstValue = ws.Cells(3, col).Value
                
                ' 獲取當前列的最後一個數值
                lastValue = ws.Cells(lastRow, col).Value
                
                ' 計算百分比差值
                If firstValue <> 0 Then
                    percentageDifference = ((lastValue - firstValue) / firstValue) * 100
                    result = Format(percentageDifference, "0.00") & "%"
                Else
                    result = "無法計算百分比" ' 如果第一個數值為 0
                End If
                
                ' 顯示結果在第 27 行
                ws.Cells(27, col).Value = result
            Else
                ' 如果當前列沒有足夠的數值
                ws.Cells(27, col).Value = "數據不足"
            End If
        Next col
    Next sheetName

Sheets("選族群").Cells(3, "B").Value = Sheets("月1").Cells(27, "B").Value
Sheets("選族群").Cells(3, "C").Value = Sheets("月2").Cells(27, "B").Value
Sheets("選族群").Cells(3, "D").Value = Sheets("月3").Cells(27, "B").Value
Sheets("選族群").Cells(3, "E").Value = Range("B3") + Range("C3") + Range("D3")
Sheets("選族群").Cells(4, "B").Value = Sheets("月1").Cells(27, "C").Value
Sheets("選族群").Cells(4, "C").Value = Sheets("月2").Cells(27, "C").Value
Sheets("選族群").Cells(4, "D").Value = Sheets("月3").Cells(27, "C").Value
Sheets("選族群").Cells(4, "E").Value = Range("B4") + Range("C4") + Range("D4")
Sheets("選族群").Cells(5, "B").Value = Sheets("月1").Cells(27, "D").Value
Sheets("選族群").Cells(5, "C").Value = Sheets("月2").Cells(27, "D").Value
Sheets("選族群").Cells(5, "D").Value = Sheets("月3").Cells(27, "D").Value
Sheets("選族群").Cells(5, "E").Value = Range("B5") + Range("C5") + Range("D5")
Sheets("選族群").Cells(6, "B").Value = Sheets("月1").Cells(27, "E").Value
Sheets("選族群").Cells(6, "C").Value = Sheets("月2").Cells(27, "E").Value
Sheets("選族群").Cells(6, "D").Value = Sheets("月3").Cells(27, "E").Value
Sheets("選族群").Cells(6, "E").Value = Range("B6") + Range("C6") + Range("D6")
Sheets("選族群").Cells(7, "B").Value = Sheets("月1").Cells(27, "F").Value
Sheets("選族群").Cells(7, "C").Value = Sheets("月2").Cells(27, "F").Value
Sheets("選族群").Cells(7, "D").Value = Sheets("月3").Cells(27, "F").Value
Sheets("選族群").Cells(7, "E").Value = Range("B7") + Range("C7") + Range("D7")
Sheets("選族群").Cells(8, "B").Value = Sheets("月1").Cells(27, "G").Value
Sheets("選族群").Cells(8, "C").Value = Sheets("月2").Cells(27, "G").Value
Sheets("選族群").Cells(8, "D").Value = Sheets("月3").Cells(27, "G").Value
Sheets("選族群").Cells(8, "E").Value = Range("B8") + Range("C8") + Range("D8")
Sheets("選族群").Cells(9, "B").Value = Sheets("月1").Cells(27, "H").Value
Sheets("選族群").Cells(9, "C").Value = Sheets("月2").Cells(27, "H").Value
Sheets("選族群").Cells(9, "D").Value = Sheets("月3").Cells(27, "H").Value
Sheets("選族群").Cells(9, "E").Value = Range("B9") + Range("C9") + Range("D9")
Sheets("選族群").Cells(10, "B").Value = Sheets("月1").Cells(27, "I").Value
Sheets("選族群").Cells(10, "C").Value = Sheets("月2").Cells(27, "I").Value
Sheets("選族群").Cells(10, "D").Value = Sheets("月3").Cells(27, "I").Value
Sheets("選族群").Cells(10, "E").Value = Range("B10") + Range("C10") + Range("D10")
Sheets("選族群").Cells(11, "B").Value = Sheets("月1").Cells(27, "J").Value
Sheets("選族群").Cells(11, "C").Value = Sheets("月2").Cells(27, "J").Value
Sheets("選族群").Cells(11, "D").Value = Sheets("月3").Cells(27, "J").Value
Sheets("選族群").Cells(11, "E").Value = Range("B11") + Range("C11") + Range("D11")
Sheets("選族群").Cells(12, "B").Value = Sheets("月1").Cells(27, "K").Value
Sheets("選族群").Cells(12, "C").Value = Sheets("月2").Cells(27, "K").Value
Sheets("選族群").Cells(12, "D").Value = Sheets("月3").Cells(27, "K").Value
Sheets("選族群").Cells(12, "E").Value = Range("B12") + Range("C12") + Range("D12")



End Sub

