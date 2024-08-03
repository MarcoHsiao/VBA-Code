Attribute VB_Name = "Module2"
Option Explicit

Sub CalculateAndSortWithErrorsAtBottom()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim uValue As Variant, vValue As Variant, wValue As Variant, xValue As Variant, yValue As Variant
    Dim oValue As Double, sValue As Double, kValue As Double
    
    ' 設定工作表
    Set ws = ThisWorkbook.Sheets("選個股")
    
    ' 找到 U 欄的最後一個非空行
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
    
    ' 從第 2 行開始計算
    For i = 2 To lastRow
        On Error Resume Next ' 當前行有錯誤時繼續處理下一行
        
        ' 獲取相關數值
        oValue = ws.Cells(i, "O").Value
        sValue = ws.Cells(i, "S").Value
        kValue = ws.Cells(i, "K").Value
        
        ' 計算 U 欄的數值並格式化
        If sValue <> 0 Then
            uValue = (oValue - sValue) / sValue
            ws.Cells(i, "U").Value = Format(uValue, "0.00")
        Else
            ws.Cells(i, "U").Value = "計算錯誤"
        End If
        
        ' 計算 V 欄的數值並格式化
        If oValue <> 0 Then
            vValue = (kValue - oValue) / oValue
            ws.Cells(i, "V").Value = Format(vValue, "0.00")
        Else
            ws.Cells(i, "V").Value = "計算錯誤"
        End If
        
        ' 計算 W 欄的數值並格式化
        If IsNumeric(ws.Cells(i, "U").Value) And IsNumeric(ws.Cells(i, "V").Value) Then
            wValue = CDbl(ws.Cells(i, "U").Value) + CDbl(ws.Cells(i, "V").Value)
            ws.Cells(i, "W").Value = Format(wValue, "0.00")
        Else
            ws.Cells(i, "W").Value = "計算錯誤"
        End If
        
        ' 計算 X 欄的數值並格式化
        If IsNumeric(ws.Cells(i, "U").Value) And IsNumeric(ws.Cells(i, "V").Value) Then
            xValue = CDbl(ws.Cells(i, "V").Value) - CDbl(ws.Cells(i, "U").Value)
            ws.Cells(i, "X").Value = Format(xValue, "0.00")
        Else
            ws.Cells(i, "X").Value = "計算錯誤"
        End If
        
        ' 計算 Y 欄的數值並格式化
        If IsNumeric(ws.Cells(i, "U").Value) And IsNumeric(ws.Cells(i, "V").Value) And IsNumeric(ws.Cells(i, "W").Value) And IsNumeric(ws.Cells(i, "X").Value) Then
            yValue = CDbl(ws.Cells(i, "U").Value) + CDbl(ws.Cells(i, "V").Value) + CDbl(ws.Cells(i, "W").Value) + CDbl(ws.Cells(i, "X").Value)
            ws.Cells(i, "Y").Value = Format(yValue, "0.00")
        Else
            ws.Cells(i, "Y").Value = "計算錯誤"
        End If
        
        On Error GoTo 0 ' 重置錯誤處理
    Next i
    
    ' 將計算錯誤的行排到最後
    ws.Range("A1:Y" & lastRow).Sort Key1:=ws.Range("Y1"), Order1:=xlAscending, Header:=xlYes
    
    ' 再次排序 Y 欄（由大到小）
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("Y2:Y" & lastRow), Order:=xlDescending
        .SetRange ws.Range("A1:Y" & lastRow) ' 假設 A 欄有標題，根據實際範圍調整
        .Header = xlYes ' 有標題
        .Apply
    End With
    
Sheets("選族群").Cells(3, "G").Value = Sheets("選個股").Cells(2, "A").Value
Sheets("選族群").Cells(4, "G").Value = Sheets("選個股").Cells(3, "A").Value
Sheets("選族群").Cells(5, "G").Value = Sheets("選個股").Cells(4, "A").Value
Sheets("選族群").Cells(6, "G").Value = Sheets("選個股").Cells(5, "A").Value
Sheets("選族群").Cells(7, "G").Value = Sheets("選個股").Cells(6, "A").Value
Sheets("選族群").Cells(8, "G").Value = Sheets("選個股").Cells(7, "A").Value
Sheets("選族群").Cells(9, "G").Value = Sheets("選個股").Cells(8, "A").Value
Sheets("選族群").Cells(10, "G").Value = Sheets("選個股").Cells(9, "A").Value
Sheets("選族群").Cells(11, "G").Value = Sheets("選個股").Cells(10, "A").Value
Sheets("選族群").Cells(12, "G").Value = Sheets("選個股").Cells(11, "A").Value

End Sub

