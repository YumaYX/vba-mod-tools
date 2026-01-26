'######### GetValueByID_Hash
'========================================
' ハッシュでIDから値取得（ID列・取得列は自動検索）
' ws           : 対象ワークシート
' idHeader     : ID列の見出し名
' idValue      : 検索するID
' targetHeader : 取得したい列の見出し名
' headerRow    : 見出し行番号（省略可、通常1）
' 戻り値       : 該当セルの値（見つからなければ""）
'========================================
Function GetValueByID_Hash(ws As Worksheet, _
                           idHeader As String, _
                           idValue As Variant, _
                           targetHeader As String, _
                           Optional headerRow As Long = 1) As Variant
    GetValueByID_Hash = "" ' 見出しが見つからない
    
    ' ID列と取得列の列番号を取得
    Dim idCol     As Long: idCol     = GetColumnByHeader(ws, idHeader, headerRow)
    Dim targetCol As Long: targetCol = GetColumnByHeader(ws, targetHeader, headerRow)
    If idCol = 0 Or targetCol = 0 Then Exit Function    

    ' 最終行
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    ' Dictionary作成
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long: For r = headerRow + 1 To lastRow
        If Not dict.Exists(ws.Cells(r, idCol).Value) Then dict(ws.Cells(r, idCol).Value) = ws.Cells(r, targetCol).Value
    Next r

    ' IDで値取得
    If dict.Exists(idValue) Then GetValueByID_Hash = dict(idValue)
End Function
