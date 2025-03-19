# excel_greper

##add excel vba
```
Sub SearchExcelFilesAndOutputResults()
    Dim folderPath As String
    Dim searchString As String
    Dim resultWb As Workbook
    Dim resultWs As Worksheet
    Dim resultRow As Long
    
    ' 検索するフォルダパスをセルから取得
    folderPath = ThisWorkbook.Sheets("Sheet1").Range("A1").Value
    
    ' 検索する文字列をセルから取得
    searchString = ThisWorkbook.Sheets("Sheet1").Range("B1").Value
    
    ' 結果を出力する新しいワークブックを作成
    Set resultWb = Workbooks.Add
    Set resultWs = resultWb.Sheets(1)
    
    ' ヘッダを設定
    resultWs.Cells(1, 1).Value = "ファイルパス"
    resultWs.Cells(1, 2).Value = "シート名"
    resultWs.Cells(1, 3).Value = "セルアドレス"
    resultWs.Cells(1, 4).Value = "内容"
    resultWs.Cells(1, 5).Value = "リンク"
    
    resultRow = 2 ' 結果の出力開始行
    
    ' 再帰的にフォルダを検索
    Call SearchFolder(folderPath, searchString, resultWs, resultRow)
    
    ' オブジェクトを解放
    Set resultWb = Nothing
    Set resultWs = Nothing
End Sub

Sub SearchFolder(folderPath As String, searchString As String, resultWs As Worksheet, ByRef resultRow As Long)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim found As Range
    Dim linkAddress As String
    
    ' ファイルシステムオブジェクトを作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' フォルダ内のすべてのファイルをループ
    For Each file In folder.Files
        ' Excelファイルのみを対象
        If fso.GetExtensionName(file.Name) = "xlsx" Or fso.GetExtensionName(file.Name) = "xls" Then
            ' ワークブックを開く
            Set wb = Workbooks.Open(file.Path)
            
            ' ワークシートをループ
            For Each ws In wb.Worksheets
                ' ワークシート内を検索
                Set found = ws.Cells.Find(What:=searchString, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                If Not found Is Nothing Then
                    Do
                        ' 結果を出力
                        resultWs.Cells(resultRow, 1).Value = file.Path
                        resultWs.Cells(resultRow, 2).Value = ws.Name
                        resultWs.Cells(resultRow, 3).Value = found.Address
                        resultWs.Cells(resultRow, 4).Value = found.Value
                        ' ハイパーリンクのアドレスを作成
                        linkAddress = "'" & file.Path & "'#" & ws.Name & "!" & found.Address
                        ' ハイパーリンクを設定
                        resultWs.Hyperlinks.Add Anchor:=resultWs.Cells(resultRow, 5), Address:="", SubAddress:=linkAddress, TextToDisplay:="リンク"
                        resultRow = resultRow + 1
                        Set found = ws.Cells.FindNext(found)
                    Loop While Not found Is Nothing And found.Address <> ws.Cells.Find(What:=searchString, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address
                End If
            Next ws
            
            ' ワークブックを閉じる
            wb.Close SaveChanges:=False
        End If
    Next file
    
    ' サブフォルダをループ
    For Each subFolder In folder.SubFolders
        ' 再帰的にサブフォルダを検索
        Call SearchFolder(subFolder.Path, searchString, resultWs, resultRow)
    Next subFolder
    
    ' オブジェクトを解放
    Set fso = Nothing
    Set folder = Nothing
End Sub