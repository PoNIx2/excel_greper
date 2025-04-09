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

Sub SearchCamelCaseInExcelFiles()
    Dim FolderPath As String
    Dim CamelCasePattern As String
    Dim FileSystem As Object
    Dim Folder As Object
    Dim File As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim Cell As Range
    Dim ResultSheet As Worksheet
    Dim OutputRow As Long
    Dim SubFolder As Object
    
    ' フォルダパスと正規表現パターンをセルから取得
    FolderPath = ThisWorkbook.Sheets(1).Range("A1").Value ' シート1のセルA1からフォルダパスを取得
    CamelCasePattern = ThisWorkbook.Sheets(1).Range("B1").Value ' シート1のセルB1から正規表現パターンを取得
    
    If FolderPath = "" Or CamelCasePattern = "" Then
        MsgBox "フォルダパスまたは正規表現パターンが指定されていません。セル A1 と B1 を確認してください。", vbExclamation
        Exit Sub
    End If
    
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    
    ' 出力用シートを作成
    On Error Resume Next
    Set ResultSheet = ThisWorkbook.Worksheets("CamelCaseResults")
    If ResultSheet Is Nothing Then
        Set ResultSheet = ThisWorkbook.Worksheets.Add
        ResultSheet.Name = "CamelCaseResults"
    End If
    On Error GoTo 0
    
    ' 結果シートのヘッダー設定
    ResultSheet.Cells.Clear
    ResultSheet.Cells(1, 1).Value = "ファイル名"
    ResultSheet.Cells(1, 2).Value = "シート名"
    ResultSheet.Cells(1, 3).Value = "セルアドレス"
    ResultSheet.Cells(1, 4).Value = "セルの値"
    OutputRow = 2
    
    ' フォルダ内のファイルを検索
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set Folder = FileSystem.GetFolder(FolderPath)
    Call ProcessFolder(Folder, CamelCasePattern, ResultSheet, OutputRow)
    
    MsgBox "検索が完了しました。結果はシート 'CamelCaseResults' に出力されました。", vbInformation
End Sub

Sub ProcessFolder(Folder As Object, Pattern As String, ResultSheet As Worksheet, ByRef OutputRow As Long)
    Dim File As Object
    Dim SubFolder As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim Cell As Range
    Dim RegExp As Object
    
    ' 正規表現オブジェクトの作成
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Pattern = Pattern
    RegExp.IgnoreCase = False
    RegExp.Global = False
    
    ' フォルダ内の各ファイルをチェック
    For Each File In Folder.Files
        If InStr(File.Name, ".xls") > 0 And Not InStr(File.Name, "~$") > 0 Then ' 一時ファイルを除外
            On Error Resume Next
            Set wb = Workbooks.Open(File.Path, ReadOnly:=True)
            On Error GoTo 0
            
            If Not wb Is Nothing Then
                For Each ws In wb.Worksheets
                    For Each Cell In ws.UsedRange
                        If RegExp.Test(Cell.Value) Then
                            ' 結果を出力
                            ResultSheet.Cells(OutputRow, 1).Value = File.Name
                            ResultSheet.Cells(OutputRow, 2).Value = ws.Name
                            ResultSheet.Cells(OutputRow, 3).Value = Cell.Address
                            ResultSheet.Cells(OutputRow, 4).Value = Cell.Value
                            OutputRow = OutputRow + 1
                        End If
                    Next Cell
                Next ws
                wb.Close SaveChanges:=False
            End If
        End If
    Next File
    
    ' サブフォルダを再帰的に処理
    For Each SubFolder In Folder.SubFolders
        Call ProcessFolder(SubFolder, Pattern, ResultSheet, OutputRow)
    Next SubFolder
End Sub