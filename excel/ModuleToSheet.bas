Attribute VB_Name = "ModuleToSheet"
Option Explicit


Sub ExecToExcel()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "目的のフォルダを選択してください"
        .AllowMultiSelect = False
        If Not .Show = -1 Then
            MsgBox "中止しました"
            Exit Sub
        End If

        folderPath = .SelectedItems(1)
    End With
    
    Dim arrFiles As Variant
    arrFiles = GetFilesWithExtensions( _
        folderPath, _
        Array(".bas", ".cls", ".frm") _
    )
    If UBound(arrFiles) < 0 Then Exit Sub

    ' 新しいワークブックを作成
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    Call CreateTextFilesToNewWorkbook( _
        arrFiles, _
        wb _
    )

    arrFiles = GetFilesWithExtensions( _
        folderPath, _
        Array(".frx") _
    )
    If Not UBound(arrFiles) < 0 Then
        Call ExportFRXFilesToSheet( _
            arrFiles, _
            wb _
        )
    End If

    Dim SavePath As String
    SavePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If SavePath <> "False" Then
        wb.SaveAs fileName:=SavePath, FileFormat:=xlOpenXMLWorkbook
    End If

    ' ワークブックを閉じる
    wb.Close SaveChanges:=False

End Sub


Private Function GetFilesWithExtensions( _
        ByVal folderPath As String, _
        ByVal extensions As Variant) As Variant

    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim filePaths As Collection
    Dim ext As Variant
    Dim filePathArray() As String
    Dim i As Integer
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    Set filePaths = New Collection
    
    ' フォルダ内のファイルをループして指定された拡張子のファイルを収集
    For Each file In folder.Files
        For Each ext In extensions
            If LCase(fileSystem.GetExtensionName(file.path)) = LCase(Replace(ext, ".", "")) Then
                filePaths.Add file.path
                Exit For
            End If
        Next ext
    Next file
    
    ' コレクションから配列に変換
    ReDim filePathArray(1 To filePaths.Count)
    For i = 1 To filePaths.Count
        filePathArray(i) = filePaths(i)
    Next i
    
    GetFilesWithExtensions = filePathArray
End Function


Sub ExportFRXFilesToSheet(frxFilePaths As Variant, wb As Workbook)
    Dim ws As Worksheet
    Dim frxFilePath As Variant
    Dim frxFileName As String
    Dim frxContent As String
    
    ' 各frxファイルごとに新しいシートを作成
    For Each frxFilePath In frxFilePaths
        ' 新しいシートを作成
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        
        ' ファイル名を取得してシート名に設定
        frxFileName = Mid(frxFilePath, InStrRev(frxFilePath, "\") + 1)
        
        ' ファイル内容を16進数表現に変換してシートに書き出す
        frxContent = ConvertFRXToHex(frxFilePath)
        ws.Cells(1, 1).Value = EncryptLine(frxFileName)
        ws.Cells(2, 1).Value = frxContent
    Next frxFilePath
    
    Set ws = Nothing
End Sub


Function ConvertFRXToHex(filePath As Variant) As String
    Dim fileNum As Integer
    Dim byteArray() As Byte
    Dim hexString As String
    Dim i As Integer
    
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    ReDim byteArray(LOF(fileNum) - 1)
    Get #fileNum, , byteArray
    Close #fileNum
    
    hexString = ""
    For i = LBound(byteArray) To UBound(byteArray)
        hexString = hexString & Right("0" & Hex(byteArray(i)), 2)
    Next i
    
    ConvertFRXToHex = hexString
End Function


Private Sub CreateTextFilesToNewWorkbook( _
        ByVal filePaths As Variant, _
        ByVal wb As Workbook)

    Dim ws As Worksheet
    Dim filePath As String
    Dim fileName As String
    Dim FileLine As String
    Dim fileNum As Integer
    Dim i As Integer
    Dim rowNum As Integer
    Dim EncryptedLine As String
    Dim EncryptedFileName As String
    
    ' 選択されたファイルごとに処理
    For i = LBound(filePaths) To UBound(filePaths)
        filePath = filePaths(i)
        
        ' ファイル名を取得して暗号化
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
        EncryptedFileName = EncryptLine(fileName)
        
        ' デフォルトのシートを取得
        Set ws = wb.Sheets(wb.Sheets.Count)
        
        ' シートの1行目に暗号化されたファイル名を書き込み
        ws.Cells(1, 1).Value = EncryptedFileName
        
        ' テキストファイルの内容を読み込み、2行目以降に書き込み
        fileNum = FreeFile
        Open filePath For Input As #fileNum
        rowNum = 2
        Do While Not EOF(fileNum)
            Line Input #fileNum, FileLine
            EncryptedLine = EncryptLine(FileLine)
            ws.Cells(rowNum, 1).Value = EncryptedLine
            rowNum = rowNum + 1
        Loop
        Close #fileNum
        
        ' 次のシートを追加
        If i < UBound(filePaths) Then
            wb.Sheets.Add After:=wb.Sheets(wb.Sheets.Count)
        End If
    Next i
    
    Set ws = Nothing
    Set wb = Nothing

End Sub


Private Function EncryptLine( _
        ByVal Line As String) As String

    Dim EncryptedLine As String
    Dim i As Integer
    Dim CharCode As Integer
    
    EncryptedLine = ""
    For i = 1 To Len(Line)
        CharCode = Asc(Mid(Line, i, 1))
        EncryptedLine = EncryptedLine & Chr(CharCode + 1)
    Next i
    
    EncryptLine = EncryptedLine

End Function

'*******************
Sub ExecToModule()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Excelファイルを選択してください"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = False
        If Not .Show = -1 Then
            MsgBox "中止します"
        End If
        
        filePath = .SelectedItems(1)
    End With
    
    Dim dirPath As String
    dirPath = GetFolderPath(filePath)
    
    Call ExportSheetsToTextFilesWithDecryption( _
        filePath, _
        dirPath _
    )

End Sub


Private Function GetFolderPath( _
        ByVal filePath As String) As String

    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダーパスを取得
    GetFolderPath = fileSystem.GetParentFolderName(filePath)
    
    Set fileSystem = Nothing
End Function


Private Sub ExportSheetsToTextFilesWithDecryption( _
        ByVal xlsxFilePath As String, _
        ByVal outputFolderPath As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim SavePath As String
    Dim DecryptedLine As String
    Dim fileNum As Integer
    Dim rowNum As Integer
    Dim EncryptedFileName As String
    Dim DecryptedFileName As String
    Dim fileExt As String
    Dim i As Long
    
    ' 指定されたExcelファイルを開く
    Set wb = Workbooks.Open(xlsxFilePath)
    
    ' 各シートを処理
    For Each ws In wb.Sheets
        ' 1行目の暗号化されたファイル名を復元
        EncryptedFileName = ws.Cells(1, 1).Value
        DecryptedFileName = DecryptLine(EncryptedFileName)
        
        ' ファイル名と拡張子を分離
        fileExt = LCase(Right(DecryptedFileName, 3))
        
        ' 保存先のパスを設定（拡張子なし）
        SavePath = outputFolderPath & "\" & DecryptedFileName
        
        If fileExt = "frx" Then
            ' バイナリファイルとして保存
            Call ExportSheetToFRXFiles(ws, outputFolderPath)
        Else
            ' テキストファイルとして保存
            fileNum = FreeFile
            Open SavePath For Output As #fileNum
            
            ' 2行目以降の内容を復元して書き込み
            rowNum = 2
            Do
                DecryptedLine = DecryptLine(ws.Cells(rowNum, 1).Value)
                Print #fileNum, DecryptedLine
                rowNum = rowNum + 1
            Loop Until ws.Cells(rowNum, 1).Value = "" And ws.Cells(rowNum + 1, 1).Value = ""
            
            Close #fileNum
        End If
    Next ws
    
    ' ワークブックを閉じる
    wb.Close SaveChanges:=False
    
    Set ws = Nothing
    Set wb = Nothing

End Sub


Sub ExportSheetToFRXFiles( _
        ByVal ws As Worksheet, _
        ByVal outputFolderPath As String)

    Dim fileSystem As Object
    Dim filePath As String
    Dim fileName As String
    Dim fileContent As String
    Dim fileNum As Integer
    Dim rowNum As Integer
    Dim byteArray() As Byte
    Dim hexString As String
    Dim i As Integer
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    rowNum = 1
    
    ' 各FRXファイルを処理
    Do While ws.Cells(rowNum, 1).Value <> ""
        ' ファイル名を取得
        fileName = DecryptLine(ws.Cells(rowNum, 1).Value)
        filePath = outputFolderPath & "\" & fileName
        
        ' ファイル内容を16進数表現からバイト配列に変換
        hexString = ws.Cells(rowNum + 1, 1).Value
        ReDim byteArray(Len(hexString) \ 2 - 1)
        For i = 0 To UBound(byteArray)
            byteArray(i) = CByte("&H" & Mid(hexString, 2 * i + 1, 2))
        Next i
        
        ' バイト配列をファイルに書き込み
        fileNum = FreeFile
        Open filePath For Binary As #fileNum
        Put #fileNum, , byteArray
        Close #fileNum
        
        rowNum = rowNum + 2
    Loop
    
    Set fileSystem = Nothing

End Sub


Private Function DecryptLine( _
        ByVal Line As String) As String

    Dim DecryptedLine As String
    Dim i As Integer
    Dim CharCode As Integer
    
    DecryptedLine = ""
    For i = 1 To Len(Line)
        CharCode = Asc(Mid(Line, i, 1))
        DecryptedLine = DecryptedLine & Chr(CharCode - 1)
    Next i
    
    DecryptLine = DecryptedLine

End Function

