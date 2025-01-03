Attribute VB_Name = "ModuleToSheet"
Option Explicit


Sub ExecToExcel()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�ړI�̃t�H���_��I�����Ă�������"
        .AllowMultiSelect = False
        If Not .Show = -1 Then
            MsgBox "���~���܂���"
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

    ' �V�������[�N�u�b�N���쐬
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

    ' ���[�N�u�b�N�����
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
    
    ' �t�H���_���̃t�@�C�������[�v���Ďw�肳�ꂽ�g���q�̃t�@�C�������W
    For Each file In folder.Files
        For Each ext In extensions
            If LCase(fileSystem.GetExtensionName(file.path)) = LCase(Replace(ext, ".", "")) Then
                filePaths.Add file.path
                Exit For
            End If
        Next ext
    Next file
    
    ' �R���N�V��������z��ɕϊ�
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
    
    ' �efrx�t�@�C�����ƂɐV�����V�[�g���쐬
    For Each frxFilePath In frxFilePaths
        ' �V�����V�[�g���쐬
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        
        ' �t�@�C�������擾���ăV�[�g���ɐݒ�
        frxFileName = Mid(frxFilePath, InStrRev(frxFilePath, "\") + 1)
        
        ' �t�@�C�����e��16�i���\���ɕϊ����ăV�[�g�ɏ����o��
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
    
    ' �I�����ꂽ�t�@�C�����Ƃɏ���
    For i = LBound(filePaths) To UBound(filePaths)
        filePath = filePaths(i)
        
        ' �t�@�C�������擾���ĈÍ���
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
        EncryptedFileName = EncryptLine(fileName)
        
        ' �f�t�H���g�̃V�[�g���擾
        Set ws = wb.Sheets(wb.Sheets.Count)
        
        ' �V�[�g��1�s�ڂɈÍ������ꂽ�t�@�C��������������
        ws.Cells(1, 1).Value = EncryptedFileName
        
        ' �e�L�X�g�t�@�C���̓��e��ǂݍ��݁A2�s�ڈȍ~�ɏ�������
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
        
        ' ���̃V�[�g��ǉ�
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
        .Title = "Excel�t�@�C����I�����Ă�������"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = False
        If Not .Show = -1 Then
            MsgBox "���~���܂�"
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
    
    ' �t�H���_�[�p�X���擾
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
    
    ' �w�肳�ꂽExcel�t�@�C�����J��
    Set wb = Workbooks.Open(xlsxFilePath)
    
    ' �e�V�[�g������
    For Each ws In wb.Sheets
        ' 1�s�ڂ̈Í������ꂽ�t�@�C�����𕜌�
        EncryptedFileName = ws.Cells(1, 1).Value
        DecryptedFileName = DecryptLine(EncryptedFileName)
        
        ' �t�@�C�����Ɗg���q�𕪗�
        fileExt = LCase(Right(DecryptedFileName, 3))
        
        ' �ۑ���̃p�X��ݒ�i�g���q�Ȃ��j
        SavePath = outputFolderPath & "\" & DecryptedFileName
        
        If fileExt = "frx" Then
            ' �o�C�i���t�@�C���Ƃ��ĕۑ�
            Call ExportSheetToFRXFiles(ws, outputFolderPath)
        Else
            ' �e�L�X�g�t�@�C���Ƃ��ĕۑ�
            fileNum = FreeFile
            Open SavePath For Output As #fileNum
            
            ' 2�s�ڈȍ~�̓��e�𕜌����ď�������
            rowNum = 2
            Do
                DecryptedLine = DecryptLine(ws.Cells(rowNum, 1).Value)
                Print #fileNum, DecryptedLine
                rowNum = rowNum + 1
            Loop Until ws.Cells(rowNum, 1).Value = "" And ws.Cells(rowNum + 1, 1).Value = ""
            
            Close #fileNum
        End If
    Next ws
    
    ' ���[�N�u�b�N�����
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
    
    ' �eFRX�t�@�C��������
    Do While ws.Cells(rowNum, 1).Value <> ""
        ' �t�@�C�������擾
        fileName = DecryptLine(ws.Cells(rowNum, 1).Value)
        filePath = outputFolderPath & "\" & fileName
        
        ' �t�@�C�����e��16�i���\������o�C�g�z��ɕϊ�
        hexString = ws.Cells(rowNum + 1, 1).Value
        ReDim byteArray(Len(hexString) \ 2 - 1)
        For i = 0 To UBound(byteArray)
            byteArray(i) = CByte("&H" & Mid(hexString, 2 * i + 1, 2))
        Next i
        
        ' �o�C�g�z����t�@�C���ɏ�������
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

