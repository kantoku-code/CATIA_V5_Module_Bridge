VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVbaModuleManagerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEBUG_MODE = False                        'デバッグモード
Private Const APP_NAME = "ModuleBridge"                 'アプリネーム
Private Const VERSION = "0.0.6"                         'バージョン
Private Const THIS_PROJECT_NAME = "ModuleBridgeProject" 'このプロジェクト名
Private Const USER_JSON_NAME = "user_data.json"         'データ保存ファイル名
Private Const SEPARATOR = "|"                           'パスの区切り文字-VBA-JSON対策
Private Const SOURCE_FOLDER_NAME = "source"             'CATVBA内に作成する際のフォルダ

Private mVbe As Object                                  'VBAエディタ
Private mUtil As clsVBAUtilityLib                       'ユーティリティ
Private mProjectLst As Collection                       'プロジェクトリスト
Private mUserDict As Object                             'ユーザーデータ辞書
Private mJson As clsJsonConverter                       'JsonParser


Private Sub Class_Initialize()

    Set mUtil = New clsVBAUtilityLib
    mUtil.DebugMode = DEBUG_MODE

    Set mVbe = mUtil.GetVbe()
    Set mJson = New clsJsonConverter
    Set mProjectLst = get_project_list()

    Set mUserDict = get_user_data()
    
End Sub


Private Sub Class_Terminate()
    
    Dim path As String
    path = get_this_dir_path() & "\" & USER_JSON_NAME

    Call cleanUp_userdata
    save_json path, mUserDict
    
End Sub

'***プロパティ***
'アプリタイトル
Property Get title() As String
    title = APP_NAME & " Ver" & VERSION
End Property


'このプロジェクト名
Property Get project_name() As String
   project_name = THIS_PROJECT_NAME
End Property


'***メソッド***
'プロジェクトの取得
'return:VBProjectのコレクション
Function get_project_list() As Collection

    Dim lst As Collection
    Set lst = New Collection

    Dim proj As Object
    For Each proj In mVbe.vBProjects
        lst.Add proj
    Next
    
    Set get_project_list = lst
    
End Function


'プロジェクト名の取得
'return:プロジェクト名
Function get_project_name_list() As Collection

    Dim lst As Collection
    Set lst = New Collection

    Dim proj As Object
    For Each proj In mProjectLst
        lst.Add proj.name
    Next
    
    Set get_project_name_list = lst
    
End Function


'プロジェクト内のモジュール取得
'param: projectIdx-プロジェクトのインデックス
'return: モジュールコレクション
Function get_module_list( _
        ByVal projectIdx As Long) As Collection

    Dim lst As Collection
    Set lst = New Collection
    
    Set get_module_list = lst

    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    
    Dim comp As Object
    For Each comp In proj.VBComponents
        If is_target_module_type(comp) Then
            lst.Add comp
        End If
    Next
    
End Function


'プロジェクト内のモジュール名取得
'param: projectIdx-プロジェクトのインデックス
'return: モジュール名コレクション
Function get_module_name_list( _
        ByVal projectIdx As Long) As Collection

    Dim lst As Collection
    Set lst = New Collection

    Set get_module_name_list = lst
    
    Dim comp As Object
    For Each comp In get_module_list(projectIdx)
        If is_target_module_type(comp) Then
            lst.Add comp.name
        End If
    Next
    
End Function


'指定モジュール内関数名ﾘｽﾄ取得
'param: comp-モジュール
'return: 関数名の文字列
Private Function get_fanction_names( _
        ByVal comp As Object) As String

    Dim lst As Collection
    Set lst = New Collection
    
    Dim buf As String
    Dim i As Long
    With comp.CodeModule
        For i = 1 To .CountOfLines
            If buf <> .ProcOfLine(i, 0) Then
                buf = .ProcOfLine(i, 0)
                lst.Add "  [" + buf + "]"
            End If
        Next i
    End With
    
    get_fanction_names = Join( _
        mUtil.CollectionToArrayLet(lst), _
        vbCrLf _
    )

End Function


'vbaのモジュールか
'vbext_ct_StdModule 1 標準モジュール
'vbext_ct_ClassModule 2 クラス モジュール
'vbext_ct_MSForm 3 Microsoft Forms
'vbext_ct_ActiveXDesigner 11 ActiveX デザイナー
'vbext_ct_Document 100 ドキュメント モジュール
Private Function is_target_module_type( _
        ByVal module As Object) As Boolean

    Select Case module.Type
        Case 1, 2, 3
            is_target_module_type = True
        Case Else
            is_target_module_type = False
    End Select

End Function


'モジュール情報取得
'param: projectIdx-プロジェクトのインデックス
'return: プロジェクトのザックリ情報
Function get_module_info( _
        ByVal projectIdx As Long, _
        ByVal moduleName As String) As String

    get_module_info = vbNullString

    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    
    Dim comp As Object
    On Error Resume Next
        Set comp = proj.VBComponents(moduleName)
    On Error GoTo 0
    If comp Is Nothing Then Exit Function

    'プロジェクト全体の行数
    Dim totalCharCount As Long
    totalCharCount = 0
    
    Dim compX As Object
    For Each compX In proj.VBComponents
        totalCharCount = totalCharCount + GetValidCharacterCount(compX)
    Next

    get_module_info = _
        "プロジェクト名 : " & proj.name & vbCrLf & _
        "総ステップ数　 : " & totalCharCount & vbCrLf & _
        "ファイルパス 　: " & proj.Filename & vbCrLf & _
        "モジュール名 　: " & moduleName & "." & get_extension(comp) & vbCrLf & _
        "モジュールステップ数: " & GetValidCharacterCount(comp) & vbCrLf & _
        "関数名リスト : " & vbCrLf & get_fanction_names(comp)

End Function


'コンポーネントのコメント部分の除いた文字数の取得
'param: vbComp-対象コンポーネント
Private Function GetValidCharacterCount( _
        ByVal vbComp As Object) As Long

    Dim codeMod As Object ' CodeModule
    Dim totalChars As Long
    Dim i As Long
    Dim codeLine As String
    
    ' CodeModuleオブジェクトを取得
    Set codeMod = vbComp.CodeModule
    
    totalChars = 0
    ' 各行をチェックしてコメントを除外
    For i = 1 To codeMod.CountOfLines
        codeLine = Trim(codeMod.Lines(i, 1))
        If Left(codeLine, 1) <> "'" Then
            totalChars = totalChars + Len(codeLine)
        End If
    Next i
    
    ' 結果を返す
    GetValidCharacterCount = totalChars

End Function


'プロジェクト内のvbaモジュールを削除しインポート
'param: projectIdx-プロジェクトのインデックス
Sub import_project( _
        ByVal projectIdx As Long)

    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim Key As String
    Key = Replace(proj.Filename, "\", SEPARATOR)
    
    If Not mUserDict.exists(Key) Then Exit Sub
    
    Dim path As String
    path = mUserDict(Key)("path")
    If Not mUtil.GetFso().FolderExists(path) Then Exit Sub
    
    Call remove_modules(projectIdx)
    Call import_modules(proj, path)

End Sub


'プロジェクト名のフォルダの取得-なきゃ作る
'param: projectIdx-プロジェクトのインデックス
'param: parentDirPath-親フォルダパス
'return: 作成したフォルダパス
Function get_dir_by_project_name( _
        ByVal projectIdx As Long, _
        ByVal parentDirPath As String) As String
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    
    Dim path As String
    path = parentDirPath & "\" & proj.name
    
    Dim fso As Object
    Set fso = mUtil.GetFso()
    
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    
    get_dir_by_project_name = path
    
End Function


'フォルダ内のファイルのインポート
'param: proj-プロジェクト
'param: dirPath-インポートするフォルダパス
Private Sub import_modules( _
        ByVal proj As Object, _
        ByVal dirPath As String)
    
    Dim fileLst As Collection
    Set fileLst = get_target_file_path_list( _
        dirPath, _
        Array( _
            "cls", _
            "frm", _
            "bas" _
        ) _
    )

    Dim path As Variant ' String
    For Each path In fileLst
        On Error Resume Next
            proj.VBComponents.import path
        On Error GoTo 0
    Next

End Sub


'指定パス内の指定拡張子のファイルの取得
'param: path-フォルダパス
'param: targetExtensions-対象拡張子
'return: ファイルパスコレクション
Private Function get_target_file_path_list( _
        ByVal path As String, _
        ByVal targetExtensions As Variant) As Collection

    Dim lst As Collection
    Set lst = New Collection
    Set get_target_file_path_list = lst

    
    Dim fso As Object
    Set fso = mUtil.GetFso()
    
    Dim dir As Object
    Set dir = fso.GetFolder(path)

    Dim file As Object
    Dim ext As String
    For Each file In dir.Files
        ext = fso.GetExtensionName(file.name)
        
        If Not UBound(Filter(targetExtensions, ext)) < 0 Then
            lst.Add file.path
        End If
    Next
    
End Function


'プロジェクト内のvbaモジュールを削除
'param: projectIdx-プロジェクトのインデックス
Private Sub remove_modules( _
        ByVal projectIdx As Long)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    
    Dim lst As Collection
    Set lst = get_module_name_list(projectIdx)
    
    Dim name As Variant ' String
    Dim comp As Object
    For Each name In get_module_name_list(projectIdx)
        Set comp = proj.VBComponents(name)
        proj.VBComponents.Remove comp
    Next

End Sub


'プロジェクト内のvbaモジュールをエクスポート
'param: projectIdx-プロジェクトのインデックス
'param: dirPath-フォルダパス
'param: isOpen(opt)-開くか?
Sub export_project( _
        ByVal projectIdx As Long, _
        ByVal dirPath As String, _
        Optional ByVal isOpen As Boolean = True)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    
    Dim name 'As String
    Dim comp As Object
    For Each name In get_module_name_list(projectIdx)
        Set comp = proj.VBComponents(name)
        
        On Error Resume Next
    
        comp.Export dirPath & "\" & comp.name & "." & get_extension(comp)
        
        If Not Err.Number = 0 Then
            MsgBox "書き出しに失敗しました。.git以外を削除し再トライしてください"
            Exit Sub
        End If
        
        On Error GoTo 0
    Next
    
    update_userdata proj, dirPath
    
    If isOpen Then
        open_folder_dy_explorer dirPath
    End If
    
End Sub


'プロジェクト内のvbaモジュールをCATVBAフォルダ内にエクスポート
'param: projectIdx-プロジェクトのインデックス
'param: isOpen(opt)-開くか?
Sub export_project_child_folder( _
        ByVal projectIdx As Long, _
        Optional ByVal isOpen As Boolean = True)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim fso As Object
    Set fso = mUtil.GetFso()

    Dim dirPath As String
    dirPath = fso.GetParentFolderName(proj.Filename) & "\" & SOURCE_FOLDER_NAME

    If Not fso.FolderExists(dirPath) Then
        fso.CreateFolder (dirPath)
    End If

    Call export_project(projectIdx, dirPath)
    
End Sub


'プロジェクト内のvbaモジュールを上書きエクスポート
'param: projectIdx-プロジェクトのインデックス
Sub overwriting_project( _
        ByVal projectIdx As Long)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim Key As String
    Key = Replace(proj.Filename, "\", SEPARATOR)

    Dim dirPath As String
    dirPath = mUserDict(Key)("path")
    
    Dim fso As Object
    Set fso = mUtil.GetFso()
    
    '見つからない場合は削除で良いのか？
    If Not fso.FolderExists(dirPath) Then Exit Sub
    
    Call remove_module_files(dirPath)
    Call export_project(projectIdx, dirPath)

End Sub


'プロジェクトフォルダのオープンのみ
'param: projectIdx-プロジェクトのインデックス
Sub open_folder( _
        ByVal projectIdx As Long)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim Key As String
    Key = Replace(proj.Filename, "\", SEPARATOR)

    Dim dirPath As String
    dirPath = mUserDict(Key)("path")
    
    Dim fso As Object
    Set fso = mUtil.GetFso()
    
    '見つからない場合は削除で良いのか？
    If Not fso.FolderExists(dirPath) Then Exit Sub
    
    Call open_folder_dy_explorer(dirPath)

End Sub


'ユーザーデータの更新
'vba-jsonでのキーで"\\"の扱いが悪いためSEPARATORに置換
'param: proj-プロジェクト
'param: path-保存先Dirパス
Private Sub update_userdata( _
        ByVal proj As Object, _
        ByVal path As String)
    
    Dim Key As String
    Key = Replace(proj.Filename, "\", SEPARATOR)
    
    If mUserDict.exists(Key) Then
        mUserDict.Remove Key
    End If

    Dim info As Object
    Set info = mUtil.InitDict()
    info.Add "name", proj.name
    info.Add "path", path

    mUserDict.Add Key, info
    
    Call cleanUp_userdata
        
End Sub


'エクスポート用の拡張子取得
'https://learn.microsoft.com/ja-jp/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model
'vbext_ct_StdModule 1 標準モジュール
'vbext_ct_ClassModule 2 クラス モジュール
'vbext_ct_MSForm 3 Microsoft Forms
'vbext_ct_ActiveXDesigner 11 ActiveX デザイナー
'vbext_ct_Document 100 ドキュメント モジュール
Private Function get_extension( _
        ByVal module As Object) As String

    Dim extension As String
    Select Case module.Type
        Case 1 'vbext_ct_StdModule
            extension = "bas"
        Case 2 'vbext_ct_ClassModule
            extension = "cls"
        Case 3 'vbext_ct_MSForm
            extension = "frm"
        Case Else
            extension = ""
    End Select
    
    get_extension = extension

End Function


'モジュールファイルの削除
Private Sub remove_module_files( _
        ByVal dirPath As String)
        
    Dim fso As Object
    Set fso = mUtil.GetFso()
    
    If Not fso.FolderExists(dirPath) Then Exit Sub

    Dim paths As Collection
    Set paths = get_target_file_path_list( _
        dirPath, _
        Array( _
            "cls", _
            "frm", _
            "frx", _
            "bas" _
        ) _
    )

    Dim path As Variant ' String
    For Each path In paths
        On Error Resume Next
            fso.DeleteFile path
        On Error GoTo 0
    Next
    
    'ウェイト
    Call mUtil.Wait(1)

End Sub


'ｱｸﾃｨﾌﾞﾌﾟﾛｼﾞｪｸﾄ取得
'return: このプロジェクトオブジェクト
Function get_this_project() As Object

    Set get_this_project = Nothing
    
    Dim thisProj As Object
    Dim proj As Object
    For Each proj In mVbe.vBProjects
        If proj.name = THIS_PROJECT_NAME Then
            Set get_this_project = proj
            Exit Function
        End If
    Next

End Function


'アクティブプロジェクトのフォルダパス取得
'return: このプロジェクトのフォルダパス
Private Function get_this_dir_path() As String

    get_this_dir_path = mUtil.GetFso().GetParentFolderName( _
        get_this_project().Filename _
    )

End Function


'フォルダをエクスプローラーで開く
'param: dirPath-フォルダパス
Private Sub open_folder_dy_explorer( _
        ByVal dirPath As String)

    On Error Resume Next
        Shell "C:\Windows\Explorer.exe " & dirPath, vbNormalFocus
    On Error GoTo 0

End Sub
        

'ユーザーデータの取得
'return: dict
Private Function get_user_data() As Object

    Dim path As String
    path = get_this_dir_path() & "\" & USER_JSON_NAME
    
    mUtil.Dump "userdata_path:" & path
    
    Dim dict As Object
    If mUtil.GetFso().FileExists(path) Then
        Set dict = load_json(path)
        dict("ver") = VERSION
    Else
        Set dict = mUtil.InitDict()
        dict.Add "ver", VERSION
    End If
    
    Set get_user_data = dict
    
End Function


'ユーザーデータのクリーンアップ
Private Sub cleanUp_userdata()

    Dim fso As Object
    Set fso = mUtil.GetFso()

    Dim removeLst As Collection
    Set removeLst = New Collection
    
    Dim filePath As String
    Dim dirPath As String

    Dim Key As Variant 'String
    For Each Key In mUserDict.keys()
        If Key = "ver" Then GoTo continue

        filePath = Replace(Key, SEPARATOR, "\")
        If Not fso.FileExists(filePath) Then
            removeLst.Add Key
            GoTo continue
        End If
        
        dirPath = mUserDict(Key)("path")
        If Not fso.FolderExists(dirPath) Then
            removeLst.Add Key
            GoTo continue
        End If
continue:
    Next
    
    For Each Key In removeLst
        mUserDict.Remove Key
    Next
    
End Sub


'プロジェクトが過去に保存されているか
'param: projectIdx-プロジェクトのインデックス
Function has_user_data( _
        ByVal projectIdx As Long) As Boolean

    has_user_data = False

    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    
    has_user_data = mUserDict.exists(Replace(proj.Filename, "\", SEPARATOR))

End Function


'jsonの書き込み
'param: path-jsonファイルパス
'param: 書き込むdict
Private Sub save_json( _
        ByVal path As String, _
        ByVal dict As Object)
    
    Call mUtil.WriteFile( _
        path, _
        mJson.ConvertToJson(dict, 3) _
    )
    
End Sub


'jsonの読み込み
'param: path-jsonファイルパス
'return: dict
Private Function load_json( _
        ByVal path As String) As Object
    
    Set load_json = Nothing
    
    If Not mUtil.GetFso().FileExists(path) Then Exit Function
    
    Dim dict As Object
    Set dict = mJson.ParseJson(mUtil.ReadFile(path))
    
    Set load_json = dict
    
End Function

