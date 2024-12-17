VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbaModuleManegerView 
   Caption         =   "UserForm2"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9195.001
   OleObjectBlob   =   "VbaModuleManegerView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VbaModuleManegerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mUtil As clsVBAUtilityLib
Private mModuleMgr As clsVbaModuleManagerModel


Private Sub UserForm_Initialize()

    Set mUtil = New clsVBAUtilityLib
    Set mModuleMgr = New clsVbaModuleManagerModel

    Me.Caption = mModuleMgr.title
    
    Call update_comboBox

End Sub


'** ｲﾍﾞﾝﾄ **
Private Sub btnExport_Click()
    Dim res As Boolean
    res = export_project

    Call update_listbox
    
    If res Then
        show_msg "エクスポート終了"
    End If
End Sub


Private Sub btnimport_Click()
    Call import_project
    show_msg "インポート終了" & vbCrLf & "プロジェクトを保存してください"
End Sub


Private Sub btnOverwrite_Click()
    Call overwriting_project
    show_msg "エクスポート終了"
End Sub


Private Sub btnFinish_Click()
    Call finish
End Sub


Private Sub ListBox1_Change()
    Call update_info_txt
End Sub


Private Sub ComboBox1_Change()
    Call update_listbox
End Sub

Private Sub btnOpen_Click()
    'フォルダーopenのみ
    Call mModuleMgr.open_folder( _
        Me.ComboBox1.ListIndex + 1 _
    )
End Sub

'** ｻﾎﾟｰﾄ **
'終了
Private Sub finish()
    Me.Hide
    Unload Me
End Sub


'モジュールリストの更新
Private Sub update_listbox()

    With Me.ListBox1
        .Clear
        .ListIndex = -1
    End With

    If Me.ComboBox1.ListIndex < 0 Then Exit Sub

    Dim name As Variant
    For Each name In mModuleMgr.get_module_name_list(Me.ComboBox1.ListIndex + 1)
        Call Me.ListBox1.AddItem(name)
    Next
    
    Dim btnEnabled As Boolean
    If mModuleMgr.has_user_data(Me.ComboBox1.ListIndex + 1) Then
        btnEnabled = True
    Else
        btnEnabled = False
    End If
    
    With Me
        .btnOverwrite.Enabled = btnEnabled
        .btnImport.Enabled = btnEnabled
        .btnOpen.Enabled = btnEnabled
    End With

    'このproject
    If Me.ComboBox1.value = mModuleMgr.project_name Then
        Me.btnImport.Enabled = False
    End If

End Sub


'情報テキストの更新
Private Sub update_info_txt()

    If Me.ComboBox1.ListIndex < 0 Then
        Me.TextBox1.Text = vbNullString
        Exit Sub
    End If
    
    Dim value As String
    If Me.ListBox1.ListIndex < 0 Then
        value = ""
    Else
        value = Me.ListBox1.value
    End If
    
    Me.TextBox1.Text = mModuleMgr.get_module_info( _
        Me.ComboBox1.ListIndex + 1, _
        value _
    )

End Sub


'ComboBox初期設定
Private Sub update_comboBox()

    Dim projects As Collection
    Set projects = mModuleMgr.get_project_name_list()
    If projects.count < 1 Then
        MsgBox "対象プロジェクトがありません"
        Call finish
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To projects.count
        Call Me.ComboBox1.AddItem(projects.Item(i))
    Next
    ComboBox1.ListIndex = 0

End Sub


'Listから指定文字のｲﾝﾃﾞｯｸｽ取得
'param: value-検索文字
'param: lst-検索対象コレクション
'return: 該当インデックス
Function get_index_by_list( _
        ByVal value As Variant, _
        ByVal lst As Collection) As Long

    Dim i As Long
    For i = 1 To lst.count
        If lst.Item(i) = value Then
            get_index_by_list = i - 1
            Exit Function
        End If
    Next

    get_index_by_list = -1

End Function


'プロジェクトのインポート
Private Sub import_project()

    Dim projIdx As Long
    projIdx = Me.ComboBox1.ListIndex + 1
    
    Call mModuleMgr.import_project( _
        projIdx _
    )
    
    Call update_listbox

End Sub


'プロジェクトのエクスポート
Private Function export_project() As Boolean

    export_project = False
    
    Dim projIdx As Long
    projIdx = Me.ComboBox1.ListIndex + 1

    Dim msg As String
    msg = "CATVBAファイルのフォルダ内に作成しますか？" & vbCrLf & _
        "(はい-フォルダ内に作成　いいえ-ダイアログで指定)"
    
    Select Case MsgBox(msg, vbYesNoCancel + vbQuestion, mModuleMgr.title)
        Case vbYes
            'プロジェクトフォルダ内
            mModuleMgr.export_project_child_folder ( _
                projIdx _
            )
        Case vbNo
            'ダイアログで指定
            Dim dirPath As String
            dirPath = get_folder_path()
            If dirPath = vbNullString Then Exit Function
                
            Dim path As String
            path = mModuleMgr.get_dir_by_project_name( _
                projIdx, _
                dirPath _
            )
            
            Call mModuleMgr.export_project( _
                projIdx, _
                path _
            )
        Case Else
            'キャンセル
            Exit Function
    End Select
    
    export_project = True

End Function


'フォルダパス取得ダイアログ
'return: フォルダパス
Private Function get_folder_path() As String

    Dim dirPicker As New clsFolderPicker
    get_folder_path = dirPicker.show_folder_picker()

'    Dim folderPath As String
'    With Application.FileDialog(msoFileDialogFolderPicker)
'        .title = "フォルダを選択してください"
'        If .Show = -1 Then
'            folderPath = .SelectedItems(1)
'            get_folder_path = folderPath
'        Else
'            get_folder_path = vbNullString
'        End If
'    End With

End Function


'プロジェクトの上書きエクスポート
Private Sub overwriting_project()

    Call mModuleMgr.overwriting_project( _
        Me.ComboBox1.ListIndex + 1 _
    )

End Sub


'メッセージ
Private Sub show_msg( _
        ByVal msg As String)

    MsgBox msg, vbOKOnly, Me.Caption

End Sub


