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

Private Const DEBUG_MODE = False                        '�f�o�b�O���[�h
Private Const APP_NAME = "ModuleBridge"                 '�A�v���l�[��
Private Const VERSION = "0.0.5"                         '�o�[�W����
Private Const THIS_PROJECT_NAME = "ModuleBridgeProject" '���̃v���W�F�N�g��
Private Const USER_JSON_NAME = "user_data.json"         '�f�[�^�ۑ��t�@�C����
Private Const SEPARATOR = "|"                           '�p�X�̋�؂蕶��-VBA-JSON�΍�
Private Const SOURCE_FOLDER_NAME = "source"             'CATVBA���ɍ쐬����ۂ̃t�H���_

Private mVbe As Object                                  'VBA�G�f�B�^
Private mUtil As clsVBAUtilityLib                       '���[�e�B���e�B
Private mProjectLst As Collection                       '�v���W�F�N�g���X�g
Private mUserDict As Object                             '���[�U�[�f�[�^����
Private mJson As clsJsonConverter                       'JsonParser


Private Sub Class_Initialize()

    Set mUtil = New clsVBAUtilityLib
    mUtil.DEBUG_MODE = DEBUG_MODE

    Set mVbe = mUtil.get_vbe()
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

'***�v���p�e�B***
'�A�v���^�C�g��
Property Get title() As String
    title = APP_NAME & " Ver" & VERSION
End Property


'���̃v���W�F�N�g��
Property Get project_name() As String
   project_name = THIS_PROJECT_NAME
End Property


'***���\�b�h***
'�v���W�F�N�g�̎擾
'return:VBProject�̃R���N�V����
Function get_project_list() As Collection

    Dim lst As Collection
    Set lst = New Collection

    Dim proj As Object
    For Each proj In mVbe.vBProjects
        lst.Add proj
    Next
    
    Set get_project_list = lst
    
End Function


'�v���W�F�N�g���̎擾
'return:�v���W�F�N�g��
Function get_project_name_list() As Collection

    Dim lst As Collection
    Set lst = New Collection

    Dim proj As Object
    For Each proj In mProjectLst
        lst.Add proj.name
    Next
    
    Set get_project_name_list = lst
    
End Function


'�v���W�F�N�g���̃��W���[���擾
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
'return: ���W���[���R���N�V����
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


'�v���W�F�N�g���̃��W���[�����擾
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
'return: ���W���[�����R���N�V����
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


'�w�胂�W���[�����֐���ؽĎ擾
'param: comp-���W���[��
'return: �֐����̕�����
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
        mUtil.collection_to_array_let(lst), _
        vbCrLf _
    )

End Function


'vba�̃��W���[����
'vbext_ct_StdModule 1 �W�����W���[��
'vbext_ct_ClassModule 2 �N���X ���W���[��
'vbext_ct_MSForm 3 Microsoft Forms
'vbext_ct_ActiveXDesigner 11 ActiveX �f�U�C�i�[
'vbext_ct_Document 100 �h�L�������g ���W���[��
Private Function is_target_module_type( _
        ByVal module As Object) As Boolean

    Select Case module.Type
        Case 1, 2, 3
            is_target_module_type = True
        Case Else
            is_target_module_type = False
    End Select

End Function


'���W���[�����擾
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
'return: �v���W�F�N�g�̃U�b�N�����
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

    get_module_info = _
        "�v���W�F�N�g�� : " & proj.name & vbCrLf & _
        "�t�@�C���p�X : " & proj.fileName & vbCrLf & _
        "���W���[���� : " & moduleName & "." & get_extension(comp) & vbCrLf & _
        "�֐������X�g : " & vbCrLf & get_fanction_names(comp)

End Function


'�v���W�F�N�g����vba���W���[�����폜���C���|�[�g
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
Sub import_project( _
        ByVal projectIdx As Long)

    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim key As String
    key = Replace(proj.fileName, "\", SEPARATOR)
    
    If Not mUserDict.Exists(key) Then Exit Sub
    
    Dim path As String
    path = mUserDict(key)("path")
    If Not mUtil.get_fso().FolderExists(path) Then Exit Sub
    
    Call remove_modules(projectIdx)
    Call import_modules(proj, path)

End Sub


'�v���W�F�N�g���̃t�H���_�̎擾-�Ȃ�����
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
'param: parentDirPath-�e�t�H���_�p�X
'return: �쐬�����t�H���_�p�X
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
    Set fso = mUtil.get_fso()
    
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    
    get_dir_by_project_name = path
    
End Function


'�t�H���_���̃t�@�C���̃C���|�[�g
'param: proj-�v���W�F�N�g
'param: dirPath-�C���|�[�g����t�H���_�p�X
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


'�w��p�X���̎w��g���q�̃t�@�C���̎擾
'param: path-�t�H���_�p�X
'param: targetExtensions-�Ώۊg���q
'return: �t�@�C���p�X�R���N�V����
Private Function get_target_file_path_list( _
        ByVal path As String, _
        ByVal targetExtensions As Variant) As Collection

    Dim lst As Collection
    Set lst = New Collection
    Set get_target_file_path_list = lst

    
    Dim fso As Object
    Set fso = mUtil.get_fso()
    
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


'�v���W�F�N�g����vba���W���[�����폜
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
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


'�v���W�F�N�g����vba���W���[�����G�N�X�|�[�g
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
'param: dirPath-�t�H���_�p�X
'param: isOpen(opt)-�J����?
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
        comp.Export dirPath & "\" & comp.name & "." & get_extension(comp)
    Next
    
    update_userdata proj, dirPath
    
    If isOpen Then
        open_folder_dy_explorer dirPath
    End If
    
End Sub


'�v���W�F�N�g����vba���W���[����CATVBA�t�H���_���ɃG�N�X�|�[�g
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
'param: isOpen(opt)-�J����?
Sub export_project_child_folder( _
        ByVal projectIdx As Long, _
        Optional ByVal isOpen As Boolean = True)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim fso As Object
    Set fso = mUtil.get_fso()

    Dim dirPath As String
    dirPath = fso.GetParentFolderName(proj.fileName) & "\" & SOURCE_FOLDER_NAME

    If Not fso.FolderExists(dirPath) Then
        fso.CreateFolder (dirPath)
    End If

    Call export_project(projectIdx, dirPath)
    
End Sub


'�v���W�F�N�g����vba���W���[�����㏑���G�N�X�|�[�g
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
Sub overwriting_project( _
        ByVal projectIdx As Long)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim key As String
    key = Replace(proj.fileName, "\", SEPARATOR)

    Dim dirPath As String
    dirPath = mUserDict(key)("path")
    
    Dim fso As Object
    Set fso = mUtil.get_fso()
    
    '������Ȃ��ꍇ�͍폜�ŗǂ��̂��H
    If Not fso.FolderExists(dirPath) Then Exit Sub
    
    Call remove_module_files(dirPath)
    Call export_project(projectIdx, dirPath)

End Sub


'�v���W�F�N�g�t�H���_�̃I�[�v���̂�
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
Sub open_folder( _
        ByVal projectIdx As Long)
    
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub

    Dim key As String
    key = Replace(proj.fileName, "\", SEPARATOR)

    Dim dirPath As String
    dirPath = mUserDict(key)("path")
    
    Dim fso As Object
    Set fso = mUtil.get_fso()
    
    '������Ȃ��ꍇ�͍폜�ŗǂ��̂��H
    If Not fso.FolderExists(dirPath) Then Exit Sub
    
    Call open_folder_dy_explorer(dirPath)

End Sub


'���[�U�[�f�[�^�̍X�V
'vba-json�ł̃L�[��"\\"�̈�������������SEPARATOR�ɒu��
'param: proj-�v���W�F�N�g
'param: path-�ۑ���Dir�p�X
Private Sub update_userdata( _
        ByVal proj As Object, _
        ByVal path As String)
    
    Dim key As String
    key = Replace(proj.fileName, "\", SEPARATOR)
    
    If mUserDict.Exists(key) Then
        mUserDict.Remove key
    End If

    Dim info As Object
    Set info = mUtil.init_dict()
    info.Add "name", proj.name
    info.Add "path", path

    mUserDict.Add key, info
    
    Call cleanUp_userdata
        
End Sub


'�G�N�X�|�[�g�p�̊g���q�擾
'https://learn.microsoft.com/ja-jp/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model
'vbext_ct_StdModule 1 �W�����W���[��
'vbext_ct_ClassModule 2 �N���X ���W���[��
'vbext_ct_MSForm 3 Microsoft Forms
'vbext_ct_ActiveXDesigner 11 ActiveX �f�U�C�i�[
'vbext_ct_Document 100 �h�L�������g ���W���[��
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


'���W���[���t�@�C���̍폜
Private Sub remove_module_files( _
        ByVal dirPath As String)
        
    Dim fso As Object
    Set fso = mUtil.get_fso()
    
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

End Sub


'��è����ۼު�Ď擾
'return: ���̃v���W�F�N�g�I�u�W�F�N�g
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


'�A�N�e�B�u�v���W�F�N�g�̃t�H���_�p�X�擾
'return: ���̃v���W�F�N�g�̃t�H���_�p�X
Private Function get_this_dir_path() As String

    get_this_dir_path = mUtil.get_fso().GetParentFolderName( _
        get_this_project().fileName _
    )

End Function


'�t�H���_���G�N�X�v���[���[�ŊJ��
'param: dirPath-�t�H���_�p�X
Private Sub open_folder_dy_explorer( _
        ByVal dirPath As String)

    On Error Resume Next
        Shell "C:\Windows\Explorer.exe " & dirPath, vbNormalFocus
    On Error GoTo 0

End Sub
        

'���[�U�[�f�[�^�̎擾
'return: dict
Private Function get_user_data() As Object

    Dim path As String
    path = get_this_dir_path() & "\" & USER_JSON_NAME
    
    mUtil.dump "userdata_path:" & path
    
    Dim dict As Object
    If mUtil.get_fso().FileExists(path) Then
        Set dict = load_json(path)
        dict("ver") = VERSION
    Else
        Set dict = mUtil.init_dict()
        dict.Add "ver", VERSION
    End If
    
    Set get_user_data = dict
    
End Function


'���[�U�[�f�[�^�̃N���[���A�b�v
Private Sub cleanUp_userdata()

    Dim fso As Object
    Set fso = mUtil.get_fso()

    Dim removeLst As Collection
    Set removeLst = New Collection
    
    Dim filePath As String
    Dim dirPath As String

    Dim key As Variant 'String
    For Each key In mUserDict.keys()
        If key = "ver" Then GoTo continue

        filePath = Replace(key, SEPARATOR, "\")
        If Not fso.FileExists(filePath) Then
            removeLst.Add key
            GoTo continue
        End If
        
        dirPath = mUserDict(key)("path")
        If Not fso.FolderExists(dirPath) Then
            removeLst.Add key
            GoTo continue
        End If
continue:
    Next
    
    For Each key In removeLst
        mUserDict.Remove key
    Next
    
End Sub


'�v���W�F�N�g���ߋ��ɕۑ�����Ă��邩
'param: projectIdx-�v���W�F�N�g�̃C���f�b�N�X
Function has_user_data( _
        ByVal projectIdx As Long) As Boolean

    has_user_data = False

    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.Item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    
    has_user_data = mUserDict.Exists(Replace(proj.fileName, "\", SEPARATOR))

End Function


'json�̏�������
'param: path-json�t�@�C���p�X
'param: ��������dict
Private Sub save_json( _
        ByVal path As String, _
        ByVal dict As Object)
    
    Call mUtil.write_file( _
        path, _
        mJson.ConvertToJson(dict, 3) _
    )
    
End Sub


'json�̓ǂݍ���
'param: path-json�t�@�C���p�X
'return: dict
Private Function load_json( _
        ByVal path As String) As Object
    
    Set load_json = Nothing
    
    If Not mUtil.get_fso().FileExists(path) Then Exit Function
    
    Dim dict As Object
    Set dict = mJson.ParseJson(mUtil.read_file(path))
    
    Set load_json = dict
    
End Function
