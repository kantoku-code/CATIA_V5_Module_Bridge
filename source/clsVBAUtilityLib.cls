VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAUtilityLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/
'/_/ clsVbaUtilityLib
'/_/ 主に配列、辞書
'/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Option Explicit

'Win API
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Const VER = "0.0.16"

'デバッグモード
Private m_DebugMode As Boolean


Private Sub Class_Initialize()
    m_DebugMode = True
End Sub


Private Sub Class_Terminate()

End Sub


Property Get DebugMode() As Boolean
    DebugMode = m_DebugMode
End Property


Property Let DebugMode(ByVal mode As Boolean)
    m_DebugMode = mode
End Property


Function GetVer() As String
    GetVer = VER
End Function


'ダンプ
'param:string
'return:
Sub Dump( _
        ByVal Txt As String)
    
    If Not m_DebugMode Then Exit Sub
    Debug.Print Txt
End Sub


'変数のダンプ
'param: v-変数
'param: Txt(opt): ヘッダーコメント
'return:
Sub DumpV( _
        ByVal X As Variant, _
        Optional ByVal Txt As String = "")
    
    If Not m_DebugMode Then Exit Sub
    
    Debug.Print Txt & DumpVariable(X)

End Sub


''' @param x As Variant
''' @return As String
Function DumpVariable( _
        ByVal X As Variant) As String

    On Error Resume Next

    Dim ty As String: ty = TypeName(X)
    Select Case ty
        Case "Boolean":     DumpVariable = X
        Case "Integer":     DumpVariable = X & "%"
        Case "Long":        DumpVariable = X & "&"
        #If Win64 Then
        Case "LongLong":    DumpVariable = X & "^"
        #End If
        Case "Single":      DumpVariable = X & "!"
        Case "Double":      DumpVariable = X & "#"
        Case "Currency":    DumpVariable = X & "@"
        Case "Byte":        DumpVariable = "CByte(" & X & ")"
        Case "Decimal":     DumpVariable = "CDec(" & X & ")"
        Case "Date":
            Dim d As String, t As String
            If Abs(X) >= 1 Then d = Month(X) & "/" & Day(X) & "/" & Year(X)
            If Not IsInt(X) Then t = Format(X, "h:nn:ss AM/PM")
            DumpVariable = "#" & Trim(d & " " & t) & "#"
        Case "String"
            If StrPtr(X) = 0 Then
                DumpVariable = "vbNullString"
            Else
                DumpVariable = """" & Replace(X, """", """""") & """"
            End If
        Case "Error"
            If IsMissing(X) Then
                DumpVariable = "Missing"
            Else
                DumpVariable = "CVErr(" & ReMatch(CStr(X), "\d+")(0) & ")"
            End If
        Case "ErrObject"
            DumpVariable = "Err " & X.Number
        Case "Empty", "Null", "Nothing", "Unknown"
            DumpVariable = ty
        Case Else
            If IsObject(X) Then
                Select Case ty
                    Case "Dictionary"
                        Dim keys As Variant
                        keys = X.keys()
                        Dim ar2 As Variant
                        ReDim ar2(X.count - 1)
                        Dim j As Long
                        For j = 0 To X.count - 1: ar2(j) = DumpVariable(keys(j)) & ":" & DumpVariable(X(keys(j))): Next
                        DumpVariable = "Dictionary(" & Join(ar2, ", ") & ")"
                    Case "Collection"
                        Dim ar3 As Variant
                        ReDim ar3(X.count - 1)
                        Dim k As Long
                        For k = 1 To X.count: ar3(k - 1) = DumpVariable(X.Item(k)): Next
                        DumpVariable = "Collection(" & Join(ar3, ", ") & ")"
                    Case Else
                        DumpVariable = ToStr(X)
                End Select
            ElseIf IsArray(X) Then
                Dim rnk As Integer: rnk = ArrRank(X)
                If rnk = 1 Then
                    Dim lb As Long: lb = LBound(X)
                    Dim ub As Long: ub = UBound(X)
                    Dim ar As Variant
                    If ub - lb < 0 Then
                        ar = Array()
                    Else
                        Dim mx As Long: mx = 8 - 1
                        Dim xb As Long: xb = IIf(ub - lb < mx, ub, lb + mx)
                        ReDim ar(lb To xb)
                        Dim i As Long
                        For i = lb To xb: ar(i) = DumpVariable(X(i)): Next
                    End If
                    DumpVariable = "Array(" & Join(ar, ", ") & IIf(xb < ub, ", ...", "") & ")"
                Else
                    DumpVariable = Replace(ty, "()", "(" & String(rnk - 1, ",") & ")")
                End If
            Else
                Err.Raise 51
        End If
    End Select
    
    On Error GoTo 0
    
End Function


''' @param num As Variant(Of Numeric Or Date)
''' @return As Boolean
Function IsInt( _
        ByVal num As Variant) As Boolean

    If IsDate(num) Then num = CDbl(num)
    If Not IsNumeric(num) Then Err.Raise 13
    
    IsInt = num = Fix(num)

End Function


''' @param expr As String
''' @param ptrnFind As String
''' @param iCase As Boolean
''' @return As Variant(Of Array(Of String))
Private Function ReMatch( _
        ByVal expr As String, ByVal ptrnFind As String, _
        Optional ByVal iCase As Boolean = False) As Variant
    
    Dim ret As Variant: ret = Array()
    
    Dim regex As Object:  Set regex = CreateRegExp(ptrnFind, IIf(iCase, "i", ""))
    Dim ms As Object:     Set ms = regex.Execute(expr)
    If ms.count < 1 Then: GoTo Ending
    
    Dim sms As Object:    Set sms = ms(0).SubMatches
    ReDim ret(sms.count)
    
    ret(0) = ms.Item(0).value
    Dim i As Integer
    For i = 1 To UBound(ret): ret(i) = sms.Item(i - 1): Next
    
Ending:
    ReMatch = ret
End Function


''' @param x As Variant
''' @return As String
Private Function ToStr(ByVal X As Variant) As String
    If IsObject(X) Then
        On Error GoTo Err438
        ToStr = X.ToStr()
        On Error GoTo 0
    ElseIf IsArray(X) Then
        ToStr = TypeName(X)
    Else
        ToStr = X
    End If
    
    GoTo Escape
    
Err438:
    Dim e As ErrObject: Set e = Err
    Select Case e.Number
        Case 438: ToStr = TypeName(X): Resume Next
        Case Else: Err.Raise e.Number, e.source, e.Description, e.HelpFile, e.HelpContext
    End Select
    
Escape:
End Function


''' @param arr As Variant(Of Array(Of T))
''' @return As Integer
Private Function ArrRank( _
        ByVal arr As Variant) As Integer
        
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim X As Long
    Dim i As Integer: i = 0
    On Error Resume Next
    While Err.Number = 0: X = UBound(arr, IncrPre(i)): Wend
    ArrRank = i - 1

End Function


''' @param n As Variant
''' @param stepVal As Variant
''' @return As Variant
Private Function IncrPre( _
        ByRef n As Variant, _
        Optional ByVal stepVal As Variant = 1) As Variant

    n = n + stepVal: IncrPre = n

End Function


'*****ｼｽﾃﾑな関数*****
'APC取得
'param:
'return:Object(Of IApc)
Function GetApc() As Object
    Dim COMObjectName As String
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #End If
    
    Dim oApc As Object
    On Error Resume Next
    Set oApc = CreateObject(COMObjectName)
    On Error GoTo 0
    
    If oApc Is Nothing Then
        Set oApc = CreateObject("MSAPC.Apc")
    End If
    
    Set GetApc = oApc
End Function


'VBE取得
'param:
'return:Object(Of IApc)
Function GetVbe() As Object
    Set GetVbe = GetApc().VBE
End Function


'実行中のvbprojectのファルダ
'param:
'return:string
Function GetThisDir() As String
    Dim oApc As Object
    Set oApc = GetApc()

    Dim projFilePath As String
    projFilePath = oApc.ExecutingProject.VBProject.Filename
    
    GetThisDir = GetFso.GetParentFolderName(projFilePath)
    
End Function


'Scripting.Dictionary
'param:compareMode-Long
'return:Object(Of Dictionary)
Function InitDict( _
        Optional compareMode As Long = vbBinaryCompare) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.compareMode = compareMode
    
    Set InitDict = dict
End Function


'文字の配列をキーとし、数をバリューとした辞書作成
'param:txtArray-array(Of string)
'param:compareMode-Long
'return:Object(Of Dictionary)
Function InitDictByArrayCount( _
        ByVal txtArray As Variant, _
        Optional compareMode As Long = vbBinaryCompare) As Object

    Dim dict As Object
    Set dict = InitDict(compareMode)

    Dim Key As Variant
    For Each Key In txtArray
        If Not dict.exists(Key) Then
            dict(Key) = dict(Key) + 1
        Else
            Call dict.Add(Key, 0)
        End If
    Next

    Set InitDictByArrayCount = dict
End Function


'同一の辞書か - キーとバリューで比較　階層でもOK
'param:dict1-Object(Of Dictionary)
'param:dict2-Object(Of Dictionary
'return:boolean
Function IsEqualDict( _
        ByVal dict1 As Object, _
        ByVal dict2 As Object) As Boolean

    If dict1.count <> dict2.count Then
        IsEqualDict = False
        Exit Function
    End If
    
    Dim Key As Variant
    For Each Key In dict1.keys
        If Not dict2.exists(Key) Then
            IsEqualDict = False
            Exit Function
        End If
        If TypeName(dict1(Key)) = "Dictionary" And TypeName(dict2(Key)) = "Dictionary" Then
            ' If both values are dictionaries, recursively compare them
            If Not IsEqualDict(dict1(Key), dict2(Key)) Then
                IsEqualDict = False
                Exit Function
            End If
        ElseIf dict1(Key) <> dict2(Key) Then
            IsEqualDict = False
            Exit Function
        End If
    Next Key

    IsEqualDict = True
    
End Function


'辞書のディープコピー
'param:source-Object(Of Dictionary)
'return:Object(Of Dictionary)
Function DeepCopyDict( _
        ByVal source As Variant) As Object

    Dim Key As Variant
    Dim newDict As Object
    Dim value As Variant
    Set newDict = CreateObject("Scripting.Dictionary")
    For Each Key In source.keys
        value = source.Item(Key)
        If TypeName(value) = "Dictionary" Then
            Set value = DeepCopyDict(value)
        End If

        newDict.Add Key, value
    Next Key

    Set DeepCopyDict = newDict

End Function


'*****配列な関数*****
'サイズ指定の配列作成 - 初期値はletのみ
'param:count-long
'param:value
'return:Variant(Of Array)
Function InitArray( _
        ByVal count As Long, _
        Optional ByVal value As Variant = Empty) As Variant

    Dim ary() As Variant
    If count < 0 Then
        ary = Array()
    Else
        ReDim ary(count)
    End If
    
    If IsEmpty(value) Then
        InitArray = ary
    End If
    
    Dim i As Long
    For i = 0 To UBound(ary)
        ary(i) = value
    Next
    InitArray = ary

End Function


'連結した配列の作成 - 連結じゃなく新たな配列作成
'param:ary1-Variant(Of Array)
'param:ary2-Variant(Of Array)
'return:Variant(Of Array)
Function JointArray( _
        ByVal ary1 As Variant, _
        ByVal ary2 As Variant) As Variant

    Select Case True
        Case Not IsArray(ary1) And Not IsArray(ary2)
            JointArray = Array() 'Empty
            Exit Function
        Case Not IsArray(ary1)
            JointArray = ary2
            Exit Function
        Case Not IsArray(ary2)
            JointArray = ary1
            Exit Function
        Case UBound(ary1) < 0 And UBound(ary2) > -1
            JointArray = ary2
            Exit Function
        Case UBound(ary1) > -1 And UBound(ary2) < 0
            JointArray = ary1
            Exit Function
        Case UBound(ary1) < 0 And UBound(ary2) < 0
            JointArray = Array() 'Empty
            Exit Function
        End Select
        
        Dim StCount&: StCount = UBound(ary1)
        ReDim Preserve ary1(UBound(ary1) + UBound(ary2) + 1)
        
        Dim i&
        If IsObject(ary2(0)) Then
            For i = StCount + 1 To UBound(ary1)
                Set ary1(i) = ary2(i - StCount - 1)
            Next
        Else
            For i = StCount + 1 To UBound(ary1)
                ary1(i) = ary2(i - StCount - 1)
            Next
        End If
        
        JointArray = ary1

End Function


'スライス
'param:ary-Variant(Of Array)
'param:startIdx-Long
'param:endIdx-Long
'return:Variant(Of Array)
Function Slice( _
        ByVal ary As Variant, _
        ByVal startIdx As Long, _
        ByVal endIdx As Long) As Variant

    If Not IsArray(ary) Then Exit Function
    If endIdx - startIdx < 0 Then Exit Function
    If startIdx < 0 Then Exit Function
    If endIdx > UBound(ary) Then Exit Function
    
    Dim rngAry() As Variant
    ReDim rngAry(endIdx - startIdx)

    Dim i As Long
    For i = startIdx To endIdx
        rngAry(i - startIdx) = ary(i)
    Next
    
    Slice = rngAry

End Function


'配列のディープコピー
'param:ary-Variant(Of Array)
'return:Variant(Of Array)
Function DeepCopyArray( _
        ByVal ary As Variant) As Variant

    If Not IsArray(ary) Then Exit Function

    DeepCopyArray = Slice(ary, 0, UBound(ary))

End Function


'配列が一致しているか？
'param:ary1-Variant(Of Array)
'param:ary2-Variant(Of Array)
'return:Boolean
Function IsEqualArray( _
        ByVal ary1 As Variant, _
        ByVal ary2 As Variant) As Boolean

    IsEqualArray = False
    
    If Not IsArray(ary1) Or Not IsArray(ary2) Then Exit Function
    If Not UBound(ary1) = UBound(ary2) Then Exit Function
    
    Dim i As Long
    For i = 0 To UBound(ary1)
        If Not ary1(i) = ary2(i) Then Exit Function
    Next
    
    IsEqualArray = True

End Function


'コレクションの配列化 - let/setは最初の要素で判断
'コレクション内にLetとSetの要素を混ぜないで!!
'param:lst-collection
'return:Variant(Of Array)
Function CollectionToArray( _
        ByVal lst As Variant) As Variant

    If lst.count < 1 Then
        CollectionToArray = Array()
        Exit Function
    End If
    
    If IsObject(lst.Item(1)) Then
        CollectionToArray = CollectionToArraySet(lst)
    Else
        CollectionToArray = CollectionToArrayLet(lst)
    End If

End Function


'-コレクションの配列化-set
'param:lst-collection
'return:Variant(Of Array)
Function CollectionToArraySet( _
        ByVal lst As Variant) As Variant

    Dim lstCount As Long
    lstCount = lst.count

    If lstCount < 1 Then
        CollectionToArraySet = Array()
        Exit Function
    End If

    Dim ary() As Variant
    ReDim ary(lstCount - 1)

    Dim i As Long
    For i = 1 To lst.count
        Set ary(i - 1) = lst(i)
    Next

    CollectionToArraySet = ary

End Function


'-コレクションの配列化-let
'param:lst-collection
'return:Variant(Of Array)
Function CollectionToArrayLet( _
        ByVal lst As Variant) As Variant

    Dim lstCount As Long
    lstCount = lst.count

    If lstCount < 1 Then
        CollectionToArrayLet = Array()
        Exit Function
    End If

    Dim ary() As Variant
    ReDim ary(lstCount - 1)

    Dim i As Long
    For i = 1 To lst.count
        Let ary(i - 1) = lst(i)
    Next

    CollectionToArrayLet = ary

End Function


'配列のコレクション化
'param:Variant(Of Array)
'return:collection
Function ArrayToCollection( _
        ByVal ary As Variant) As Collection

    Dim lst As Collection
    Set lst = New Collection

    If UBound(ary) < 0 Then Exit Function
    
    Dim v As Variant
    For Each v In ary
        lst.Add v
    Next

    Set ArrayToCollection = lst

End Function


'配列のスライス
'param: arr-Variant(Of Array)
'param: StartIndex(opt)-スタートインデックス
'param: EndIndex(opt)-エンドインデックス
'return:- スライスされた配列
Function ArraySlice( _
        ByVal arr As Variant, _
        Optional ByVal StartIndex As Long = 0, _
        Optional ByVal EndIndex As Long = -1) As Variant

    ' EndIndex が省略された場合、配列の最後のインデックスを設定
    If EndIndex = -1 Then
        EndIndex = UBound(arr)
    End If
    
    ' エラーチェック
    If StartIndex < LBound(arr) Or EndIndex > UBound(arr) Or StartIndex > EndIndex Then
        ArraySlice = Array() ' 空の配列を返す
        Exit Function
    End If
    
    ' 結果配列のサイズを設定
    Dim result() As Variant
    ReDim result(0 To EndIndex - StartIndex)
    
    ' 配列をスライス
    Dim j As Long
    j = 0
    
    Dim i As Long
    For i = StartIndex To EndIndex
        If IsObject(arr(0)) Then
            Set result(j) = arr(i)
        Else
            Let result(j) = arr(i)
        End If
        j = j + 1
    Next i
    
    ArraySlice = result

End Function


'************************************************************************
'*  バブルソート
'*  param: arrTmp-値型の配列
'*  return: ソートされた配列
'************************************************************************
Function BubbleSort( _
        ByVal arrTmp As Variant) As Variant

    Dim i As Long, j As Long
    Dim vTmp As Variant

    For i = LBound(arrTmp) To UBound(arrTmp) - 1
        For j = i + 1 To UBound(arrTmp)
            If ExtractNumericPart(arrTmp(j)) < ExtractNumericPart(arrTmp(i)) Then
                vTmp = arrTmp(i)
                arrTmp(i) = arrTmp(j)
                arrTmp(j) = vTmp
            End If
        Next j
    Next i

    BubbleSort = arrTmp

End Function


'------------------------------------------------------------------------
'-  数値部分の抽出
'-  param: sTxt-ターゲット文字列
'-  return: 数値位置インデックス
'------------------------------------------------------------------------
Private Function ExtractNumericPart( _
        ByVal sTxt As String) As Long

    Dim sNumericPart As String
    Dim i As Long

    For i = 1 To Len(sTxt)
        If IsNumeric(Mid(sTxt, i, 1)) Then
            sNumericPart = sNumericPart & Mid(sTxt, i, 1)
        End If
    Next i

    ExtractNumericPart = CLng(sNumericPart)

End Function


'*****IOな関数*****
'FileSystemObject
'return:Object(Of FileSystemObject)
Function GetFso() As Object
    Set GetFso = CreateObject("Scripting.FileSystemObject")
End Function


'Dirﾊﾟｽ/ﾌｧｲﾙ名/拡張子 分割
'param:FullPath-ファイルパス
'return:Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
Function SplitPathName( _
        ByVal fullpath As String) As Variant

    Dim path(2) As Variant ' String
    With GetFso
        path(0) = .GetParentFolderName(fullpath)
        path(1) = .GetBaseName(fullpath)
        path(2) = .GetExtensionName(fullpath)
    End With

    SplitPathName = path

End Function


'Dirﾊﾟｽ/ﾌｧｲﾙ名/拡張子 連結
'param:Path-Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
'return:ファイルパス
Function JointPathName( _
        ByVal path As Variant) As String

    JointPathName = ""

    If Not IsArray(path) Then Exit Function '未対応
    If Not UBound(path) = 2 Then Exit Function '未対応
    
    JointPathName = GetFso().BuildPath(path(0), path(1) & "." & path(2))

End Function


'重複しないパスの取得
'重複する場合 (ファイルベース名)_(ナンバリング) とする
'param:Path-ファイルパス
'return:重複無しパス
Function GetUniquePath( _
        ByVal oldPath As String) As String
        
    Dim path() As Variant
    path = SplitPathName(oldPath)
    path(2) = "." & path(2)

    'ファイル名に使用出来ない文字の置き換え
    path(1) = ReplaceInvalidChars(path(1))

    Dim fso As Object
    Set fso = GetFso()

    Dim newPath As String
    newPath = path(0) + "\" + path(1)
    If Not fso.FileExists(newPath + path(2)) Then
        GetUniquePath = newPath + path(2)
        Exit Function
    End If

    Dim tempName As String
    Dim i As Long
    i = 0
    Do
        i = i + 1
        tempName = newPath + "_" + CStr(i) + path(2)
        If Not fso.FileExists(tempName) Then
            GetUniquePath = tempName
            Exit Function
        End If
    Loop

End Function


'************************************************************************
'*  ファイル名に使用できない文字を[_]に置き換える
'*  param: sPath-ファイル名文字列
'*  return: ファイル名に使用出来る文字列
'************************************************************************
Function ReplaceInvalidChars( _
        ByVal sPath As String) As String

    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    oRegex.Pattern = "[\/:*?""<>|]"
    oRegex.Global = True
    
    ReplaceInvalidChars = oRegex.Replace(sPath, "_")

End Function


'ファイルの読み込み
'1行で読み込んでいるので改行でsplitして
'param:path-string
'return:string
Function ReadFile( _
        ByVal path As String) As String

    ReadFile = ""

    On Error Resume Next

    With GetFso
        ReadFile = .OpenTextFile(path).ReadAll
        .Close
    End With

    On Error GoTo 0

End Function


'ファイルの書き出し
'param:path-string
'param:txt-string
'param:IsAppending(opt)-追記の場合True
Sub WriteFile( _
        ByVal path As String, _
        ByVal Txt As String, _
        Optional ByVal IsAppending = False)

    Dim iIoMode As Integer
    iIoMode = IIf(IsAppending, 8, 2)

    With GetFso().OpenTextFile(path, iIoMode, True)
        On Error Resume Next
        .Write Txt
        If Not Err.Number = 0 Then
            .Write "*** Write Errar ***"
        End If
        On Error GoTo 0
        .Close
    End With

End Sub


'スリープ
'param: dMillisecond-ミリ秒
Sub Wait( _
        ByVal dMillisecond As Double)

    Sleep dMillisecond

End Sub

