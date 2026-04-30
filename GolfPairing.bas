Attribute VB_Name = "GolfPairing"
Option Explicit

' ============================================================
' ゴルフコンペ 組み合わせ自動作成マクロ
' ============================================================
' 【使い方】
'   1. このファイルをExcelのVBAエディタにインポート
'   2. まず CreateSampleData を実行してサンプルデータを作成
'   3. 「参加者選択」シートで参加者の「参加」列に○を入力
'   4. 「組み合わせ作成」ボタンをクリック
' ============================================================

' --- シート名定数 ---
Private Const SH_MASTER  As String = "会員マスタ"
Private Const SH_NG      As String = "NGリスト"
Private Const SH_SELECT  As String = "参加者選択"
Private Const SH_HISTORY As String = "組み合わせ履歴"
Private Const SH_RESULT  As String = "組み合わせ結果"

' --- 動作パラメータ定数 ---
Private Const DEFAULT_GROUP_SIZE As Integer = 4
Private Const MAX_RETRY          As Integer = 2000

' ============================================================
' ユーティリティ：シート存在確認
' ============================================================
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

' ============================================================
' 初期設定：参加者選択シートを作成・更新する
' ============================================================
Sub SetupSelectionSheet()
    If Not SheetExists(SH_MASTER) Then
        MsgBox "「" & SH_MASTER & "」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    Dim wsMaster As Worksheet
    Dim wsSelect As Worksheet
    Set wsMaster = Sheets(SH_MASTER)

    If Not SheetExists(SH_SELECT) Then
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = SH_SELECT
    End If
    Set wsSelect = Sheets(SH_SELECT)

    ' 既存の参加チェック状態を退避
    Dim existingChecks As Object
    Set existingChecks = CreateObject("Scripting.Dictionary")
    Dim exRow As Long
    exRow = wsSelect.Cells(Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    If exRow > 1 Then
        For i = 2 To exRow
            Dim mNo As String
            mNo = CStr(wsSelect.Cells(i, 2).Value)
            If mNo <> "" Then
                existingChecks(mNo) = wsSelect.Cells(i, 1).Value
            End If
        Next i
    End If

    wsSelect.Cells.Clear

    ' ヘッダー
    With wsSelect
        .Cells(1, 1).Value = "参加"
        .Cells(1, 2).Value = "会員番号"
        .Cells(1, 3).Value = "企業名"
        .Cells(1, 4).Value = "役職名"
        .Cells(1, 5).Value = "漢字氏名"
        .Cells(1, 6).Value = "カナ氏名"
        With .Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With
        .Columns(1).ColumnWidth = 6
        .Columns(2).ColumnWidth = 12
        .Columns(3).ColumnWidth = 22
        .Columns(4).ColumnWidth = 14
        .Columns(5).ColumnWidth = 14
        .Columns(6).ColumnWidth = 20
    End With

    ' 会員マスタからデータ転記
    Dim lastRow As Long
    lastRow = wsMaster.Cells(Rows.Count, 1).End(xlUp).Row
    Dim row As Long
    row = 2
    For i = 2 To lastRow
        If wsMaster.Cells(i, 1).Value <> "" Then
            Dim memberNo As String
            memberNo = CStr(wsMaster.Cells(i, 1).Value)
            wsSelect.Cells(row, 2).Value = wsMaster.Cells(i, 1).Value
            wsSelect.Cells(row, 3).Value = wsMaster.Cells(i, 2).Value
            wsSelect.Cells(row, 4).Value = wsMaster.Cells(i, 3).Value
            wsSelect.Cells(row, 5).Value = wsMaster.Cells(i, 4).Value
            wsSelect.Cells(row, 6).Value = wsMaster.Cells(i, 5).Value
            If existingChecks.Exists(memberNo) Then
                wsSelect.Cells(row, 1).Value = existingChecks(memberNo)
            End If
            If row Mod 2 = 0 Then
                wsSelect.Rows(row).Interior.Color = RGB(242, 242, 242)
            End If
            row = row + 1
        End If
    Next i

    ' 操作パラメータ欄
    With wsSelect
        .Cells(1, 8).Value = "【操作方法】"
        .Cells(2, 8).Value = "1. A列に ○ を入力して参加者を選択"
        .Cells(3, 8).Value = "2. 1組あたり人数を設定（デフォルト4）"
        .Cells(4, 8).Value = "3. 参照する回を入力（空欄=最新回）"
        .Cells(5, 8).Value = "4. 「組み合わせ作成」ボタンを押す"
        .Cells(1, 8).Font.Bold = True
        .Cells(7, 8).Value = "1組あたり人数："
        .Cells(7, 9).Value = DEFAULT_GROUP_SIZE
        .Cells(8, 8).Value = "参照する回（空欄=最新）："
        .Cells(8, 9).Value = ""
        .Cells(7, 8).Font.Bold = True
        .Cells(8, 8).Font.Bold = True
        .Columns(8).ColumnWidth = 28
        .Columns(9).ColumnWidth = 10
    End With

    ' ボタン配置
    Dim btn As Button
    On Error Resume Next
    wsSelect.Buttons.Delete
    On Error GoTo 0

    Set btn = wsSelect.Buttons.Add( _
        wsSelect.Cells(10, 8).Left, wsSelect.Cells(10, 8).Top, 160, 28)
    btn.Caption = "組み合わせ作成"
    btn.OnAction = "CreatePairing"

    Set btn = wsSelect.Buttons.Add( _
        wsSelect.Cells(13, 8).Left, wsSelect.Cells(13, 8).Top, 160, 28)
    btn.Caption = "全員選択"
    btn.OnAction = "SelectAll"

    Set btn = wsSelect.Buttons.Add( _
        wsSelect.Cells(15, 8).Left, wsSelect.Cells(15, 8).Top, 160, 28)
    btn.Caption = "選択解除"
    btn.OnAction = "DeselectAll"

    MsgBox "参加者選択シートを更新しました。" & vbCrLf & _
           "「参加」列に ○ を入力して参加者を選択してください。", vbInformation
End Sub

' ============================================================
' 全員選択 / 選択解除
' ============================================================
Sub SelectAll()
    If Not SheetExists(SH_SELECT) Then Exit Sub
    Dim ws As Worksheet
    Set ws = Sheets(SH_SELECT)
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value <> "" Then
            ws.Cells(i, 1).Value = "○"
        End If
    Next i
End Sub

Sub DeselectAll()
    If Not SheetExists(SH_SELECT) Then Exit Sub
    Dim ws As Worksheet
    Set ws = Sheets(SH_SELECT)
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        ws.Cells(i, 1).Value = ""
    Next i
End Sub

' ============================================================
' 参加者リスト取得（参加者選択シートの○印から）
' ============================================================
Private Function GetParticipants(wsSelect As Worksheet, _
                                  ByRef participants() As String) As Integer
    Dim lastRow As Long
    lastRow = wsSelect.Cells(Rows.Count, 2).End(xlUp).Row

    Dim count As Integer
    count = 0
    Dim i As Long
    Dim chk As String

    ' まずカウント
    For i = 2 To lastRow
        chk = Trim(CStr(wsSelect.Cells(i, 1).Value))
        If (chk = "○" Or chk = "〇" Or chk = "o" Or chk = "O") Then
            If wsSelect.Cells(i, 5).Value <> "" Then count = count + 1
        End If
    Next i

    If count = 0 Then
        GetParticipants = 0
        Exit Function
    End If

    ReDim participants(0 To count - 1)
    Dim idx As Integer
    idx = 0
    For i = 2 To lastRow
        chk = Trim(CStr(wsSelect.Cells(i, 1).Value))
        If (chk = "○" Or chk = "〇" Or chk = "o" Or chk = "O") Then
            If wsSelect.Cells(i, 5).Value <> "" Then
                participants(idx) = CStr(wsSelect.Cells(i, 5).Value)
                idx = idx + 1
            End If
        End If
    Next i

    GetParticipants = count
End Function

' ============================================================
' NGペア辞書の読み込み（双方向化）
' ============================================================
Private Function LoadNGPairs() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Not SheetExists(SH_NG) Then
        Set LoadNGPairs = dict
        Exit Function
    End If

    Dim wsNG As Worksheet
    Set wsNG = Sheets(SH_NG)
    Dim lastRow As Long
    lastRow = wsNG.Cells(Rows.Count, 1).End(xlUp).Row

    Dim i As Long, j As Long
    For i = 2 To lastRow
        Dim name1 As String
        name1 = Trim(CStr(wsNG.Cells(i, 2).Value))
        If name1 = "" Then GoTo NextNGRow

        Dim lastCol As Long
        lastCol = wsNG.Cells(i, Columns.Count).End(xlToLeft).Column

        For j = 3 To lastCol
            Dim name2 As String
            name2 = Trim(CStr(wsNG.Cells(i, j).Value))
            If name2 = "" Then GoTo NextNGCol

            ' name1 → name2 登録
            If Not dict.Exists(name1) Then
                Set dict(name1) = CreateObject("Scripting.Dictionary")
            End If
            dict(name1)(name2) = True

            ' name2 → name1 登録（双方向化）
            If Not dict.Exists(name2) Then
                Set dict(name2) = CreateObject("Scripting.Dictionary")
            End If
            dict(name2)(name1) = True

NextNGCol:
        Next j
NextNGRow:
    Next i

    Set LoadNGPairs = dict
End Function

' ============================================================
' 前回組み合わせデータの読み込み
' requestedRound=-1 のとき最新回を使用
' ============================================================
Private Function LoadPreviousRoundData(requestedRound As Integer, _
                                        ByRef actualRound As Integer) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    actualRound = 0

    If Not SheetExists(SH_HISTORY) Then
        Set LoadPreviousRoundData = dict
        Exit Function
    End If

    Dim wsH As Worksheet
    Set wsH = Sheets(SH_HISTORY)
    Dim lastRow As Long
    lastRow = wsH.Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Set LoadPreviousRoundData = dict
        Exit Function
    End If

    ' 最大回番号を取得
    Dim maxRound As Integer
    maxRound = 0
    Dim i As Long
    Dim rn As Integer
    For i = 2 To lastRow
        On Error Resume Next
        rn = CInt(wsH.Cells(i, 1).Value)
        On Error GoTo 0
        If rn > maxRound Then maxRound = rn
    Next i
    If maxRound = 0 Then
        Set LoadPreviousRoundData = dict
        Exit Function
    End If

    ' 参照回の決定
    If requestedRound <= 0 Or requestedRound > maxRound Then
        actualRound = maxRound
    Else
        actualRound = requestedRound
    End If

    ' 指定回の組番号ごとにメンバーを収集
    Dim grpDict As Object
    Set grpDict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        On Error Resume Next
        rn = CInt(wsH.Cells(i, 1).Value)
        On Error GoTo 0
        If rn = actualRound Then
            Dim grpKey As String
            grpKey = CStr(wsH.Cells(i, 3).Value)
            Dim mName As String
            mName = Trim(CStr(wsH.Cells(i, 5).Value))
            If Not grpDict.Exists(grpKey) Then
                Set grpDict(grpKey) = New Collection
            End If
            grpDict(grpKey).Add mName
        End If
    Next i

    ' 組内ペアを「前回同組辞書」に展開
    Dim gKey As Variant
    For Each gKey In grpDict.Keys
        Dim members As Collection
        Set members = grpDict(gKey)
        Dim ci As Integer, cj As Integer
        For ci = 1 To members.Count
            For cj = 1 To members.Count
                If ci <> cj Then
                    Dim n1 As String, n2 As String
                    n1 = members(ci)
                    n2 = members(cj)
                    If Not dict.Exists(n1) Then
                        Set dict(n1) = CreateObject("Scripting.Dictionary")
                    End If
                    dict(n1)(n2) = True
                End If
            Next cj
        Next ci
    Next gKey

    Set LoadPreviousRoundData = dict
End Function

' ============================================================
' グループサイズ計算
' 最大組と最小組の差が1以内になるよう配分
' ============================================================
Private Sub CalculateGroupSizes(n As Integer, targetSize As Integer, _
                                 ByRef numGroups As Integer, _
                                 ByRef groupSizes() As Integer)
    ' floor(n/target) と ceil(n/(target+1)) の大きい方を組数とする
    Dim fl As Integer
    Dim cl As Integer
    fl = Int(n / targetSize)
    If fl < 1 Then fl = 1
    cl = Int((n + targetSize) / (targetSize + 1))
    If cl < 1 Then cl = 1
    numGroups = IIf(fl > cl, fl, cl)

    ReDim groupSizes(0 To numGroups - 1)
    Dim baseSize As Integer
    Dim remainder As Integer
    baseSize = Int(n / numGroups)
    remainder = n Mod numGroups

    Dim i As Integer
    For i = 0 To numGroups - 1
        ' remainder 個の組だけ 1 人多くする
        groupSizes(i) = IIf(i < remainder, baseSize + 1, baseSize)
    Next i
End Sub

' ============================================================
' Fisher-Yates シャッフル
' ============================================================
Private Sub ShuffleArray(arr() As Integer, n As Integer)
    Dim i As Integer, j As Integer, tmp As Integer
    For i = n - 1 To 1 Step -1
        j = Int(Rnd() * (i + 1))
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
    Next i
End Sub

' ============================================================
' 次の回番号を取得（履歴の最大値 + 1）
' ============================================================
Private Function GetNextRoundNumber() As Integer
    If Not SheetExists(SH_HISTORY) Then
        GetNextRoundNumber = 1
        Exit Function
    End If
    Dim wsH As Worksheet
    Set wsH = Sheets(SH_HISTORY)
    Dim lastRow As Long
    lastRow = wsH.Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        GetNextRoundNumber = 1
        Exit Function
    End If
    Dim maxRound As Integer
    maxRound = 0
    Dim i As Long
    Dim rn As Integer
    For i = 2 To lastRow
        On Error Resume Next
        rn = CInt(wsH.Cells(i, 1).Value)
        On Error GoTo 0
        If rn > maxRound Then maxRound = rn
    Next i
    GetNextRoundNumber = maxRound + 1
End Function

' ============================================================
' ランダム組み合わせ生成
' 戻り値 True=成功 / False=失敗（NG制約を満たせない）
' groupAssign(i) : 参加者 i が属するグループ番号(0始まり)
' softViolations(i) : True=前回同組制約を緩和して配置
' ============================================================
Private Function RunPairing(participants() As String, n As Integer, _
                             numGroups As Integer, groupSizes() As Integer, _
                             ngDict As Object, prevDict As Object, _
                             ByRef groupAssign() As Integer, _
                             ByRef softViolations() As Boolean) As Boolean
    ReDim groupAssign(0 To n - 1)
    ReDim softViolations(0 To n - 1)

    Dim attempt As Integer
    For attempt = 1 To MAX_RETRY

        ' シャッフル用インデックス配列
        Dim shuffled() As Integer
        ReDim shuffled(0 To n - 1)
        Dim i As Integer
        For i = 0 To n - 1
            shuffled(i) = i
        Next i
        Call ShuffleArray(shuffled, n)

        ' 各グループの現在人数とメンバー辞書
        Dim grpCount() As Integer
        ReDim grpCount(0 To numGroups - 1)
        Dim grpMembers() As Object
        ReDim grpMembers(0 To numGroups - 1)
        Dim g As Integer
        For g = 0 To numGroups - 1
            Set grpMembers(g) = CreateObject("Scripting.Dictionary")
        Next g

        Dim tempAssign() As Integer
        ReDim tempAssign(0 To n - 1)
        Dim failed As Boolean
        failed = False

        For i = 0 To n - 1
            Dim pIdx As Integer
            pIdx = shuffled(i)
            Dim pName As String
            pName = participants(pIdx)

            ' 配置可能グループを収集
            Dim validG() As Integer
            Dim softOKG() As Integer
            Dim vCnt As Integer, sCnt As Integer
            vCnt = 0
            sCnt = 0
            ReDim validG(0 To numGroups - 1)
            ReDim softOKG(0 To numGroups - 1)

            For g = 0 To numGroups - 1
                If grpCount(g) < groupSizes(g) Then
                    ' ハード制約：NG同士でないこと
                    If Not HasNGConflict(pName, g, grpMembers, ngDict) Then
                        validG(vCnt) = g
                        vCnt = vCnt + 1
                        ' ソフト制約：前回同組でないこと
                        If Not HasPrevConflict(pName, g, grpMembers, prevDict) Then
                            softOKG(sCnt) = g
                            sCnt = sCnt + 1
                        End If
                    End If
                End If
            Next g

            If vCnt = 0 Then
                failed = True
                Exit For
            End If

            ' ソフト制約を優先し、なければハード制約のみ満たすグループへ
            Dim chosen As Integer
            If sCnt > 0 Then
                chosen = softOKG(Int(Rnd() * sCnt))
            Else
                chosen = validG(Int(Rnd() * vCnt))
            End If

            tempAssign(pIdx) = chosen
            grpCount(chosen) = grpCount(chosen) + 1
            grpMembers(chosen)(pName) = True
        Next i

        If Not failed Then
            ' ソフト違反フラグを記録してリターン
            For i = 0 To n - 1
                pName = participants(i)
                g = tempAssign(i)
                groupAssign(i) = g
                softViolations(i) = CheckSoftViolation(pName, g, grpMembers, prevDict)
            Next i
            RunPairing = True
            Exit Function
        End If

    Next attempt

    RunPairing = False
End Function

' --- NG衝突チェック（ハード制約）---
Private Function HasNGConflict(pName As String, g As Integer, _
                                grpMembers() As Object, ngDict As Object) As Boolean
    If Not ngDict.Exists(pName) Then
        HasNGConflict = False
        Exit Function
    End If
    Dim ngList As Object
    Set ngList = ngDict(pName)
    Dim m As Variant
    For Each m In grpMembers(g).Keys
        If ngList.Exists(CStr(m)) Then
            HasNGConflict = True
            Exit Function
        End If
    Next m
    HasNGConflict = False
End Function

' --- 前回同組チェック（ソフト制約・配置前）---
Private Function HasPrevConflict(pName As String, g As Integer, _
                                  grpMembers() As Object, prevDict As Object) As Boolean
    If Not prevDict.Exists(pName) Then
        HasPrevConflict = False
        Exit Function
    End If
    Dim prevList As Object
    Set prevList = prevDict(pName)
    Dim m As Variant
    For Each m In grpMembers(g).Keys
        If prevList.Exists(CStr(m)) Then
            HasPrevConflict = True
            Exit Function
        End If
    Next m
    HasPrevConflict = False
End Function

' --- ソフト違反チェック（配置後・自分以外を確認）---
Private Function CheckSoftViolation(pName As String, g As Integer, _
                                     grpMembers() As Object, prevDict As Object) As Boolean
    If Not prevDict.Exists(pName) Then
        CheckSoftViolation = False
        Exit Function
    End If
    Dim prevList As Object
    Set prevList = prevDict(pName)
    Dim m As Variant
    For Each m In grpMembers(g).Keys
        If CStr(m) <> pName Then
            If prevList.Exists(CStr(m)) Then
                CheckSoftViolation = True
                Exit Function
            End If
        End If
    Next m
    CheckSoftViolation = False
End Function

' ============================================================
' 結果シートへ出力
' ============================================================
Private Sub OutputResults(participants() As String, n As Integer, _
                           numGroups As Integer, groupSizes() As Integer, _
                           groupAssign() As Integer, softViolations() As Boolean, _
                           roundNum As Integer, refRound As Integer)
    If Not SheetExists(SH_RESULT) Then
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = SH_RESULT
    End If
    Dim wsR As Worksheet
    Set wsR = Sheets(SH_RESULT)
    wsR.Cells.Clear
    wsR.Cells.Interior.ColorIndex = xlNone
    wsR.Cells.Font.Color = RGB(0, 0, 0)

    ' タイトル
    With wsR.Cells(1, 1)
        .Value = "第" & roundNum & "回 ゴルフコンペ 組み合わせ表"
        .Font.Bold = True
        .Font.Size = 16
    End With
    wsR.Cells(2, 1).Value = "作成日時：" & Format(Now(), "yyyy/mm/dd hh:mm")
    If refRound > 0 Then
        wsR.Cells(3, 1).Value = "参照前回：第" & refRound & "回"
    Else
        wsR.Cells(3, 1).Value = "参照前回：なし（初回）"
    End If
    wsR.Cells(4, 1).Value = "参加者数：" & n & "名 ／ " & numGroups & "組"

    With wsR.Cells(5, 1)
        .Value = "※ オレンジ色 = 前回同組制約を緩和したペア（幹事要確認）"
        .Font.Color = RGB(200, 80, 0)
        .Font.Bold = True
    End With

    ' グループ別メンバーリストを生成（Collection of Array）
    Dim grpList() As Object
    ReDim grpList(0 To numGroups - 1)
    Dim g As Integer
    For g = 0 To numGroups - 1
        Set grpList(g) = New Collection
    Next g
    Dim i As Integer
    For i = 0 To n - 1
        grpList(groupAssign(i)).Add Array(participants(i), softViolations(i))
    Next i

    ' 最大組人数
    Dim maxSize As Integer
    maxSize = 0
    For g = 0 To numGroups - 1
        If groupSizes(g) > maxSize Then maxSize = groupSizes(g)
    Next g

    ' テーブルヘッダー行
    Dim startRow As Integer
    startRow = 7
    wsR.Cells(startRow, 1).Value = "組"
    Dim j As Integer
    For j = 1 To maxSize
        wsR.Cells(startRow, 1 + j).Value = j & "番"
    Next j
    wsR.Cells(startRow, 2 + maxSize).Value = "人数"
    With wsR.Rows(startRow)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' データ行
    For g = 0 To numGroups - 1
        Dim row As Integer
        row = startRow + 1 + g
        wsR.Cells(row, 1).Value = "第" & (g + 1) & "組"

        Dim mIdx As Integer
        mIdx = 1
        Dim item As Variant
        For Each item In grpList(g)
            Dim mName As String
            Dim isSoft As Boolean
            mName = item(0)
            isSoft = item(1)
            Dim cel As Range
            Set cel = wsR.Cells(row, 1 + mIdx)
            cel.Value = mName
            If isSoft Then
                cel.Interior.Color = RGB(255, 192, 0)
                cel.Font.Color = RGB(150, 50, 0)
            End If
            mIdx = mIdx + 1
        Next item

        wsR.Cells(row, 2 + maxSize).Value = grpList(g).Count

        ' 奇数グループは薄グレー（ソフト違反セルは上書きしない）
        If g Mod 2 = 1 Then
            Dim c As Integer
            For c = 1 To 2 + maxSize
                If wsR.Cells(row, c).Interior.Color <> RGB(255, 192, 0) Then
                    wsR.Cells(row, c).Interior.Color = RGB(242, 242, 242)
                End If
            Next c
        End If
    Next g

    wsR.Columns.AutoFit
    wsR.Activate
    wsR.Range("A1").Select
End Sub

' ============================================================
' 履歴シートへ保存
' ============================================================
Private Sub SaveToHistory(participants() As String, n As Integer, _
                           groupAssign() As Integer, roundNum As Integer)
    ' 履歴シートがなければ作成
    If Not SheetExists(SH_HISTORY) Then
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = SH_HISTORY
        With Sheets(SH_HISTORY)
            .Cells(1, 1).Value = "回数"
            .Cells(1, 2).Value = "開催日"
            .Cells(1, 3).Value = "組番号"
            .Cells(1, 4).Value = "会員番号"
            .Cells(1, 5).Value = "漢字氏名"
            With .Rows(1)
                .Font.Bold = True
                .Interior.Color = RGB(68, 114, 196)
                .Font.Color = RGB(255, 255, 255)
            End With
        End With
    End If

    ' 開催日を入力
    Dim dateStr As String
    dateStr = InputBox("開催日を入力してください（例：2024/05/15）", _
                       "開催日入力", Format(Date, "yyyy/mm/dd"))
    If dateStr = "" Then dateStr = Format(Date, "yyyy/mm/dd")

    ' 会員番号辞書
    Dim mNoDict As Object
    Set mNoDict = GetMemberNumbers()

    Dim wsH As Worksheet
    Set wsH = Sheets(SH_HISTORY)
    Dim nextRow As Long
    nextRow = wsH.Cells(Rows.Count, 1).End(xlUp).Row + 1

    Dim i As Integer
    For i = 0 To n - 1
        wsH.Cells(nextRow, 1).Value = roundNum
        wsH.Cells(nextRow, 2).Value = dateStr
        wsH.Cells(nextRow, 3).Value = groupAssign(i) + 1
        Dim pName As String
        pName = participants(i)
        If mNoDict.Exists(pName) Then
            wsH.Cells(nextRow, 4).Value = mNoDict(pName)
        End If
        wsH.Cells(nextRow, 5).Value = pName
        nextRow = nextRow + 1
    Next i

    wsH.Columns.AutoFit
End Sub

' ============================================================
' 会員番号辞書（氏名→会員番号）
' ============================================================
Private Function GetMemberNumbers() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If Not SheetExists(SH_SELECT) Then
        Set GetMemberNumbers = dict
        Exit Function
    End If
    Dim ws As Worksheet
    Set ws = Sheets(SH_SELECT)
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        Dim nm As String
        nm = Trim(CStr(ws.Cells(i, 5).Value))
        If nm <> "" Then
            dict(nm) = ws.Cells(i, 2).Value
        End If
    Next i
    Set GetMemberNumbers = dict
End Function

' ============================================================
' メイン：組み合わせ作成（ボタンから呼び出し）
' ============================================================
Sub CreatePairing()
    Application.ScreenUpdating = False
    Randomize

    ' シート確認
    If Not SheetExists(SH_SELECT) Then
        MsgBox "参加者選択シートがありません。" & vbCrLf & _
               "「SetupSelectionSheet」を先に実行してください。", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim wsSelect As Worksheet
    Set wsSelect = Sheets(SH_SELECT)

    ' 参加者リスト取得
    Dim participants() As String
    Dim n As Integer
    n = GetParticipants(wsSelect, participants)
    If n < 2 Then
        MsgBox "参加者が少なすぎます（最低2名）。" & vbCrLf & _
               "「参加」列に ○ を入力してください。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' パラメータ取得
    Dim groupSize As Integer
    Dim refRoundInput As Integer
    On Error Resume Next
    groupSize = CInt(wsSelect.Cells(7, 9).Value)
    If Err.Number <> 0 Or groupSize < 2 Then groupSize = DEFAULT_GROUP_SIZE
    Err.Clear
    Dim refCell As String
    refCell = Trim(CStr(wsSelect.Cells(8, 9).Value))
    If refCell = "" Then
        refRoundInput = -1
    Else
        refRoundInput = CInt(refCell)
        If Err.Number <> 0 Then refRoundInput = -1
    End If
    Err.Clear
    On Error GoTo 0

    ' NGペア・前回データ読み込み
    Dim ngDict As Object
    Set ngDict = LoadNGPairs()
    Dim prevDict As Object
    Dim actualRefRound As Integer
    Set prevDict = LoadPreviousRoundData(refRoundInput, actualRefRound)

    ' グループサイズ計算
    Dim numGroups As Integer
    Dim groupSizes() As Integer
    Call CalculateGroupSizes(n, groupSize, numGroups, groupSizes)

    ' 組み合わせ生成
    Dim groupAssign() As Integer
    Dim softViolations() As Boolean
    Dim success As Boolean
    success = RunPairing(participants, n, numGroups, groupSizes, _
                         ngDict, prevDict, groupAssign, softViolations)

    If Not success Then
        MsgBox "組み合わせの生成に失敗しました。" & vbCrLf & vbCrLf & _
               "NG設定が厳しすぎて有効な組み合わせが見つかりません。" & vbCrLf & _
               "参加者やNG設定を見直してください。", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 次の回番号
    Dim nextRound As Integer
    nextRound = GetNextRoundNumber()

    ' 結果出力
    Call OutputResults(participants, n, numGroups, groupSizes, _
                       groupAssign, softViolations, nextRound, actualRefRound)

    Application.ScreenUpdating = True

    ' ソフト違反数カウント
    Dim softCount As Integer
    Dim i As Integer
    For i = 0 To n - 1
        If softViolations(i) Then softCount = softCount + 1
    Next i

    ' 完了メッセージ＆保存確認
    Dim msg As String
    msg = "第" & nextRound & "回の組み合わせを作成しました。" & vbCrLf & vbCrLf
    If actualRefRound > 0 Then
        msg = msg & "参照した前回データ：第" & actualRefRound & "回" & vbCrLf
    End If
    If softCount > 0 Then
        msg = msg & "⚠ 前回同組制約を " & softCount & " 名について緩和しました。" & vbCrLf
        msg = msg & "（結果シートのオレンジ色セルを確認してください）" & vbCrLf
    End If
    msg = msg & vbCrLf & "この結果を履歴に保存しますか？"

    Dim ans As VbMsgBoxResult
    ans = MsgBox(msg, vbYesNo + vbQuestion, "組み合わせ完成")
    If ans = vbYes Then
        Call SaveToHistory(participants, n, groupAssign, nextRound)
        MsgBox "第" & nextRound & "回として履歴に保存しました。", vbInformation
    End If
End Sub

' ============================================================
' サンプルデータ一括作成（初回テスト用）
' ============================================================
Sub CreateSampleData()
    Call CreateSampleMaster
    Call CreateSampleNGList
    Call CreateSampleHistory
    Call SetupSelectionSheet
    MsgBox "サンプルデータの作成が完了しました！" & vbCrLf & vbCrLf & _
           "「参加者選択」シートで参加者に ○ を入力し、" & vbCrLf & _
           "「組み合わせ作成」ボタンを押してください。", vbInformation
End Sub

' --- 会員マスタ（50名）---
Private Sub CreateSampleMaster()
    Dim ws As Worksheet
    If SheetExists(SH_MASTER) Then
        Set ws = Sheets(SH_MASTER)
    Else
        Sheets.Add Before:=Sheets(1)
        ActiveSheet.Name = SH_MASTER
        Set ws = ActiveSheet
    End If
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "会員番号"
    ws.Cells(1, 2).Value = "企業名"
    ws.Cells(1, 3).Value = "役職名"
    ws.Cells(1, 4).Value = "漢字氏名"
    ws.Cells(1, 5).Value = "カナ氏名"
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim r As Integer
    r = 2
    Dim d(1 To 50, 1 To 5) As String

    d(1,1)="M001": d(1,2)="山田商事(株)":   d(1,3)="社長":  d(1,4)="山田太郎":   d(1,5)="ヤマダタロウ"
    d(2,1)="M002": d(2,2)="鈴木工業(株)":   d(2,3)="専務":  d(2,4)="鈴木次郎":   d(2,5)="スズキジロウ"
    d(3,1)="M003": d(3,2)="田中建設(株)":   d(3,3)="常務":  d(3,4)="田中三郎":   d(3,5)="タナカサブロウ"
    d(4,1)="M004": d(4,2)="佐藤電機(株)":   d(4,3)="取締役": d(4,4)="佐藤四郎":  d(4,5)="サトウシロウ"
    d(5,1)="M005": d(5,2)="伊藤食品(株)":   d(5,3)="社長":  d(5,4)="伊藤五郎":   d(5,5)="イトウゴロウ"
    d(6,1)="M006": d(6,2)="渡辺商事(株)":   d(6,3)="会長":  d(6,4)="渡辺六郎":   d(6,5)="ワタナベロクロウ"
    d(7,1)="M007": d(7,2)="中村製造(株)":   d(7,3)="社長":  d(7,4)="中村七郎":   d(7,5)="ナカムラシチロウ"
    d(8,1)="M008": d(8,2)="小林不動産(株)": d(8,3)="専務":  d(8,4)="小林八郎":   d(8,5)="コバヤシハチロウ"
    d(9,1)="M009": d(9,2)="加藤金融(株)":   d(9,3)="取締役": d(9,4)="加藤九郎":  d(9,5)="カトウクロウ"
    d(10,1)="M010": d(10,2)="吉田運輸(株)":  d(10,3)="社長": d(10,4)="吉田十郎":  d(10,5)="ヨシダジュウロウ"
    d(11,1)="M011": d(11,2)="山本医療(株)":  d(11,3)="院長": d(11,4)="山本一男":  d(11,5)="ヤマモトカズオ"
    d(12,1)="M012": d(12,2)="松本建設(株)":  d(12,3)="社長": d(12,4)="松本二男":  d(12,5)="マツモトツギオ"
    d(13,1)="M013": d(13,2)="井上商会":      d(13,3)="代表": d(13,4)="井上三男":  d(13,5)="イノウエミツオ"
    d(14,1)="M014": d(14,2)="木村電気(株)":  d(14,3)="社長": d(14,4)="木村四男":  d(14,5)="キムラヨツオ"
    d(15,1)="M015": d(15,2)="林製薬(株)":    d(15,3)="専務": d(15,4)="林五男":    d(15,5)="ハヤシイツオ"
    d(16,1)="M016": d(16,2)="清水土木(株)":  d(16,3)="常務": d(16,4)="清水六男":  d(16,5)="シミズムツオ"
    d(17,1)="M017": d(17,2)="山崎食品(株)":  d(17,3)="社長": d(17,4)="山崎七男":  d(17,5)="ヤマザキナナオ"
    d(18,1)="M018": d(18,2)="池田金属(株)":  d(18,3)="会長": d(18,4)="池田八男":  d(18,5)="イケダヤツオ"
    d(19,1)="M019": d(19,2)="橋本化学(株)":  d(19,3)="社長": d(19,4)="橋本九男":  d(19,5)="ハシモトコノオ"
    d(20,1)="M020": d(20,2)="阿部商事(株)":  d(20,3)="取締役": d(20,4)="阿部十男": d(20,5)="アベトオオ"
    d(21,1)="M021": d(21,2)="石川機械(株)":  d(21,3)="社長": d(21,4)="石川一郎":  d(21,5)="イシカワイチロウ"
    d(22,1)="M022": d(22,2)="前田建工(株)":  d(22,3)="専務": d(22,4)="前田二郎":  d(22,5)="マエダジロウ"
    d(23,1)="M023": d(23,2)="後藤物産(株)":  d(23,3)="社長": d(23,4)="後藤三郎":  d(23,5)="ゴトウサブロウ"
    d(24,1)="M024": d(24,2)="村上水産(株)":  d(24,3)="代表": d(24,4)="村上四郎":  d(24,5)="ムラカミシロウ"
    d(25,1)="M025": d(25,2)="長谷川観光(株)": d(25,3)="社長": d(25,4)="長谷川五郎": d(25,5)="ハセガワゴロウ"
    d(26,1)="M026": d(26,2)="近藤製鉄(株)":  d(26,3)="会長": d(26,4)="近藤六郎":  d(26,5)="コンドウロクロウ"
    d(27,1)="M027": d(27,2)="藤田商会":      d(27,3)="代表": d(27,4)="藤田七郎":  d(27,5)="フジタシチロウ"
    d(28,1)="M028": d(28,2)="坂本電子(株)":  d(28,3)="社長": d(28,4)="坂本八郎":  d(28,5)="サカモトハチロウ"
    d(29,1)="M029": d(29,2)="岡田印刷(株)":  d(29,3)="専務": d(29,4)="岡田九郎":  d(29,5)="オカダクロウ"
    d(30,1)="M030": d(30,2)="松田自動車(株)": d(30,3)="社長": d(30,4)="松田十郎":  d(30,5)="マツダジュウロウ"
    d(31,1)="M031": d(31,2)="中島航空(株)":  d(31,3)="取締役": d(31,4)="中島一夫": d(31,5)="ナカジマカズオ"
    d(32,1)="M032": d(32,2)="和田保険(株)":  d(32,3)="社長": d(32,4)="和田二夫":  d(32,5)="ワダツギオ"
    d(33,1)="M033": d(33,2)="高橋証券(株)":  d(33,3)="専務": d(33,4)="高橋三夫":  d(33,5)="タカハシミツオ"
    d(34,1)="M034": d(34,2)="浜田観光(株)":  d(34,3)="社長": d(34,4)="浜田四夫":  d(34,5)="ハマダヨツオ"
    d(35,1)="M035": d(35,2)="西村医院":      d(35,3)="院長": d(35,4)="西村五夫":  d(35,5)="ニシムライツオ"
    d(36,1)="M036": d(36,2)="原田商事(株)":  d(36,3)="社長": d(36,4)="原田六夫":  d(36,5)="ハラダムツオ"
    d(37,1)="M037": d(37,2)="内田建設(株)":  d(37,3)="常務": d(37,4)="内田七夫":  d(37,5)="ウチダナナオ"
    d(38,1)="M038": d(38,2)="増田電気(株)":  d(38,3)="社長": d(38,4)="増田八夫":  d(38,5)="マスダヤツオ"
    d(39,1)="M039": d(39,2)="市川製造(株)":  d(39,3)="取締役": d(39,4)="市川九夫": d(39,5)="イチカワコノオ"
    d(40,1)="M040": d(40,2)="千葉商会":      d(40,3)="代表": d(40,4)="千葉十夫":  d(40,5)="チバトオオ"
    d(41,1)="M041": d(41,2)="秋田農業(株)":  d(41,3)="社長": d(41,4)="秋田一郎":  d(41,5)="アキタイチロウ"
    d(42,1)="M042": d(42,2)="島田水産(株)":  d(42,3)="専務": d(42,4)="島田二郎":  d(42,5)="シマダジロウ"
    d(43,1)="M043": d(43,2)="工藤製薬(株)":  d(43,3)="社長": d(43,4)="工藤三郎":  d(43,5)="クドウサブロウ"
    d(44,1)="M044": d(44,2)="川口電子(株)":  d(44,3)="会長": d(44,4)="川口四郎":  d(44,5)="カワグチシロウ"
    d(45,1)="M045": d(45,2)="太田不動産(株)": d(45,3)="社長": d(45,4)="太田五郎":  d(45,5)="オオタゴロウ"
    d(46,1)="M046": d(46,2)="丸山商事(株)":  d(46,3)="取締役": d(46,4)="丸山六郎": d(46,5)="マルヤマロクロウ"
    d(47,1)="M047": d(47,2)="上田機械(株)":  d(47,3)="社長": d(47,4)="上田七郎":  d(47,5)="ウエダシチロウ"
    d(48,1)="M048": d(48,2)="横田建工(株)":  d(48,3)="専務": d(48,4)="横田八郎":  d(48,5)="ヨコタハチロウ"
    d(49,1)="M049": d(49,2)="宮田物産(株)":  d(49,3)="社長": d(49,4)="宮田九郎":  d(49,5)="ミヤタクロウ"
    d(50,1)="M050": d(50,2)="福田運輸(株)":  d(50,3)="代表": d(50,4)="福田十郎":  d(50,5)="フクダジュウロウ"

    Dim i As Integer
    For i = 1 To 50
        ws.Cells(r, 1).Value = d(i, 1)
        ws.Cells(r, 2).Value = d(i, 2)
        ws.Cells(r, 3).Value = d(i, 3)
        ws.Cells(r, 4).Value = d(i, 4)
        ws.Cells(r, 5).Value = d(i, 5)
        If r Mod 2 = 0 Then ws.Rows(r).Interior.Color = RGB(242, 242, 242)
        r = r + 1
    Next i
    ws.Columns.AutoFit
End Sub

' --- NGリスト（サンプル）---
Private Sub CreateSampleNGList()
    Dim ws As Worksheet
    If SheetExists(SH_NG) Then
        Set ws = Sheets(SH_NG)
    Else
        Sheets.Add After:=Sheets(SH_MASTER)
        ActiveSheet.Name = SH_NG
        Set ws = ActiveSheet
    End If
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "会員番号"
    ws.Cells(1, 2).Value = "漢字氏名"
    ws.Cells(1, 3).Value = "NG会員1"
    ws.Cells(1, 4).Value = "NG会員2"
    ws.Cells(1, 5).Value = "NG会員3"
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' サンプルNGペア
    Dim ng(1 To 12, 1 To 5) As String
    ng(1,1)="M001": ng(1,2)="山田太郎": ng(1,3)="鈴木次郎": ng(1,4)="田中三郎": ng(1,5)=""
    ng(2,1)="M005": ng(2,2)="伊藤五郎": ng(2,3)="渡辺六郎": ng(2,4)="中村七郎": ng(2,5)=""
    ng(3,1)="M010": ng(3,2)="吉田十郎": ng(3,3)="山本一男": ng(3,4)="": ng(3,5)=""
    ng(4,1)="M015": ng(4,2)="林五男":   ng(4,3)="清水六男": ng(4,4)="山崎七男": ng(4,5)=""
    ng(5,1)="M020": ng(5,2)="阿部十男": ng(5,3)="石川一郎": ng(5,4)="": ng(5,5)=""
    ng(6,1)="M025": ng(6,2)="長谷川五郎": ng(6,3)="近藤六郎": ng(6,4)="": ng(6,5)=""
    ng(7,1)="M030": ng(7,2)="松田十郎": ng(7,3)="中島一夫": ng(7,4)="和田二夫": ng(7,5)=""
    ng(8,1)="M035": ng(8,2)="西村五夫": ng(8,3)="原田六夫": ng(8,4)="": ng(8,5)=""
    ng(9,1)="M040": ng(9,2)="千葉十夫": ng(9,3)="秋田一郎": ng(9,4)="島田二郎": ng(9,5)=""
    ng(10,1)="M045": ng(10,2)="太田五郎": ng(10,3)="丸山六郎": ng(10,4)="": ng(10,5)=""
    ng(11,1)="M022": ng(11,2)="前田二郎": ng(11,3)="後藤三郎": ng(11,4)="": ng(11,5)=""
    ng(12,1)="M033": ng(12,2)="高橋三夫": ng(12,3)="浜田四夫": ng(12,4)="": ng(12,5)=""

    Dim i As Integer
    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = ng(i, 1)
        ws.Cells(i + 1, 2).Value = ng(i, 2)
        ws.Cells(i + 1, 3).Value = ng(i, 3)
        ws.Cells(i + 1, 4).Value = ng(i, 4)
        ws.Cells(i + 1, 5).Value = ng(i, 5)
        If (i + 1) Mod 2 = 0 Then ws.Rows(i + 1).Interior.Color = RGB(242, 242, 242)
    Next i
    ws.Columns.AutoFit
End Sub

' --- 組み合わせ履歴（第1回サンプル：24名・6組）---
Private Sub CreateSampleHistory()
    Dim ws As Worksheet
    If SheetExists(SH_HISTORY) Then
        Set ws = Sheets(SH_HISTORY)
    Else
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = SH_HISTORY
        Set ws = ActiveSheet
    End If
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "回数"
    ws.Cells(1, 2).Value = "開催日"
    ws.Cells(1, 3).Value = "組番号"
    ws.Cells(1, 4).Value = "会員番号"
    ws.Cells(1, 5).Value = "漢字氏名"
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim h(1 To 24, 1 To 5) As String
    h(1,1)="1":  h(1,2)="2024/03/15": h(1,3)="1": h(1,4)="M001": h(1,5)="山田太郎"
    h(2,1)="1":  h(2,2)="2024/03/15": h(2,3)="1": h(2,4)="M006": h(2,5)="渡辺六郎"
    h(3,1)="1":  h(3,2)="2024/03/15": h(3,3)="1": h(3,4)="M011": h(3,5)="山本一男"
    h(4,1)="1":  h(4,2)="2024/03/15": h(4,3)="1": h(4,4)="M016": h(4,5)="清水六男"
    h(5,1)="1":  h(5,2)="2024/03/15": h(5,3)="2": h(5,4)="M002": h(5,5)="鈴木次郎"
    h(6,1)="1":  h(6,2)="2024/03/15": h(6,3)="2": h(6,4)="M007": h(6,5)="中村七郎"
    h(7,1)="1":  h(7,2)="2024/03/15": h(7,3)="2": h(7,4)="M012": h(7,5)="松本二男"
    h(8,1)="1":  h(8,2)="2024/03/15": h(8,3)="2": h(8,4)="M017": h(8,5)="山崎七男"
    h(9,1)="1":  h(9,2)="2024/03/15": h(9,3)="3": h(9,4)="M003": h(9,5)="田中三郎"
    h(10,1)="1": h(10,2)="2024/03/15": h(10,3)="3": h(10,4)="M008": h(10,5)="小林八郎"
    h(11,1)="1": h(11,2)="2024/03/15": h(11,3)="3": h(11,4)="M013": h(11,5)="井上三男"
    h(12,1)="1": h(12,2)="2024/03/15": h(12,3)="3": h(12,4)="M018": h(12,5)="池田八男"
    h(13,1)="1": h(13,2)="2024/03/15": h(13,3)="4": h(13,4)="M004": h(13,5)="佐藤四郎"
    h(14,1)="1": h(14,2)="2024/03/15": h(14,3)="4": h(14,4)="M009": h(14,5)="加藤九郎"
    h(15,1)="1": h(15,2)="2024/03/15": h(15,3)="4": h(15,4)="M014": h(15,5)="木村四男"
    h(16,1)="1": h(16,2)="2024/03/15": h(16,3)="4": h(16,4)="M019": h(16,5)="橋本九男"
    h(17,1)="1": h(17,2)="2024/03/15": h(17,3)="5": h(17,4)="M005": h(17,5)="伊藤五郎"
    h(18,1)="1": h(18,2)="2024/03/15": h(18,3)="5": h(18,4)="M010": h(18,5)="吉田十郎"
    h(19,1)="1": h(19,2)="2024/03/15": h(19,3)="5": h(19,4)="M015": h(19,5)="林五男"
    h(20,1)="1": h(20,2)="2024/03/15": h(20,3)="5": h(20,4)="M020": h(20,5)="阿部十男"
    h(21,1)="1": h(21,2)="2024/03/15": h(21,3)="6": h(21,4)="M021": h(21,5)="石川一郎"
    h(22,1)="1": h(22,2)="2024/03/15": h(22,3)="6": h(22,4)="M022": h(22,5)="前田二郎"
    h(23,1)="1": h(23,2)="2024/03/15": h(23,3)="6": h(23,4)="M023": h(23,5)="後藤三郎"
    h(24,1)="1": h(24,2)="2024/03/15": h(24,3)="6": h(24,4)="M024": h(24,5)="村上四郎"

    Dim i As Integer
    For i = 1 To 24
        ws.Cells(i + 1, 1).Value = CInt(h(i, 1))
        ws.Cells(i + 1, 2).Value = h(i, 2)
        ws.Cells(i + 1, 3).Value = CInt(h(i, 3))
        ws.Cells(i + 1, 4).Value = h(i, 4)
        ws.Cells(i + 1, 5).Value = h(i, 5)
        If (i + 1) Mod 2 = 0 Then ws.Rows(i + 1).Interior.Color = RGB(242, 242, 242)
    Next i
    ws.Columns.AutoFit
End Sub

