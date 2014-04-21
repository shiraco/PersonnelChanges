Attribute VB_Name = "ConvertConcurrentPostTable"

Option Explicit

'
' 横に並んだ兼務の人の所属情報を縦に並べるように変換する
' 兼務数は3、出向数は2まで対応
'
' Args:
'  sourceReadingType  "NEW_ONLY": 新所属のみ転記 , "NEW_AND_OLD": 新旧所属両方転記
'
' author koji shiraishi
' since 2014/04/21
' version 1.1
'
Sub ConvertConcurrentPostTable(sourceReadingType As String)

    '=======================================================================================================
    ' 開始処理
    '=======================================================================================================

    ' 【debug】処理時間の計測start
    Dim startTime, stopTime As Variant
    startTime = Time 'ここから実行時間のカウントを開始します

    ' 【debug】描画OFF
    Application.ScreenUpdating = False

    '=======================================================================================================
    ' 初期設定
    '=======================================================================================================

    ' Excelテーブルの範囲の定義
    ' Excel にあらかじめ "source_table", "target_table" という名前でテーブルを定義しておく
    Dim sourceTable As Range
    Dim targetTable As Range
    Set sourceTable = Range("source_table")
    Set targetTable = Range("target_table")

    ' パラメータ
    Const SPACE_INDENT_STR As String = "　　"             ' インデント幅を定義

    ' それぞれの列の位置（インデックス）を定数（Const）として定義
    ' COL_S_* は sourceTable （変換前テーブル）での列位置
    Const COL_S_COMMON_PREFIX_START As Integer = 1        ' 事由名称
    Const COL_S_COMMON_PREFIX_END As Integer = 4          ' 氏名

    Const COL_S_NEW_REPEAT1_0A As Integer = 5             ' 新本務所属
    Const COL_S_NEW_REPEAT1_0B As Integer = 9             ' 新本務役職
    Const COL_S_NEW_REPEAT1_1A As Integer = 10            ' 新兼務所属１
    Const COL_S_NEW_REPEAT1_1B As Integer = 14            ' 新兼務役職１
    Const COL_S_NEW_REPEAT1_2A As Integer = 15            ' 新兼務所属２
    Const COL_S_NEW_REPEAT1_2B As Integer = 19            ' 新兼務役職２
    Const COL_S_NEW_REPEAT1_3A As Integer = 20            ' 新兼務所属３
    Const COL_S_NEW_REPEAT1_3B As Integer = 24            ' 新兼務役職３

    Const COL_S_NEW_REPEAT2_1A As Integer = 28            ' 新出向先１
    Const COL_S_NEW_REPEAT2_2A As Integer = 32            ' 新出向先２

    Const COL_S_NEW_SUFFIX_START As Integer = 33          ' 新事業所
    Const COL_S_NEW_SUFFIX_END As Integer = 35            ' 新職種*削除
    Const COL_S_NEW_END As Integer = COL_S_NEW_SUFFIX_END ' 新職種*削除

    Const COL_S_OLD_REPEAT1_0A As Integer = 36            ' 旧本務所属
    Const COL_S_OLD_REPEAT1_0B As Integer = 40            ' 旧本務役職
    Const COL_S_OLD_REPEAT1_1A As Integer = 41            ' 旧兼務所属１
    Const COL_S_OLD_REPEAT1_1B As Integer = 45            ' 旧兼務役職１
    Const COL_S_OLD_REPEAT1_2A As Integer = 46            ' 旧兼務所属２
    Const COL_S_OLD_REPEAT1_2B As Integer = 50            ' 旧兼務役職２
    Const COL_S_OLD_REPEAT1_3A As Integer = 51            ' 旧兼務所属３
    Const COL_S_OLD_REPEAT1_3B As Integer = 55            ' 旧兼務役職３

    Const COL_S_OLD_REPEAT2_1A As Integer = 59            ' 旧出向先１
    Const COL_S_OLD_REPEAT2_2A As Integer = 63            ' 旧出向先２

    Const COL_S_OLD_SUFFIX_START As Integer = 64          ' 旧事業所
    Const COL_S_OLD_SUFFIX_END As Integer = 66            ' 旧職種*削除
    Const COL_S_OLD_END As Integer = COL_S_OLD_SUFFIX_END ' 旧職種*削除

    'COL_T_* は targetTable （変換後テーブル）での列位置
    Const COL_T_NEW_REASON1 As Integer = 1                ' 事由名称１
    Const COL_T_NEW_REASON2 As Integer = 2                ' 事由名称２
    Const COL_T_NEW_CHANGE_DATE As Integer = 3            ' 発令日
    Const COL_T_NEW_CONCURRENT_POST_LABEL As Integer = 5  ' 氏名
    Const COL_T_NEW_UNIFY1_A As Integer = 6               ' 新所属
    Const COL_T_NEW_UNIFY1_B As Integer = 7               ' 新役職
    Const COL_T_NEW_UNIFY2_A As Integer = 8               ' 新出向先
    Const COL_T_OLD_UNIFY1_A As Integer = 10              ' 旧所属
    Const COL_T_OLD_UNIFY1_B As Integer = 11              ' 旧役職
    Const COL_T_OLD_UNIFY2_A As Integer = 12              ' 旧出向先

    Const CELL_T_NEW_REASON1_STR As String = "A7"         ' 事由名称１
    Const CELL_T_NEW_REASON2_STR As String = "B7"         ' 事由名称２
    Const CELL_T_NEW_CHANGE_DATE_STR As String = "C7"     ' 発令日
    Const CELL_T_EMPLOYEE_NO_STR As String = "$D7"        ' 社員番号列の最上行
    Const CELL_T_NEW_POST_NAME_STR As String = "F7"       ' 新所属の最上行
    Const CELL_T_NEW_ACOMPANY_NAME_STR As String = "H7"   ' 新出向先の最上行
    Const CELL_T_OLD_POST_NAME_STR As String = "J7"       ' 旧所属の最上行
    Const CELL_T_OLD_ACOMPANY_NAME_STR As String = "L7"   ' 旧出向先の最上行

    '=======================================================================================================
    ' source_table -> target_table への転記処理
    '=======================================================================================================

    ' 行パラメータ
    ' その人の兼務数（本務除く）
    Dim newConcurrentPosts, oldConcurrentPosts As Integer
    newConcurrentPosts = 0
    oldConcurrentPosts = 0
    ' その人の出向会社数
    Dim newAssigneeCompanies, oldAssigneeCompanies As Integer
    newAssigneeCompanies = 0
    oldAssigneeCompanies = 0
    ' その人が使用する行数（通常1、兼務数、出向数によって増加）
    Dim usingRows, newUsingRows, oldUsingRows As Integer
    usingRows = 1
    newUsingRows = 1
    oldUsingRows = 1
    ' その人の本務において使用する行数（出向数に応じて増加）
    Dim newUsingRowsInMainPost, oldUsingRowsInMainPost As Integer
    newUsingRowsInMainPost = 1
    oldUsingRowsInMainPost = 1

    ' sourceTable上の読込インデックス位置 (r, c)
    Dim r, c As Long

    ' targetTable上の書込インデックス位置 (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    ' for ループの index
    Dim i As Long

    Dim sourceTableReadingCols As Long
    sourceTableReadingCols = IIf(sourceReadingType = "NEW_ONLY", COL_S_NEW_END, sourceTable.columns.Count)

    For r = 1 To sourceTable.Rows.Count
        For c = 1 To sourceTableReadingCols

            ' スキップ対象の列であれば何もしない
            ' FIXME もう少しまともなスキップのさせ方でスキップさせる
            ' ※ source_table の列が変更になった場合、要修正
            ' 新のスキップ対象
            If COL_S_NEW_REPEAT1_0A < c And c < COL_S_NEW_REPEAT1_0B Or _
               COL_S_NEW_REPEAT1_1A < c And c < COL_S_NEW_REPEAT1_1B Or _
               COL_S_NEW_REPEAT1_2A < c And c < COL_S_NEW_REPEAT1_2B Or _
               COL_S_NEW_REPEAT1_3A < c And c < COL_S_NEW_REPEAT1_3B Or _
               c = 25 Or c = 26 Or c = 27 Or c = 29 Or c = 30 Or c = 31 Or c = 34 Or c = 35 Then

                'NOP

            ' 旧のスキップ対象
            ElseIf sourceReadingType <> "NEW_ONLY" And _
               (COL_S_OLD_REPEAT1_0A < c And c < COL_S_OLD_REPEAT1_0B Or _
                COL_S_OLD_REPEAT1_1A < c And c < COL_S_OLD_REPEAT1_1B Or _
                COL_S_OLD_REPEAT1_2A < c And c < COL_S_OLD_REPEAT1_2B Or _
                COL_S_OLD_REPEAT1_3A < c And c < COL_S_OLD_REPEAT1_3B Or _
                c = 56 Or c = 57 Or c = 58 Or c = 60 Or c = 61 Or c = 62 Or c = 65 Or c = 66) Then

                'NOP

            ' スキップ対象外列
            Else
                '----------------------------------------------------
                ' common's field
                '----------------------------------------------------

                If COL_S_COMMON_PREFIX_START <= c And c <= COL_S_COMMON_PREFIX_END Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                    ' 「事由名称２」列のスペースを１列分確保
                    If c = COL_T_NEW_REASON1 Then
                        target_c = target_c + 1 ' 列移動
                    End If

                '----------------------------------------------------
                ' new's field
                '----------------------------------------------------

                ' 新本務所属
                ElseIf c = COL_S_NEW_REPEAT1_0A Then
                    ' 出向数事前判定
                    If sourceTable(r, COL_S_NEW_REPEAT2_1A) <> "" Then
                        newAssigneeCompanies = IIf(sourceTable(r, COL_S_NEW_REPEAT2_2A) <> "", 2, 1)
                        newUsingRowsInMainPost = IIf(newAssigneeCompanies >= 2, newAssigneeCompanies, 1)

                        ' target_table の範囲を自動拡張させるためにセルにダミー値（SPACE_INDENT_STR）を設定
                        If newUsingRowsInMainPost >= 2 Then
                            For i = 1 To newUsingRowsInMainPost - 1
                                targetTable(target_r + i, COL_T_NEW_UNIFY2_A) = SPACE_INDENT_STR
                            Next i
                        End If
                    End If

                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' 新本務役職
                ElseIf c = COL_S_NEW_REPEAT1_0B Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' 新兼務1所属
                ElseIf c = COL_S_NEW_REPEAT1_1A Then
                    target_c = target_c - 2 ' 所属列に戻る

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 1

                        target_r = target_r + (newConcurrentPosts + (newUsingRowsInMainPost - 1)) ' 兼務行へ改行する

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "（兼務１）")
                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' 新兼務1役職
                ElseIf c = COL_S_NEW_REPEAT1_1B Then
                    If newConcurrentPosts >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (newConcurrentPosts + (newUsingRowsInMainPost - 1)) ' 兼務行へ改行した分戻る
                    End If

                ' 新兼務2所属
                ElseIf c = COL_S_NEW_REPEAT1_2A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 2

                        target_r = target_r + (newConcurrentPosts + (newUsingRowsInMainPost - 1))

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "（兼務２）")
                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' 新兼務2役職
                ElseIf c = COL_S_NEW_REPEAT1_2B Then
                    If newConcurrentPosts >= 2 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (newConcurrentPosts + (newUsingRowsInMainPost - 1))
                    End If

                ' 新兼務3所属
                ElseIf c = COL_S_NEW_REPEAT1_3A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 3

                        target_r = target_r + (newConcurrentPosts + (newUsingRowsInMainPost - 1))

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "（兼務３）")
                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' 新兼務3役職
                ElseIf c = COL_S_NEW_REPEAT1_3B Then
                    If newConcurrentPosts >= 3 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (newConcurrentPosts + (newUsingRowsInMainPost - 1))
                    End If

                ' 新出向先1
                ElseIf c = COL_S_NEW_REPEAT2_1A Then
                    If newAssigneeCompanies >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)
                    End If

                ' 新出向先2
                ElseIf c = COL_S_NEW_REPEAT2_2A Then
                    target_c = target_c - 1

                    If newAssigneeCompanies >= 2 Then
                        target_r = target_r + (newUsingRowsInMainPost - 1)

                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (newUsingRowsInMainPost - 1)
                   End If

                ' suffix field
                ElseIf COL_S_NEW_SUFFIX_START <= c And c <= COL_S_NEW_SUFFIX_END Then
                    newUsingRows = newUsingRowsInMainPost + newConcurrentPosts
                    targetTable(target_r, target_c) = sourceTable(r, c)

                '----------------------------------------------------
                ' old's field  (sourceReadingType <> "NEW_ONLY")
                '----------------------------------------------------

                ' 旧本務所属
                ElseIf c = COL_S_OLD_REPEAT1_0A Then
                    ' 出向数事前判定
                    If sourceTable(r, COL_S_OLD_REPEAT2_1A) <> "" Then
                        oldAssigneeCompanies = IIf(sourceTable(r, COL_S_OLD_REPEAT2_2A) <> "", 2, 1)
                        oldUsingRowsInMainPost = IIf(oldAssigneeCompanies >= 2, oldAssigneeCompanies, 1)

                        ' target_table の範囲を自動拡張させるためにセルにダミー値（SPACE_INDENT_STR）を設定
                        If oldUsingRowsInMainPost >= 2 Then
                            For i = 1 To oldUsingRowsInMainPost - 1
                                targetTable(target_r + i, COL_T_OLD_UNIFY2_A) = SPACE_INDENT_STR
                            Next i
                        End If
                    End If

                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' 旧本務役職
                ElseIf c = COL_S_OLD_REPEAT1_0B Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' 旧兼務1所属
                ElseIf c = COL_S_OLD_REPEAT1_1A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        oldConcurrentPosts = 1

                        target_r = target_r + (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))

                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' 旧兼務1役職
                ElseIf c = COL_S_OLD_REPEAT1_1B Then
                    If oldConcurrentPosts >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))
                    End If

                ' 旧兼務2所属
                ElseIf c = COL_S_OLD_REPEAT1_2A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        oldConcurrentPosts = 2

                        target_r = target_r + (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))

                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' 旧兼務2役職
                ElseIf c = COL_S_OLD_REPEAT1_2B Then
                    If oldConcurrentPosts >= 2 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))
                    End If

                '旧兼務3所属
                ElseIf c = COL_S_OLD_REPEAT1_3A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        oldConcurrentPosts = 3

                        target_r = target_r + (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))

                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' 旧兼務3役職
                ElseIf c = COL_S_OLD_REPEAT1_3B Then
                    If oldConcurrentPosts >= 3 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))
                    End If

                ' 旧出向先1
                ElseIf c = COL_S_OLD_REPEAT2_1A Then
                    If oldAssigneeCompanies >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)
                    End If

                ' 旧出向先2
                ElseIf c = COL_S_OLD_REPEAT2_2A Then
                    target_c = target_c - 1

                    If oldAssigneeCompanies >= 2 Then
                        target_r = target_r + (oldUsingRowsInMainPost - 1)

                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (oldUsingRowsInMainPost - 1)
                   End If

                ' suffix field
                ElseIf COL_S_OLD_SUFFIX_START <= c And c <= COL_S_OLD_SUFFIX_END Then
                    oldUsingRows = oldUsingRowsInMainPost + oldConcurrentPosts
                    targetTable(target_r, target_c) = sourceTable(r, c)

                End If
                '----------------------------------------------------
                ' common process
                '----------------------------------------------------

                target_c = target_c + 1 ' 列移動
            End If

        Next

        ' 改行処理
        usingRows = IIf(newUsingRows >= oldUsingRows, newUsingRows, oldUsingRows)
        target_c = 1                    ' 列移動
        target_r = target_r + usingRows ' 行移動

        ' 行パラメータのリセット
        newConcurrentPosts = 0
        newAssigneeCompanies = 0
        newUsingRows = 1
        newUsingRowsInMainPost = 1
        oldConcurrentPosts = 0
        oldAssigneeCompanies = 0
        oldUsingRows = 1
        oldUsingRowsInMainPost = 1
        usingRows = 1

    Next

    '=======================================================================================================
    ' target_table の書式の設定
    '=======================================================================================================

    ' 【書式の初期化】条件付書式をクリア
    Set targetTable = Range("target_table") ' targetTable が処理前より拡張されているので、改めて範囲を再定義する
    With targetTable.ListObject.Range
        .FormatConditions.Delete      ' 既に条件付書式が定義されていたら、条件付書式をクリアする（条件付じゃない書式はクリアしない）
    End With

    ' 【全列の書式設定（その人物における2行目以降が対象となる）】その行の"社員番号"列が（空白であれば）兼務行とみなしその行の上側の罫線を無くす
    With targetTable.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(" & CELL_T_EMPLOYEE_NO_STR & ")")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' 【事由名称、発令日列の書式設定】自身のセルが空白の場合、上側の罫線を無くす
    Call SetFormatConditionsTopLineNone(targetTable.columns(COL_T_NEW_REASON1), CELL_T_NEW_REASON1_STR)
    Call SetFormatConditionsTopLineNone(targetTable.columns(COL_T_NEW_REASON2), CELL_T_NEW_REASON2_STR)
    Call SetFormatConditionsTopLineNone(targetTable.columns(COL_T_NEW_CHANGE_DATE), CELL_T_NEW_CHANGE_DATE_STR)

    ' 【所属列の書式設定（その人物における2行目以降が対象となる）】
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_NEW_POST_NAME_STR)
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_B), CELL_T_EMPLOYEE_NO_STR, CELL_T_NEW_POST_NAME_STR)
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_OLD_UNIFY1_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_OLD_POST_NAME_STR)
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_OLD_UNIFY1_B), CELL_T_EMPLOYEE_NO_STR, CELL_T_OLD_POST_NAME_STR)

    ' 【出向先列の書式設定（その人物における2行目以降が対象となる）】
    Call SetAssigneeCompanyFormatConditions(targetTable.columns(COL_T_NEW_UNIFY2_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_NEW_POST_NAME_STR, CELL_T_NEW_ACOMPANY_NAME_STR)
    Call SetAssigneeCompanyFormatConditions(targetTable.columns(COL_T_OLD_UNIFY2_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_OLD_POST_NAME_STR, CELL_T_OLD_ACOMPANY_NAME_STR)

    '=======================================================================================================
    ' 終了処理
    '=======================================================================================================

    ' 【debug】描画ON
    Application.ScreenUpdating = True

    ' 【debug】処理時間の計測end
    stopTime = Time
    stopTime = stopTime - startTime
    MsgBox "所要時間は" & Minute(stopTime) & "分" & Second(stopTime) & "秒 でした"

End Sub

'
' 指定したCell (氏名の下の欄) に引数の文字列（兼務ラベル）を右寄せした上でセットするサブルーチン
'
Sub SetConcurrentPostLabel(target As Range, label As String)

    With target
        .Value = label
        .HorizontalAlignment = xlRight
    End With

End Sub

'
' 指定したColumns (事由名称、発令日) に条件付書式をセットするサブルーチン
'
Sub SetFormatConditionsTopLineNone(columns As Range, referenceCell1Str As String)

    With columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(NOT(ISBLANK(" & referenceCell1Str & ")))").Borders(xlTop).LineStyle = xlContinuous
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "))").Borders(xlTop).LineStyle = xlNone
    End With

End Sub

'
' 指定したColumns (所属、役職列) に条件付書式をセットするサブルーチン
'
Sub SetConcurrentPostFormatConditions(columns As Range, referenceCell1Str As String, referenceCell2Str As String)

    With columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), NOT(ISBLANK(" & referenceCell2Str & ")))").Borders(xlTop).LineStyle = xlDash
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), ISBLANK(" & referenceCell2Str & "))").Borders(xlTop).LineStyle = xlNone
    End With

End Sub

'
' 指定したColumns (出向先列) に条件付書式をセットするサブルーチン
'
Sub SetAssigneeCompanyFormatConditions(columns As Range, referenceCell1Str As String, referenceCell2Str As String, selfCellStr As String)

    With columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), NOT(ISBLANK(" & referenceCell2Str & ")))").Borders(xlTop).LineStyle = xlDash
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), ISBLANK(" & referenceCell2Str & "), NOT(ISBLANK(" & selfCellStr & ")))").Borders(xlTop).LineStyle = xlDot
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), ISBLANK(" & referenceCell2Str & "), ISBLANK(" & selfCellStr & "))").Borders(xlTop).LineStyle = xlNone
    End With

End Sub

