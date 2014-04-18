Attribute VB_Name = "ConverterForAllEmployees"

Option Explicit

'
' 横に並んだ兼務の人の所属情報を縦に並べるように変換する
' 兼務数は3まで対応
'
' author koji shiraishi
' since 2014/04/18
' version 1.1
'
Sub ConvertConcurrentPostTableForAllEmployees()

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
    Const STR_SPACE_INDENT As String = "　　"             ' インデント幅を定義
    Const STR_FIRST_BASE_CELL_ROW As String = "$A4"       ' 書式設定の基準とする最左上のセルからの列参照

    ' それぞれの列の位置（インデックス）を定数（Const）として定義
    ' COL_S_* は sourceTable （変換前テーブル）での列位置
    Const COL_S_COMMON_PREFIX_START As Integer = 1        ' 社員番号
    Const COL_S_COMMON_PREFIX_END As Integer = 4          ' 新所属略称

    Const COL_S_NEW_PREFIX_START As Integer = 5           ' 新所属
    Const COL_S_NEW_PREFIX_END As Integer = 9             ' 新事業所

    Const COL_S_NEW_REPEAT_1A As Integer = 10             ' 新兼務所属１
    Const COL_S_NEW_REPEAT_1B As Integer = 11             ' 新兼務所属長１
    Const COL_S_NEW_REPEAT_2A As Integer = 12             ' 新兼務所属２
    Const COL_S_NEW_REPEAT_2B As Integer = 13             ' 新兼務所属長２
    Const COL_S_NEW_REPEAT_3A As Integer = 14             ' 新兼務所属３
    Const COL_S_NEW_REPEAT_3B As Integer = 15             ' 新兼務所属長３

    Const COL_S_NEW_SUFFIX_START As Integer = 16          ' 新他社出向先
    Const COL_S_NEW_SUFFIX_END As Integer = 18            ' 新出向割合

    'COL_T_* は targetTable （変換後テーブル）での列位置
    Const COL_T_NEW_CONCURRENT_POST_LABEL As Integer = 2  ' 氏名

    Const COL_T_NEW_UNIFY1_A As Integer = 3               ' 新所属
    Const COL_T_NEW_UNIFY1_B As Integer = 4               ' 新所属組織長

    '=======================================================================================================
    ' main 処理
    '=======================================================================================================

    ' 行パラメータ
    ' その人の兼務数（本務除く）
    Dim newConcurrentPosts As Integer
    newConcurrentPosts = 0

    ' その人が使用する行数（通常1、兼務数によって増加）
    Dim usingRows, newUsingRows As Integer
    usingRows = 1
    newUsingRows = 1

    ' sourceTable上の読込インデックス位置 (r, c)
    Dim r, c As Long

    ' targetTable上の書込インデックス位置 (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    For r = 1 To sourceTable.Rows.Count
        For c = 1 To sourceTable.columns.Count

            ' スキップ対象の列であれば何もしない
            ' FIXME もう少しまともなスキップのさせ方でスキップさせる
            If c = 3 Or c = 4 Or c = 6 Or c = 7 Or c = 17 Or c = 18 Then

                'NOP

            ' スキップ対象外
            Else
                '----------------------------------------------------
                ' common's field
                '----------------------------------------------------

                If COL_S_COMMON_PREFIX_START <= c And c <= COL_S_COMMON_PREFIX_END Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                '----------------------------------------------------
                ' new's field
                '----------------------------------------------------

                ' prefix field
                ElseIf COL_S_NEW_PREFIX_START <= c And c <= COL_S_NEW_PREFIX_END Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' 新兼務1所属
                ElseIf c = COL_S_NEW_REPEAT_1A Then
                    target_c = target_c - 2
                    target_c = target_c - 1

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 1

                        target_r = target_r + newConcurrentPosts

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "（兼務１）")
                        targetTable(target_r, target_c) = STR_SPACE_INDENT & sourceTable(r, c)
                    End If

                ' 新兼務1所属長
                ElseIf c = COL_S_NEW_REPEAT_1B Then
                    If newConcurrentPosts >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - newConcurrentPosts
                    End If

                ' 新兼務2所属
                ElseIf c = COL_S_NEW_REPEAT_2A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 2

                        target_r = target_r + newConcurrentPosts

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "（兼務２）")
                        targetTable(target_r, target_c) = STR_SPACE_INDENT & sourceTable(r, c)
                    End If

                ' 新兼務2所属長
                ElseIf c = COL_S_NEW_REPEAT_2B Then
                    If newConcurrentPosts >= 2 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - newConcurrentPosts
                    End If

                ' 新兼務3所属
                ElseIf c = COL_S_NEW_REPEAT_3A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 3

                        target_r = target_r + newConcurrentPosts

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "（兼務３）")
                        targetTable(target_r, target_c) = STR_SPACE_INDENT & sourceTable(r, c)
                    End If

                ' 新兼務3所属長
                ElseIf c = COL_S_NEW_REPEAT_3B Then
                    If newConcurrentPosts >= 3 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - newConcurrentPosts
                    End If

                    target_c = target_c + 1
                
                ' suffix field
                ElseIf COL_S_NEW_SUFFIX_START <= c And c <= COL_S_NEW_SUFFIX_END Then
                    newUsingRows = newConcurrentPosts + 1
                    targetTable(target_r, target_c) = sourceTable(r, c)

                End If

                '----------------------------------------------------
                ' common process
                '----------------------------------------------------

                target_c = target_c + 1 ' 列移動
            End If

        Next

        ' 改行処理
        usingRows = newUsingRows
        target_c = 1                    ' 列移動
        target_r = target_r + usingRows ' 行移動

        ' 行パラメータのリセット
        newConcurrentPosts = 0
        newUsingRows = 1
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

    ' 【全列の書式設定（その人物における2行目以降）】その行の"社員番号"列が（空白であれば）兼務行とみなしその行の上側の罫線を無くす
    With targetTable.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(" & STR_FIRST_BASE_CELL_ROW & ")")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' 【所属列の書式設定（その人物における2行目以降）】
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_A), STR_FIRST_BASE_CELL_ROW, "C4")
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_B), STR_FIRST_BASE_CELL_ROW, "C4")

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
' 指定したColumns (所属、役職列) に条件付書式をセットするサブルーチン
'
Sub SetConcurrentPostFormatConditions(columns As Range, referenceCell1Str As String, referenceCell2Str As String)

    With columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), NOT(ISBLANK(" & referenceCell2Str & ")))").Borders(xlTop).LineStyle = xlDash
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), ISBLANK(" & referenceCell2Str & "))").Borders(xlTop).LineStyle = xlNone
    End With

End Sub

