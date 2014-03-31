Attribute VB_Name = "ConcurrentPostTableConverter"

Option Explicit

'
' 横に並んだ兼務の人の所属情報を縦に並べるように変換する
' 兼務数は3まで対応
'
' author koji shiraishi
' since 2014/03/31
' version 1.0
'
Sub ConvertConcurrentPostTable()

    '=======================================================================================================
    ' 初期設定
    '=======================================================================================================

    ' 【debug】処理時間の計測start
    Dim StartTime, StopTime As Variant
    StartTime = Time 'ここから実行時間のカウントを開始します

    ' Excelテーブルの範囲の定義
    Dim source_table As Range
    Dim target_table As Range

    ' それぞれの列の位置（インデックス）を定数（Const）として定義
    ' COL_S_* は source_table （変換前テーブル）での列位置
    Const COL_S_COMMON_PREFIX_START As Integer = 1 ' 社員番号
    Const COL_S_COMMON_PREFIX_END As Integer = 2   ' 新所属略称
    
    Const COL_S_AFT_PREFIX_START As Integer = 5   ' 新所属
    Const COL_S_AFT_PREFIX_END As Integer = 9     ' 新事業所

    Const COL_S_AFT_REPEAT_1A As Integer = 10     ' 新兼務所属１
    Const COL_S_AFT_REPEAT_1B As Integer = 11     ' 新兼務所属長１
    Const COL_S_AFT_REPEAT_2A As Integer = 12     ' 新兼務所属２
    Const COL_S_AFT_REPEAT_2B As Integer = 13     ' 新兼務所属長２
    Const COL_S_AFT_REPEAT_3A As Integer = 14     ' 新兼務所属３
    Const COL_S_AFT_REPEAT_3B As Integer = 15     ' 新兼務所属長３

    Const COL_S_AFT_SUFFIX_START As Integer = 16  ' 新他社出向先
    Const COL_S_AFT_SUFFIX_END As Integer = 18    ' 新出向割合

    'COL_T_* は target_table （変換後テーブル）での列位置
    Const COL_T_AFT_UNIFY_A As Integer = 3        ' 新所属
    Const COL_T_AFT_UNIFY_B As Integer = 4        ' 新所属組織長

    ' Dim COL_S_SKIPS As Variant                ' スキップ対象の列（定数ではないけど、変更しないので大文字で宣言）
    ' COL_S_SKIPS = Array(3, 4, 6, 7, 17, 18)   ' 表示順、新本務、新グレード 、新職種、新他社略称、新出向割合


    ' Excel にあらかじめ "source_table", "target_table" という名前でテーブルを定義しておく
    Set source_table = Range("source_table")
    Set target_table = Range("target_table")

    ' source_table上の読込インデックス位置 (r, c)
    Dim r, c As Long

    ' target_table上の書込インデックス位置 (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    ' その人の新兼務数（本務除く）
    Dim ConcurrentPosts, BefAftMaxConcurrentPosts  As Integer
    ConcurrentPosts = 0
    BefAftMaxConcurrentPosts = 0
    
    '=======================================================================================================
    ' main 処理
    '=======================================================================================================
    
    Application.ScreenUpdating = False ' 描画OFF

    For r = 1 To source_table.Rows.Count
        For c = 1 To source_table.Columns.Count

            ' スキップ対象の列であれば何もしない
            ' 配列（COL_S_SKIPS）との比較の仕方がわからないので、べた書き
            If c = 3 Or c = 4 Or c = 6 Or c = 7 Or c = 17 Or c = 18 Then
                'NOP

            ' スキップ対象外
            Else
                '----------------------------------------------------
                ' common's field
                '----------------------------------------------------
                
                If COL_S_COMMON_PREFIX_START <= c And c <= COL_S_COMMON_PREFIX_END Then
                    target_table(target_r, target_c) = source_table(r, c)


                '----------------------------------------------------
                ' after's field
                '----------------------------------------------------
                
                ' prefix field
                ElseIf COL_S_AFT_PREFIX_START <= c And c <= COL_S_AFT_PREFIX_END Then
                    target_table(target_r, target_c) = source_table(r, c)
                
                ' 新兼務1所属
                ElseIf c = COL_S_AFT_REPEAT_1A Then
                    target_r = target_r + 1
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 1

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "（兼務１）")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))

                    End If

                ' 新兼務1所属長
                ElseIf c = COL_S_AFT_REPEAT_1B Then
                    If ConcurrentPosts >= 1 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 1
                    target_c = target_c + 1

                ' 新兼務2所属
                ElseIf c = COL_S_AFT_REPEAT_2A Then
                    target_r = target_r + 2
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 2

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "（兼務２）")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' 新兼務2所属長
                ElseIf c = COL_S_AFT_REPEAT_2B Then
                    If ConcurrentPosts >= 2 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 2
                    target_c = target_c + 1

                ' 新兼務3所属
                ElseIf c = COL_S_AFT_REPEAT_3A Then
                    target_r = target_r + 3
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 3

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "（兼務３）")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' 新兼務3所属長
                ElseIf c = COL_S_AFT_REPEAT_3B Then
                    If ConcurrentPosts >= 3 Then
                       target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 3
                    target_c = target_c + 1

                ' suffix field
                ElseIf COL_S_AFT_SUFFIX_START <= c And c <= COL_S_AFT_SUFFIX_END Then
                    target_table(target_r, target_c) = source_table(r, c)

                End If

                '----------------------------------------------------
                ' common process
                '----------------------------------------------------

                target_c = target_c + 1 ' 列移動
            End If

        Next

        ' 改行処理
        target_c = 1                          ' 列移動
        target_r = target_r + 1               ' 行移動（通常分）
        target_r = target_r + concurrentPosts ' 行移動（兼務数分の加算）
        concurrentPosts = 0
        
    Next

    '=======================================================================================================
    ' 書式の設定
    '=======================================================================================================

    ' 【target_tableの書式の初期化】条件付書式をクリア＆設定
    Set target_table = Range("target_table") ' target_table が拡張されているので、改めて定義する
    With target_table.ListObject.Range
        .FormatConditions.Delete      ' 既に条件付書式が定義されていたら、条件付書式をクリアする
    End With

    ' 【兼務行の全体（全列）の書式設定】その行の社員列が（空白であれば）その行の上側の罫線を無くす
    With target_table.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($A4)")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' 【兼務行の所属列の書式設定】
    ' 所属列に関しては、その行の社員列が（空白であれば）その行の上側の罫線を点線にする
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_A), "$A4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_B), "$A4")

    Application.ScreenUpdating = True ' 描画ON

    ' 【debug】処理時間の計測end
    StopTime = Time
    StopTime = StopTime - StartTime
    MsgBox "所要時間は" & Minute(StopTime) & "分" & Second(StopTime) & "秒 でした"

End Sub

'
' 指定したCell (Range) に引数の文字列（ラベル）を右寄せした上でセットするサブルーチン
'
Sub SetConcurrentPostLabel(target As Range, label As String)

    With target
        .Value = label
        .HorizontalAlignment = xlRight
    End With

End Sub

'
' 指定したCell (Range) に所属をインデント付でセットするサブルーチン
'
Sub SetConcurrentPost(target As Range, PostName As String)

     target = "　　" & PostName ' 全角スペース×２でインデント

End Sub

'
' 指定したColumns (Range) にの条件付書式をセットするサブルーチン
'
Sub SetConcurrentPostFormatConditions(Columns As Range, referenceCellStr As String)
    
    With Columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(" & referenceCellStr & ")").Borders(xlTop).LineStyle = xlDot
    End With

End Sub
