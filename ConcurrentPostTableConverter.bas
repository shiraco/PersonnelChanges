Attribute VB_Name = "ConcurrentPostTableConverter"

Option Explicit

'
' 横に並んだ兼務の人の所属情報を縦に並べるように変換する
' 兼務数は3まで対応
'
' author koji shiraishi
' since 2014/03/31
'

Sub ConvertConcurrentPostTable()

    
    '=======================================================================================================
    ' 初期設定
    '=======================================================================================================
    
    ' Excelテーブルの範囲の定義
    Dim source_table As Range
    Dim target_table As Range

    ' それぞれの列の位置（インデックス）を定数（Const）として定義
    ' COL_S_* は source_table （変換前テーブル）での列位置
    Const COL_S_COMMON_PREFIX_START As Integer = 1   ' 事由名称１
    Const COL_S_COMMON_PREFIX_END As Integer = 6   ' 新所属略称
    
    Const COL_S_AFT_PREFIX_START As Integer = 7   ' 新所属
    Const COL_S_AFT_PREFIX_END As Integer = 12    ' 新事業所

    Const COL_S_AFT_REPEAT_1A As Integer = 13     ' 新兼務所属１
    Const COL_S_AFT_REPEAT_1B As Integer = 14     ' 新兼務所属長１
    Const COL_S_AFT_REPEAT_2A As Integer = 15     ' 新兼務所属２
    Const COL_S_AFT_REPEAT_2B As Integer = 16     ' 新兼務所属長２
    Const COL_S_AFT_REPEAT_3A As Integer = 17     ' 新兼務所属３
    Const COL_S_AFT_REPEAT_3B As Integer = 18     ' 新兼務所属長３

    Const COL_S_AFT_SUFFIX_START As Integer = 19  ' 新他社出向先
    Const COL_S_AFT_SUFFIX_END As Integer = 21    ' 新出向割合

    Const COL_S_BEF_PREFIX_START As Integer = 22  ' 旧所属
    Const COL_S_BEF_PREFIX_END As Integer = 24  ' 旧所属略称
    
    Const COL_S_BEF_REPEAT_1A As Integer = 25     ' 旧兼務所属１
    Const COL_S_BEF_REPEAT_1B As Integer = 26     ' 旧兼務所属長１
    Const COL_S_BEF_REPEAT_2A As Integer = 27     ' 旧兼務所属２
    Const COL_S_BEF_REPEAT_2B As Integer = 28     ' 旧兼務所属長２
    Const COL_S_BEF_REPEAT_3A As Integer = 29     ' 旧兼務所属３
    Const COL_S_BEF_REPEAT_3B As Integer = 30     ' 旧兼務所属長３
    
    Const COL_S_BEF_SUFFIX_START As Integer = 31  ' 旧他社出向先
    Const COL_S_BEF_SUFFIX_END As Integer = 31    ' 旧出向割合
    
    'COL_T_* は target_table （変換後テーブル）での列位置
    Const COL_T_AFT_UNIFY_A As Integer = 6        ' 新所属
    Const COL_T_AFT_UNIFY_B As Integer = 7        ' 新所属組織長
    
    Const COL_T_BEF_UNIFY_A As Integer = 10       ' 旧所属
    Const COL_T_BEF_UNIFY_B As Integer = 11       ' 旧所属組織長

    ' Dim COL_S_SKIPS As Variant                ' スキップ対象の列（定数ではないけど、変更しないので大文字で宣言）
    ' COL_S_SKIPS = Array(6, 7, 9, 10, 20, 21)   ' 表示順、新本務、新グレード 、新職種、新他社略称、新出向割合


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
            If c = 6 Or c = 7 Or c = 9 Or c = 10 Or c = 20 Or c = 21 Then
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

                '----------------------------------------------------
                ' before's field
                '----------------------------------------------------
                
                ' prefix field
                ElseIf COL_S_BEF_PREFIX_START <= c And c <= COL_S_BEF_PREFIX_END Then
                    target_table(target_r, target_c) = source_table(r, c)
                
                ' 旧兼務1所属
                ElseIf c = COL_S_BEF_REPEAT_1A Then
                    target_r = target_r + 1
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 1

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "（兼務１）")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))

                    End If

                ' 旧兼務1所属長
                ElseIf c = COL_S_BEF_REPEAT_1B Then
                    If ConcurrentPosts >= 1 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 1
                    target_c = target_c + 1

                ' 旧兼務2所属
                ElseIf c = COL_S_BEF_REPEAT_2A Then
                    target_r = target_r + 2
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 2

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "（兼務２）")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' 旧兼務2所属長
                ElseIf c = COL_S_BEF_REPEAT_2B Then
                    If ConcurrentPosts >= 2 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 2
                    target_c = target_c + 1

                ' 旧兼務3所属
                ElseIf c = COL_S_BEF_REPEAT_3A Then
                    target_r = target_r + 3
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 3
                         
                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "（兼務３）")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' 旧兼務3所属長
                ElseIf c = COL_S_BEF_REPEAT_3B Then
                    If ConcurrentPosts >= 3 Then
                       target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 3
                    target_c = target_c + 1

                ' suffix field
                ElseIf COL_S_BEF_SUFFIX_START <= c And c <= COL_S_BEF_SUFFIX_END Then
                    target_table(target_r, target_c) = source_table(r, c)

                End If

                '----------------------------------------------------
                ' common process
                '----------------------------------------------------
                
                target_c = target_c + 1 ' 列移動
            End If

            ' 新の右端(after's suffix end)での処理
            If c = COL_S_AFT_SUFFIX_END Then
                BefAftMaxConcurrentPosts = ConcurrentPosts
                ConcurrentPosts = 0
            End If
            
            ' 旧の右端(before's suffix end)での処理
            If c = COL_S_BEF_SUFFIX_END Then
                If BefAftMaxConcurrentPosts < ConcurrentPosts Then
                    BefAftMaxConcurrentPosts = ConcurrentPosts
                End If
                ConcurrentPosts = 0
            End If

        Next
        
        ' 改行処理
        target_c = 1                                   ' 列移動
        target_r = target_r + 1                        ' 行移動（通常分）
        target_r = target_r + BefAftMaxConcurrentPosts ' 行移動（兼務数分の加算）
        BefAftMaxConcurrentPosts = 0
        
    Next

    
    '=======================================================================================================
    ' 書式の設定
    '=======================================================================================================
    
    ' 【target_tableの書式の初期化】条件付書式をクリア＆設定
    Set target_table = Range("target_table") ' target_table が拡張されているので、改めて定義する
    With target_table.ListObject.Range
        .FormatConditions.Delete      ' 既に条件付書式が定義されていたら、条件付書式をクリアする
        .HorizontalAlignment = xlGeneral ' 原則左詰めに書式設定
    End With

    ' 【兼務行の全体（全列）の書式設定】その行の社員列が（空白であれば）その行の上側の罫線を無くす
    With target_table.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($D4)")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' 【兼務行の所属列の書式設定】
    ' target_table の最終位置のインデックスを取得 (last_target_r, last_tareget_c)
    Dim last_target_r As Integer
    last_target_r = target_r - 1

    ' 所属列に関しては、その行の社員列が（空白であれば）その行の上側の罫線を点線にする
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_A), "$D4", "$F4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_B), "$D4", "$F4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_BEF_UNIFY_A), "$D4", "$J4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_BEF_UNIFY_B), "$D4", "$J4")

    Application.ScreenUpdating = True ' 描画ON

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
Sub SetConcurrentPostFormatConditions(Columns As Range, referenceCellStr As String, selfCellStr As String)
    
    With Columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCellStr & "), NOT(ISBLANK(" & selfCellStr & ")))").Borders(xlTop).LineStyle = xlDot
    End With

End Sub
