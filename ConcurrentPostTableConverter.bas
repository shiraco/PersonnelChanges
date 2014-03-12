Attribute VB_Name = "ConcurrentPostTableConverter"


Option Explicit

'
' 横に並んだ兼務の人の所属情報を縦に並べるように変換する
'
' author koji shiraishi
' since 2014/03/12
'
Sub ConvertConcurrentPostTable()
    ' Excelテーブルの範囲の定義
    Dim source_table As Range
    Dim target_table As Range
        
    ' Excel にあらかじめ "source_table", "target_table" という名前でテーブルを定義しておく
    Set source_table = Range("source_table")
    Set target_table = Range("target_table")

    ' source_table上の読込インデックス位置 (r, c)
    Dim r, c As Long
    
    ' target_table上の書込インデックス位置 (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    ' その人の兼務数（本務除く）
    Dim ConcurrentPosts  As Integer
    ConcurrentPosts = 0

    ' main 処理
    For r = 1 To source_table.Rows.Count
        For c = 1 To source_table.Columns.Count
            
            ' prefix field
            If c <= 5 Then
                target_table(target_r, target_c) = source_table(r, c)
            
            ' 兼務1所属
            ElseIf c = 6 Then
                target_r = target_r + 1
                target_c = target_c - 3
                
                If source_table(r, c) <> "" Then
                    ConcurrentPosts = 1
                    target_table(target_r, target_c) = source_table(r, c)
                    
                    Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "（兼務１）")
                    
                End If
                        
            ' 兼務1所属長
            ElseIf c = 7 Then
                If ConcurrentPosts >= 1 Then
                    target_table(target_r, target_c) = source_table(r, c)
                End If
            
                target_r = target_r - 1
                target_c = target_c + 1
            
            ' 兼務2所属
            ElseIf c = 8 Then
                target_r = target_r + 2
                target_c = target_c - 3
                
                If source_table(r, c) <> "" Then
                    ConcurrentPosts = 2
                    target_table(target_r, target_c) = source_table(r, c)
                    
                    Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "（兼務２）")
                End If
            
            ' 兼務2所属長
            ElseIf c = 9 Then
                If ConcurrentPosts >= 2 Then
                    target_table(target_r, target_c) = source_table(r, c)
                End If
            
                target_r = target_r - 2
                target_c = target_c + 1
            
            ' 兼務3所属
            ElseIf c = 10 Then
                target_r = target_r + 3
                target_c = target_c - 3
                
                If source_table(r, c) <> "" Then
                    ConcurrentPosts = 3
                    target_table(target_r, target_c) = source_table(r, c)
                
                    Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "（兼務３）")
                End If
                
            ' 兼務3所属長
            ElseIf c = 11 Then
                If ConcurrentPosts >= 3 Then
                    target_table(target_r, target_c) = source_table(r, c)
                End If
            
                target_r = target_r - 3
                target_c = target_c + 1
            
            ' postfix field
            ElseIf c = 12 Then
                target_table(target_r, target_c) = source_table(r, c)
                target_c = 0 ' 最後の列なので一番左に戻る(後で +1 する)
                
                target_r = target_r + ConcurrentPosts
                ConcurrentPosts = 0
                                        
            End If
            
            target_c = target_c + 1
        Next
        
        target_r = target_r + 1
    Next

    ' 条件付書式をクリア＆設定
    Set target_table = Range("target_table") ' target_table を拡張された領域を含めて再定義
    target_table.ListObject.Range.FormatConditions.Delete ' 条件付書式クリア
    
    ' その行の社員列が（空白であれば）その行の上側の罫線を無くす
    With target_table.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($A2)")
        .Borders(xlTop).LineStyle = xlNone
    End With

    ' target_table の最終位置のインデックスを取得 (last_target_r, last_tareget_c)
    Dim last_target_r, last_target_c As Integer
    last_target_r = target_r - 1
    last_target_c = 12
    
    ' 所属＆所属長列に関しては、条件付書式クリア（Cells 直指定なのでHeader行分 +1 する）
    With Range(Cells(1, 3), Cells(last_target_r + 1, 4))
        .FormatConditions.Delete
    End With

End Sub

'
' 指定したCell (Range) に引数の文字列（ラベル）を右寄せした上でセットする
'
Sub SetConcurrentPostsLabel(target As Range, label As String)

    With target
        .Value = label
        .HorizontalAlignment = xlRight
    End With
    
End Sub
