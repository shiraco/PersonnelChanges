Attribute VB_Name = "ConcurrentPostTableConverter"


Option Explicit

'
' ���ɕ��񂾌����̐l�̏��������c�ɕ��ׂ�悤�ɕϊ�����
'
' author koji shiraishi
' since 2014/03/12
'
Sub ConvertConcurrentPostTable()
    ' Excel�e�[�u���͈̔͂̒�`
    Dim source_table As Range
    Dim target_table As Range
        
    ' Excel �ɂ��炩���� "source_table", "target_table" �Ƃ������O�Ńe�[�u�����`���Ă���
    Set source_table = Range("source_table")
    Set target_table = Range("target_table")

    ' source_table��̓Ǎ��C���f�b�N�X�ʒu (r, c)
    Dim r, c As Long
    
    ' target_table��̏����C���f�b�N�X�ʒu (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    ' ���̐l�̌������i�{�������j
    Dim ConcurrentPosts  As Integer
    ConcurrentPosts = 0

    ' main ����
    For r = 1 To source_table.Rows.Count
        For c = 1 To source_table.Columns.Count
            
            ' prefix field
            If c <= 5 Then
                target_table(target_r, target_c) = source_table(r, c)
            
            ' ����1����
            ElseIf c = 6 Then
                target_r = target_r + 1
                target_c = target_c - 3
                
                If source_table(r, c) <> "" Then
                    ConcurrentPosts = 1
                    target_table(target_r, target_c) = source_table(r, c)
                    
                    Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "�i�����P�j")
                    
                End If
                        
            ' ����1������
            ElseIf c = 7 Then
                If ConcurrentPosts >= 1 Then
                    target_table(target_r, target_c) = source_table(r, c)
                End If
            
                target_r = target_r - 1
                target_c = target_c + 1
            
            ' ����2����
            ElseIf c = 8 Then
                target_r = target_r + 2
                target_c = target_c - 3
                
                If source_table(r, c) <> "" Then
                    ConcurrentPosts = 2
                    target_table(target_r, target_c) = source_table(r, c)
                    
                    Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "�i�����Q�j")
                End If
            
            ' ����2������
            ElseIf c = 9 Then
                If ConcurrentPosts >= 2 Then
                    target_table(target_r, target_c) = source_table(r, c)
                End If
            
                target_r = target_r - 2
                target_c = target_c + 1
            
            ' ����3����
            ElseIf c = 10 Then
                target_r = target_r + 3
                target_c = target_c - 3
                
                If source_table(r, c) <> "" Then
                    ConcurrentPosts = 3
                    target_table(target_r, target_c) = source_table(r, c)
                
                    Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "�i�����R�j")
                End If
                
            ' ����3������
            ElseIf c = 11 Then
                If ConcurrentPosts >= 3 Then
                    target_table(target_r, target_c) = source_table(r, c)
                End If
            
                target_r = target_r - 3
                target_c = target_c + 1
            
            ' postfix field
            ElseIf c = 12 Then
                target_table(target_r, target_c) = source_table(r, c)
                target_c = 0 ' �Ō�̗�Ȃ̂ň�ԍ��ɖ߂�(��� +1 ����)
                
                target_r = target_r + ConcurrentPosts
                ConcurrentPosts = 0
                                        
            End If
            
            target_c = target_c + 1
        Next
        
        target_r = target_r + 1
    Next

    ' �����t�������N���A���ݒ�
    Set target_table = Range("target_table") ' target_table ���g�����ꂽ�̈���܂߂čĒ�`
    target_table.ListObject.Range.FormatConditions.Delete ' �����t�����N���A
    
    ' ���̍s�̎Ј��񂪁i�󔒂ł���΁j���̍s�̏㑤�̌r���𖳂���
    With target_table.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($A2)")
        .Borders(xlTop).LineStyle = xlNone
    End With

    ' target_table �̍ŏI�ʒu�̃C���f�b�N�X���擾 (last_target_r, last_tareget_c)
    Dim last_target_r, last_target_c As Integer
    last_target_r = target_r - 1
    last_target_c = 12
    
    ' ��������������Ɋւ��ẮA�����t�����N���A�iCells ���w��Ȃ̂�Header�s�� +1 ����j
    With Range(Cells(1, 3), Cells(last_target_r + 1, 4))
        .FormatConditions.Delete
    End With

End Sub

'
' �w�肵��Cell (Range) �Ɉ����̕�����i���x���j���E�񂹂�����ŃZ�b�g����
'
Sub SetConcurrentPostsLabel(target As Range, label As String)

    With target
        .Value = label
        .HorizontalAlignment = xlRight
    End With
    
End Sub
