Attribute VB_Name = "ConcurrentPostTableConverter"

Option Explicit

'
' ���ɕ��񂾌����̐l�̏��������c�ɕ��ׂ�悤�ɕϊ�����
' ��������3�܂őΉ�
'
' author koji shiraishi
' since 2014/03/20
'
Sub ConvertConcurrentPostTable()

    
    '=======================================================================================================
    ' �����ݒ�
    '=======================================================================================================
    
    ' Excel�e�[�u���͈̔͂̒�`
    Dim source_table As Range
    Dim target_table As Range

    ' ���ꂼ��̗�̈ʒu�i�C���f�b�N�X�j��萔�iConst�j�Ƃ��Ē�`
    ' COL_S_* �� source_table �i�ϊ��O�e�[�u���j�ł̗�ʒu
    Const COL_S_PREFIX_START As Integer = 1  ' �Ј��ԍ�
    Const COL_S_PREFIX_END As Integer = 9    ' �V���Ə�

    Const COL_S_REPEAT_1A As Integer = 10     ' �V���������P
    Const COL_S_REPEAT_1B As Integer = 11     ' �V�����������P
    Const COL_S_REPEAT_2A As Integer = 12     ' �V���������Q
    Const COL_S_REPEAT_2B As Integer = 13     ' �V�����������Q
    Const COL_S_REPEAT_3A As Integer = 14     ' �V���������R
    Const COL_S_REPEAT_3B As Integer = 15     ' �V�����������R

    Const COL_S_SUFFIX_START As Integer = 16  ' �V���Џo����
    Const COL_S_SUFFIX_END As Integer = 18    ' �V�o������

    ' COL_T_* �� target_table �i�ϊ���e�[�u���j�ł̗�ʒu
    Const COL_T_UNIFY_A As Integer = 3        ' �V����
    Const COL_T_UNIFY_B As Integer = 4        ' �V�����g�D��

    ' Dim COL_S_SKIPS As Variant                ' �X�L�b�v�Ώۂ̗�i�萔�ł͂Ȃ����ǁA�ύX���Ȃ��̂ő啶���Ő錾�j
    ' COL_S_SKIPS = Array(3, 4, 6, 7, 17, 18)   ' �\�����A�V�{���A�V�O���[�h �A�V�E��A�V���З��́A�V�o������


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

    
    '=======================================================================================================
    ' main ����
    '=======================================================================================================
    
    Application.ScreenUpdating = False ' �`��OFF

    For r = 1 To source_table.Rows.Count
        For c = 1 To source_table.Columns.Count

            ' �X�L�b�v�Ώۂ̗�ł���Ή������Ȃ�
            ' �z��iCOL_S_SKIPS�j�Ƃ̔�r�̎d�����킩��Ȃ��̂ŁA�ׂ�����
            If c = 3 Or c = 4 Or c = 6 Or c = 7 Or c = 17 Or c = 18 Then
                'NOP

            ' �X�L�b�v�ΏۊO
            Else
                ' prefix field
                If COL_S_PREFIX_START <= c And c <= COL_S_PREFIX_END Then
                    target_table(target_r, target_c) = source_table(r, c)

                ' ����1����
                ElseIf c = COL_S_REPEAT_1A Then
                    target_r = target_r + 1
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 1
                        target_table(target_r, target_c) = source_table(r, c)

                        Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "�i�����P�j")

                    End If

                ' ����1������
                ElseIf c = COL_S_REPEAT_1B Then
                    If ConcurrentPosts >= 1 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 1
                    target_c = target_c + 1

                ' ����2����
                ElseIf c = COL_S_REPEAT_2A Then
                    target_r = target_r + 2
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 2
                        target_table(target_r, target_c) = source_table(r, c)

                        Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "�i�����Q�j")
                    End If

                ' ����2������
                ElseIf c = COL_S_REPEAT_2B Then
                    If ConcurrentPosts >= 2 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 2
                    target_c = target_c + 1

                ' ����3����
                ElseIf c = COL_S_REPEAT_3A Then
                    target_r = target_r + 3
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 3
                        target_table(target_r, target_c) = source_table(r, c)

                        Call SetConcurrentPostsLabel(target_table(target_r, target_c - 1), "�i�����R�j")
                    End If

                ' ����3������
                ElseIf c = COL_S_REPEAT_3B Then
                    If ConcurrentPosts >= 3 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 3
                    target_c = target_c + 1

                ' postfix field
                ElseIf COL_S_SUFFIX_START <= c And c <= COL_S_SUFFIX_END Then
                    target_table(target_r, target_c) = source_table(r, c)

                End If

                target_c = target_c + 1
            End If

            If c = COL_S_SUFFIX_END Then
                target_c = 1                          ' �Ō�̗�Ȃ̂ň�ԍ��ɖ߂�
                target_r = target_r + ConcurrentPosts ' ���s
                ConcurrentPosts = 0
            End If

        Next

        target_r = target_r + 1
    Next

    
    '=======================================================================================================
    ' �����̐ݒ�
    '=======================================================================================================
    
    ' �ytarget_table�̏����̏������z�����t�������N���A���ݒ�
    Set target_table = Range("target_table") ' target_table ���g������Ă���̂ŁA���߂Ē�`����
    target_table.ListObject.Range.FormatConditions.Delete ' ���ɏ����t��������`����Ă�����A�����t�������N���A����

    ' �y�����s�̑S�́i�S��j�̏����ݒ�z���̍s�̎Ј��񂪁i�󔒂ł���΁j���̍s�̏㑤�̌r���𖳂���
    With target_table.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($A2)")
        .Borders(xlTop).LineStyle = xlNone
    End With

    ' �y�����s�̏�����̏����ݒ�z
    ' target_table �̍ŏI�ʒu�̃C���f�b�N�X���擾 (last_target_r, last_tareget_c)
    Dim last_target_r As Integer
    last_target_r = target_r - 1

    ' ������Ɋւ��ẮA�����t�����N���A
    ' ���̍s�̎Ј��񂪁i�󔒂ł���΁j���̍s�̏㑤�̌r���𖳂���
    With target_table.Range(Cells(1, COL_T_UNIFY_A), Cells(last_target_r, COL_T_UNIFY_B))
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($A2)").Borders(xlTop).LineStyle = xlDot
    End With

    Application.ScreenUpdating = True ' �`��ON

End Sub

'
' �w�肵��Cell (Range) �Ɉ����̕�����i���x���j���E�񂹂�����ŃZ�b�g����T�u���[�`��
'
Sub SetConcurrentPostsLabel(target As Range, label As String)

    With target
        .Value = label
        .HorizontalAlignment = xlRight
    End With

End Sub


