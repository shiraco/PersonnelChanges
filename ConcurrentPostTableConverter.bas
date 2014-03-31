Attribute VB_Name = "ConcurrentPostTableConverter"

Option Explicit

'
' ���ɕ��񂾌����̐l�̏��������c�ɕ��ׂ�悤�ɕϊ�����
' ��������3�܂őΉ�
'
' author koji shiraishi
' since 2014/03/31
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
    Const COL_S_COMMON_PREFIX_START As Integer = 1   ' ���R���̂P
    Const COL_S_COMMON_PREFIX_END As Integer = 6   ' �V��������
    
    Const COL_S_AFT_PREFIX_START As Integer = 7   ' �V����
    Const COL_S_AFT_PREFIX_END As Integer = 12    ' �V���Ə�

    Const COL_S_AFT_REPEAT_1A As Integer = 13     ' �V���������P
    Const COL_S_AFT_REPEAT_1B As Integer = 14     ' �V�����������P
    Const COL_S_AFT_REPEAT_2A As Integer = 15     ' �V���������Q
    Const COL_S_AFT_REPEAT_2B As Integer = 16     ' �V�����������Q
    Const COL_S_AFT_REPEAT_3A As Integer = 17     ' �V���������R
    Const COL_S_AFT_REPEAT_3B As Integer = 18     ' �V�����������R

    Const COL_S_AFT_SUFFIX_START As Integer = 19  ' �V���Џo����
    Const COL_S_AFT_SUFFIX_END As Integer = 21    ' �V�o������

    Const COL_S_BEF_PREFIX_START As Integer = 22  ' ������
    Const COL_S_BEF_PREFIX_END As Integer = 24  ' ����������
    
    Const COL_S_BEF_REPEAT_1A As Integer = 25     ' �����������P
    Const COL_S_BEF_REPEAT_1B As Integer = 26     ' �������������P
    Const COL_S_BEF_REPEAT_2A As Integer = 27     ' �����������Q
    Const COL_S_BEF_REPEAT_2B As Integer = 28     ' �������������Q
    Const COL_S_BEF_REPEAT_3A As Integer = 29     ' �����������R
    Const COL_S_BEF_REPEAT_3B As Integer = 30     ' �������������R
    
    Const COL_S_BEF_SUFFIX_START As Integer = 31  ' �����Џo����
    Const COL_S_BEF_SUFFIX_END As Integer = 31    ' ���o������
    
    'COL_T_* �� target_table �i�ϊ���e�[�u���j�ł̗�ʒu
    Const COL_T_AFT_UNIFY_A As Integer = 6        ' �V����
    Const COL_T_AFT_UNIFY_B As Integer = 7        ' �V�����g�D��
    
    Const COL_T_BEF_UNIFY_A As Integer = 10       ' ������
    Const COL_T_BEF_UNIFY_B As Integer = 11       ' �������g�D��

    ' Dim COL_S_SKIPS As Variant                ' �X�L�b�v�Ώۂ̗�i�萔�ł͂Ȃ����ǁA�ύX���Ȃ��̂ő啶���Ő錾�j
    ' COL_S_SKIPS = Array(6, 7, 9, 10, 20, 21)   ' �\�����A�V�{���A�V�O���[�h �A�V�E��A�V���З��́A�V�o������


    ' Excel �ɂ��炩���� "source_table", "target_table" �Ƃ������O�Ńe�[�u�����`���Ă���
    Set source_table = Range("source_table")
    Set target_table = Range("target_table")

    ' source_table��̓Ǎ��C���f�b�N�X�ʒu (r, c)
    Dim r, c As Long

    ' target_table��̏����C���f�b�N�X�ʒu (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    ' ���̐l�̐V�������i�{�������j
    Dim ConcurrentPosts, BefAftMaxConcurrentPosts  As Integer
    ConcurrentPosts = 0
    BefAftMaxConcurrentPosts = 0
    
    '=======================================================================================================
    ' main ����
    '=======================================================================================================
    
    Application.ScreenUpdating = False ' �`��OFF

    For r = 1 To source_table.Rows.Count
        For c = 1 To source_table.Columns.Count

            ' �X�L�b�v�Ώۂ̗�ł���Ή������Ȃ�
            ' �z��iCOL_S_SKIPS�j�Ƃ̔�r�̎d�����킩��Ȃ��̂ŁA�ׂ�����
            If c = 6 Or c = 7 Or c = 9 Or c = 10 Or c = 20 Or c = 21 Then
                'NOP

            ' �X�L�b�v�ΏۊO
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
                
                ' �V����1����
                ElseIf c = COL_S_AFT_REPEAT_1A Then
                    target_r = target_r + 1
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 1

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "�i�����P�j")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))

                    End If

                ' �V����1������
                ElseIf c = COL_S_AFT_REPEAT_1B Then
                    If ConcurrentPosts >= 1 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 1
                    target_c = target_c + 1

                ' �V����2����
                ElseIf c = COL_S_AFT_REPEAT_2A Then
                    target_r = target_r + 2
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 2

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "�i�����Q�j")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' �V����2������
                ElseIf c = COL_S_AFT_REPEAT_2B Then
                    If ConcurrentPosts >= 2 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 2
                    target_c = target_c + 1

                ' �V����3����
                ElseIf c = COL_S_AFT_REPEAT_3A Then
                    target_r = target_r + 3
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 3
                         
                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "�i�����R�j")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' �V����3������
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
                
                ' ������1����
                ElseIf c = COL_S_BEF_REPEAT_1A Then
                    target_r = target_r + 1
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 1

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "�i�����P�j")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))

                    End If

                ' ������1������
                ElseIf c = COL_S_BEF_REPEAT_1B Then
                    If ConcurrentPosts >= 1 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 1
                    target_c = target_c + 1

                ' ������2����
                ElseIf c = COL_S_BEF_REPEAT_2A Then
                    target_r = target_r + 2
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 2

                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "�i�����Q�j")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' ������2������
                ElseIf c = COL_S_BEF_REPEAT_2B Then
                    If ConcurrentPosts >= 2 Then
                        target_table(target_r, target_c) = source_table(r, c)
                    End If

                    target_r = target_r - 2
                    target_c = target_c + 1

                ' ������3����
                ElseIf c = COL_S_BEF_REPEAT_3A Then
                    target_r = target_r + 3
                    target_c = target_c - 3

                    If source_table(r, c) <> "" Then
                        ConcurrentPosts = 3
                         
                        Call SetConcurrentPostLabel(target_table(target_r, target_c - 1), "�i�����R�j")
                        Call SetConcurrentPost(target_table(target_r, target_c), source_table(r, c))
                    End If

                ' ������3������
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
                
                target_c = target_c + 1 ' ��ړ�
            End If

            ' �V�̉E�[(after's suffix end)�ł̏���
            If c = COL_S_AFT_SUFFIX_END Then
                BefAftMaxConcurrentPosts = ConcurrentPosts
                ConcurrentPosts = 0
            End If
            
            ' ���̉E�[(before's suffix end)�ł̏���
            If c = COL_S_BEF_SUFFIX_END Then
                If BefAftMaxConcurrentPosts < ConcurrentPosts Then
                    BefAftMaxConcurrentPosts = ConcurrentPosts
                End If
                ConcurrentPosts = 0
            End If

        Next
        
        ' ���s����
        target_c = 1                                   ' ��ړ�
        target_r = target_r + 1                        ' �s�ړ��i�ʏ핪�j
        target_r = target_r + BefAftMaxConcurrentPosts ' �s�ړ��i���������̉��Z�j
        BefAftMaxConcurrentPosts = 0
        
    Next

    
    '=======================================================================================================
    ' �����̐ݒ�
    '=======================================================================================================
    
    ' �ytarget_table�̏����̏������z�����t�������N���A���ݒ�
    Set target_table = Range("target_table") ' target_table ���g������Ă���̂ŁA���߂Ē�`����
    With target_table.ListObject.Range
        .FormatConditions.Delete      ' ���ɏ����t��������`����Ă�����A�����t�������N���A����
        .HorizontalAlignment = xlGeneral ' �������l�߂ɏ����ݒ�
    End With

    ' �y�����s�̑S�́i�S��j�̏����ݒ�z���̍s�̎Ј��񂪁i�󔒂ł���΁j���̍s�̏㑤�̌r���𖳂���
    With target_table.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($D4)")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' �y�����s�̏�����̏����ݒ�z
    ' target_table �̍ŏI�ʒu�̃C���f�b�N�X���擾 (last_target_r, last_tareget_c)
    Dim last_target_r As Integer
    last_target_r = target_r - 1

    ' ������Ɋւ��ẮA���̍s�̎Ј��񂪁i�󔒂ł���΁j���̍s�̏㑤�̌r����_���ɂ���
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_A), "$D4", "$F4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_B), "$D4", "$F4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_BEF_UNIFY_A), "$D4", "$J4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_BEF_UNIFY_B), "$D4", "$J4")

    Application.ScreenUpdating = True ' �`��ON

End Sub

'
' �w�肵��Cell (Range) �Ɉ����̕�����i���x���j���E�񂹂�����ŃZ�b�g����T�u���[�`��
'
Sub SetConcurrentPostLabel(target As Range, label As String)

    With target
        .Value = label
        .HorizontalAlignment = xlRight
    End With

End Sub

'
' �w�肵��Cell (Range) �ɏ������C���f���g�t�ŃZ�b�g����T�u���[�`��
'
Sub SetConcurrentPost(target As Range, PostName As String)

     target = "�@�@" & PostName ' �S�p�X�y�[�X�~�Q�ŃC���f���g

End Sub


'
' �w�肵��Columns (Range) �ɂ̏����t�������Z�b�g����T�u���[�`��
'
Sub SetConcurrentPostFormatConditions(Columns As Range, referenceCellStr As String, selfCellStr As String)
    
    With Columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCellStr & "), NOT(ISBLANK(" & selfCellStr & ")))").Borders(xlTop).LineStyle = xlDot
    End With

End Sub
