Attribute VB_Name = "ConcurrentPostTableConverter"

Option Explicit

'
' ���ɕ��񂾌����̐l�̏��������c�ɕ��ׂ�悤�ɕϊ�����
' ��������3�܂őΉ�
'
' author koji shiraishi
' since 2014/03/31
' version 1.0
'
Sub ConvertConcurrentPostTable()

    '=======================================================================================================
    ' �����ݒ�
    '=======================================================================================================

    ' �ydebug�z�������Ԃ̌v��start
    Dim StartTime, StopTime As Variant
    StartTime = Time '����������s���Ԃ̃J�E���g���J�n���܂�

    ' Excel�e�[�u���͈̔͂̒�`
    Dim source_table As Range
    Dim target_table As Range

    ' ���ꂼ��̗�̈ʒu�i�C���f�b�N�X�j��萔�iConst�j�Ƃ��Ē�`
    ' COL_S_* �� source_table �i�ϊ��O�e�[�u���j�ł̗�ʒu
    Const COL_S_COMMON_PREFIX_START As Integer = 1 ' �Ј��ԍ�
    Const COL_S_COMMON_PREFIX_END As Integer = 2   ' �V��������
    
    Const COL_S_AFT_PREFIX_START As Integer = 5   ' �V����
    Const COL_S_AFT_PREFIX_END As Integer = 9     ' �V���Ə�

    Const COL_S_AFT_REPEAT_1A As Integer = 10     ' �V���������P
    Const COL_S_AFT_REPEAT_1B As Integer = 11     ' �V�����������P
    Const COL_S_AFT_REPEAT_2A As Integer = 12     ' �V���������Q
    Const COL_S_AFT_REPEAT_2B As Integer = 13     ' �V�����������Q
    Const COL_S_AFT_REPEAT_3A As Integer = 14     ' �V���������R
    Const COL_S_AFT_REPEAT_3B As Integer = 15     ' �V�����������R

    Const COL_S_AFT_SUFFIX_START As Integer = 16  ' �V���Џo����
    Const COL_S_AFT_SUFFIX_END As Integer = 18    ' �V�o������

    'COL_T_* �� target_table �i�ϊ���e�[�u���j�ł̗�ʒu
    Const COL_T_AFT_UNIFY_A As Integer = 3        ' �V����
    Const COL_T_AFT_UNIFY_B As Integer = 4        ' �V�����g�D��

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
            If c = 3 Or c = 4 Or c = 6 Or c = 7 Or c = 17 Or c = 18 Then
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

                End If

                '----------------------------------------------------
                ' common process
                '----------------------------------------------------

                target_c = target_c + 1 ' ��ړ�
            End If

        Next

        ' ���s����
        target_c = 1                          ' ��ړ�
        target_r = target_r + 1               ' �s�ړ��i�ʏ핪�j
        target_r = target_r + concurrentPosts ' �s�ړ��i���������̉��Z�j
        concurrentPosts = 0
        
    Next

    '=======================================================================================================
    ' �����̐ݒ�
    '=======================================================================================================

    ' �ytarget_table�̏����̏������z�����t�������N���A���ݒ�
    Set target_table = Range("target_table") ' target_table ���g������Ă���̂ŁA���߂Ē�`����
    With target_table.ListObject.Range
        .FormatConditions.Delete      ' ���ɏ����t��������`����Ă�����A�����t�������N���A����
    End With

    ' �y�����s�̑S�́i�S��j�̏����ݒ�z���̍s�̎Ј��񂪁i�󔒂ł���΁j���̍s�̏㑤�̌r���𖳂���
    With target_table.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK($A4)")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' �y�����s�̏�����̏����ݒ�z
    ' ������Ɋւ��ẮA���̍s�̎Ј��񂪁i�󔒂ł���΁j���̍s�̏㑤�̌r����_���ɂ���
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_A), "$A4")
    Call SetConcurrentPostFormatConditions(target_table.Columns(COL_T_AFT_UNIFY_B), "$A4")

    Application.ScreenUpdating = True ' �`��ON

    ' �ydebug�z�������Ԃ̌v��end
    StopTime = Time
    StopTime = StopTime - StartTime
    MsgBox "���v���Ԃ�" & Minute(StopTime) & "��" & Second(StopTime) & "�b �ł���"

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
Sub SetConcurrentPostFormatConditions(Columns As Range, referenceCellStr As String)
    
    With Columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(" & referenceCellStr & ")").Borders(xlTop).LineStyle = xlDot
    End With

End Sub
