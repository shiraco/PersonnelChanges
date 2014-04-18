Attribute VB_Name = "ConverterForAllEmployees"

Option Explicit

'
' ���ɕ��񂾌����̐l�̏��������c�ɕ��ׂ�悤�ɕϊ�����
' ��������3�܂őΉ�
'
' author koji shiraishi
' since 2014/04/18
' version 1.1
'
Sub ConvertConcurrentPostTableForAllEmployees()

    '=======================================================================================================
    ' �J�n����
    '=======================================================================================================

    ' �ydebug�z�������Ԃ̌v��start
    Dim startTime, stopTime As Variant
    startTime = Time '����������s���Ԃ̃J�E���g���J�n���܂�

    ' �ydebug�z�`��OFF
    Application.ScreenUpdating = False

    '=======================================================================================================
    ' �����ݒ�
    '=======================================================================================================

    ' Excel�e�[�u���͈̔͂̒�`
    ' Excel �ɂ��炩���� "source_table", "target_table" �Ƃ������O�Ńe�[�u�����`���Ă���
    Dim sourceTable As Range
    Dim targetTable As Range
    Set sourceTable = Range("source_table")
    Set targetTable = Range("target_table")

    ' �p�����[�^
    Const STR_SPACE_INDENT As String = "�@�@"             ' �C���f���g�����`
    Const STR_FIRST_BASE_CELL_ROW As String = "$A4"       ' �����ݒ�̊�Ƃ���ō���̃Z������̗�Q��

    ' ���ꂼ��̗�̈ʒu�i�C���f�b�N�X�j��萔�iConst�j�Ƃ��Ē�`
    ' COL_S_* �� sourceTable �i�ϊ��O�e�[�u���j�ł̗�ʒu
    Const COL_S_COMMON_PREFIX_START As Integer = 1        ' �Ј��ԍ�
    Const COL_S_COMMON_PREFIX_END As Integer = 4          ' �V��������

    Const COL_S_NEW_PREFIX_START As Integer = 5           ' �V����
    Const COL_S_NEW_PREFIX_END As Integer = 9             ' �V���Ə�

    Const COL_S_NEW_REPEAT_1A As Integer = 10             ' �V���������P
    Const COL_S_NEW_REPEAT_1B As Integer = 11             ' �V�����������P
    Const COL_S_NEW_REPEAT_2A As Integer = 12             ' �V���������Q
    Const COL_S_NEW_REPEAT_2B As Integer = 13             ' �V�����������Q
    Const COL_S_NEW_REPEAT_3A As Integer = 14             ' �V���������R
    Const COL_S_NEW_REPEAT_3B As Integer = 15             ' �V�����������R

    Const COL_S_NEW_SUFFIX_START As Integer = 16          ' �V���Џo����
    Const COL_S_NEW_SUFFIX_END As Integer = 18            ' �V�o������

    'COL_T_* �� targetTable �i�ϊ���e�[�u���j�ł̗�ʒu
    Const COL_T_NEW_CONCURRENT_POST_LABEL As Integer = 2  ' ����

    Const COL_T_NEW_UNIFY1_A As Integer = 3               ' �V����
    Const COL_T_NEW_UNIFY1_B As Integer = 4               ' �V�����g�D��

    '=======================================================================================================
    ' main ����
    '=======================================================================================================

    ' �s�p�����[�^
    ' ���̐l�̌������i�{�������j
    Dim newConcurrentPosts As Integer
    newConcurrentPosts = 0

    ' ���̐l���g�p����s���i�ʏ�1�A�������ɂ���đ����j
    Dim usingRows, newUsingRows As Integer
    usingRows = 1
    newUsingRows = 1

    ' sourceTable��̓Ǎ��C���f�b�N�X�ʒu (r, c)
    Dim r, c As Long

    ' targetTable��̏����C���f�b�N�X�ʒu (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    For r = 1 To sourceTable.Rows.Count
        For c = 1 To sourceTable.columns.Count

            ' �X�L�b�v�Ώۂ̗�ł���Ή������Ȃ�
            ' FIXME ���������܂Ƃ��ȃX�L�b�v�̂������ŃX�L�b�v������
            If c = 3 Or c = 4 Or c = 6 Or c = 7 Or c = 17 Or c = 18 Then

                'NOP

            ' �X�L�b�v�ΏۊO
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

                ' �V����1����
                ElseIf c = COL_S_NEW_REPEAT_1A Then
                    target_c = target_c - 2
                    target_c = target_c - 1

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 1

                        target_r = target_r + newConcurrentPosts

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "�i�����P�j")
                        targetTable(target_r, target_c) = STR_SPACE_INDENT & sourceTable(r, c)
                    End If

                ' �V����1������
                ElseIf c = COL_S_NEW_REPEAT_1B Then
                    If newConcurrentPosts >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - newConcurrentPosts
                    End If

                ' �V����2����
                ElseIf c = COL_S_NEW_REPEAT_2A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 2

                        target_r = target_r + newConcurrentPosts

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "�i�����Q�j")
                        targetTable(target_r, target_c) = STR_SPACE_INDENT & sourceTable(r, c)
                    End If

                ' �V����2������
                ElseIf c = COL_S_NEW_REPEAT_2B Then
                    If newConcurrentPosts >= 2 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - newConcurrentPosts
                    End If

                ' �V����3����
                ElseIf c = COL_S_NEW_REPEAT_3A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 3

                        target_r = target_r + newConcurrentPosts

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "�i�����R�j")
                        targetTable(target_r, target_c) = STR_SPACE_INDENT & sourceTable(r, c)
                    End If

                ' �V����3������
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

                target_c = target_c + 1 ' ��ړ�
            End If

        Next

        ' ���s����
        usingRows = newUsingRows
        target_c = 1                    ' ��ړ�
        target_r = target_r + usingRows ' �s�ړ�

        ' �s�p�����[�^�̃��Z�b�g
        newConcurrentPosts = 0
        newUsingRows = 1
        usingRows = 1

    Next

    '=======================================================================================================
    ' target_table �̏����̐ݒ�
    '=======================================================================================================

    ' �y�����̏������z�����t�������N���A
    Set targetTable = Range("target_table") ' targetTable �������O���g������Ă���̂ŁA���߂Ĕ͈͂��Ē�`����
    With targetTable.ListObject.Range
        .FormatConditions.Delete      ' ���ɏ����t��������`����Ă�����A�����t�������N���A����i�����t����Ȃ������̓N���A���Ȃ��j
    End With

    ' �y�S��̏����ݒ�i���̐l���ɂ�����2�s�ڈȍ~�j�z���̍s��"�Ј��ԍ�"�񂪁i�󔒂ł���΁j�����s�Ƃ݂Ȃ����̍s�̏㑤�̌r���𖳂���
    With targetTable.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(" & STR_FIRST_BASE_CELL_ROW & ")")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' �y������̏����ݒ�i���̐l���ɂ�����2�s�ڈȍ~�j�z
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_A), STR_FIRST_BASE_CELL_ROW, "C4")
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_B), STR_FIRST_BASE_CELL_ROW, "C4")

    '=======================================================================================================
    ' �I������
    '=======================================================================================================

    ' �ydebug�z�`��ON
    Application.ScreenUpdating = True

    ' �ydebug�z�������Ԃ̌v��end
    stopTime = Time
    stopTime = stopTime - startTime
    MsgBox "���v���Ԃ�" & Minute(stopTime) & "��" & Second(stopTime) & "�b �ł���"

End Sub

'
' �w�肵��Cell (�����̉��̗�) �Ɉ����̕�����i�������x���j���E�񂹂�����ŃZ�b�g����T�u���[�`��
'
Sub SetConcurrentPostLabel(target As Range, label As String)

    With target
        .Value = label
        .HorizontalAlignment = xlRight
    End With

End Sub

'
' �w�肵��Columns (�����A��E��) �ɏ����t�������Z�b�g����T�u���[�`��
'
Sub SetConcurrentPostFormatConditions(columns As Range, referenceCell1Str As String, referenceCell2Str As String)

    With columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), NOT(ISBLANK(" & referenceCell2Str & ")))").Borders(xlTop).LineStyle = xlDash
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), ISBLANK(" & referenceCell2Str & "))").Borders(xlTop).LineStyle = xlNone
    End With

End Sub

