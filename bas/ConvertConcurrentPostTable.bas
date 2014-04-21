Attribute VB_Name = "ConvertConcurrentPostTable"

Option Explicit

'
' ���ɕ��񂾌����̐l�̏��������c�ɕ��ׂ�悤�ɕϊ�����
' ��������3�A�o������2�܂őΉ�
'
' Args:
'  sourceReadingType  "NEW_ONLY": �V�����̂ݓ]�L , "NEW_AND_OLD": �V�����������]�L
'
' author koji shiraishi
' since 2014/04/21
' version 1.1
'
Sub ConvertConcurrentPostTable(sourceReadingType As String)

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
    Const SPACE_INDENT_STR As String = "�@�@"             ' �C���f���g�����`

    ' ���ꂼ��̗�̈ʒu�i�C���f�b�N�X�j��萔�iConst�j�Ƃ��Ē�`
    ' COL_S_* �� sourceTable �i�ϊ��O�e�[�u���j�ł̗�ʒu
    Const COL_S_COMMON_PREFIX_START As Integer = 1        ' ���R����
    Const COL_S_COMMON_PREFIX_END As Integer = 4          ' ����

    Const COL_S_NEW_REPEAT1_0A As Integer = 5             ' �V�{������
    Const COL_S_NEW_REPEAT1_0B As Integer = 9             ' �V�{����E
    Const COL_S_NEW_REPEAT1_1A As Integer = 10            ' �V���������P
    Const COL_S_NEW_REPEAT1_1B As Integer = 14            ' �V������E�P
    Const COL_S_NEW_REPEAT1_2A As Integer = 15            ' �V���������Q
    Const COL_S_NEW_REPEAT1_2B As Integer = 19            ' �V������E�Q
    Const COL_S_NEW_REPEAT1_3A As Integer = 20            ' �V���������R
    Const COL_S_NEW_REPEAT1_3B As Integer = 24            ' �V������E�R

    Const COL_S_NEW_REPEAT2_1A As Integer = 28            ' �V�o����P
    Const COL_S_NEW_REPEAT2_2A As Integer = 32            ' �V�o����Q

    Const COL_S_NEW_SUFFIX_START As Integer = 33          ' �V���Ə�
    Const COL_S_NEW_SUFFIX_END As Integer = 35            ' �V�E��*�폜
    Const COL_S_NEW_END As Integer = COL_S_NEW_SUFFIX_END ' �V�E��*�폜

    Const COL_S_OLD_REPEAT1_0A As Integer = 36            ' ���{������
    Const COL_S_OLD_REPEAT1_0B As Integer = 40            ' ���{����E
    Const COL_S_OLD_REPEAT1_1A As Integer = 41            ' �����������P
    Const COL_S_OLD_REPEAT1_1B As Integer = 45            ' ��������E�P
    Const COL_S_OLD_REPEAT1_2A As Integer = 46            ' �����������Q
    Const COL_S_OLD_REPEAT1_2B As Integer = 50            ' ��������E�Q
    Const COL_S_OLD_REPEAT1_3A As Integer = 51            ' �����������R
    Const COL_S_OLD_REPEAT1_3B As Integer = 55            ' ��������E�R

    Const COL_S_OLD_REPEAT2_1A As Integer = 59            ' ���o����P
    Const COL_S_OLD_REPEAT2_2A As Integer = 63            ' ���o����Q

    Const COL_S_OLD_SUFFIX_START As Integer = 64          ' �����Ə�
    Const COL_S_OLD_SUFFIX_END As Integer = 66            ' ���E��*�폜
    Const COL_S_OLD_END As Integer = COL_S_OLD_SUFFIX_END ' ���E��*�폜

    'COL_T_* �� targetTable �i�ϊ���e�[�u���j�ł̗�ʒu
    Const COL_T_NEW_REASON1 As Integer = 1                ' ���R���̂P
    Const COL_T_NEW_REASON2 As Integer = 2                ' ���R���̂Q
    Const COL_T_NEW_CHANGE_DATE As Integer = 3            ' ���ߓ�
    Const COL_T_NEW_CONCURRENT_POST_LABEL As Integer = 5  ' ����
    Const COL_T_NEW_UNIFY1_A As Integer = 6               ' �V����
    Const COL_T_NEW_UNIFY1_B As Integer = 7               ' �V��E
    Const COL_T_NEW_UNIFY2_A As Integer = 8               ' �V�o����
    Const COL_T_OLD_UNIFY1_A As Integer = 10              ' ������
    Const COL_T_OLD_UNIFY1_B As Integer = 11              ' ����E
    Const COL_T_OLD_UNIFY2_A As Integer = 12              ' ���o����

    Const CELL_T_NEW_REASON1_STR As String = "A7"         ' ���R���̂P
    Const CELL_T_NEW_REASON2_STR As String = "B7"         ' ���R���̂Q
    Const CELL_T_NEW_CHANGE_DATE_STR As String = "C7"     ' ���ߓ�
    Const CELL_T_EMPLOYEE_NO_STR As String = "$D7"        ' �Ј��ԍ���̍ŏ�s
    Const CELL_T_NEW_POST_NAME_STR As String = "F7"       ' �V�����̍ŏ�s
    Const CELL_T_NEW_ACOMPANY_NAME_STR As String = "H7"   ' �V�o����̍ŏ�s
    Const CELL_T_OLD_POST_NAME_STR As String = "J7"       ' �������̍ŏ�s
    Const CELL_T_OLD_ACOMPANY_NAME_STR As String = "L7"   ' ���o����̍ŏ�s

    '=======================================================================================================
    ' source_table -> target_table �ւ̓]�L����
    '=======================================================================================================

    ' �s�p�����[�^
    ' ���̐l�̌������i�{�������j
    Dim newConcurrentPosts, oldConcurrentPosts As Integer
    newConcurrentPosts = 0
    oldConcurrentPosts = 0
    ' ���̐l�̏o����А�
    Dim newAssigneeCompanies, oldAssigneeCompanies As Integer
    newAssigneeCompanies = 0
    oldAssigneeCompanies = 0
    ' ���̐l���g�p����s���i�ʏ�1�A�������A�o�����ɂ���đ����j
    Dim usingRows, newUsingRows, oldUsingRows As Integer
    usingRows = 1
    newUsingRows = 1
    oldUsingRows = 1
    ' ���̐l�̖{���ɂ����Ďg�p����s���i�o�����ɉ����đ����j
    Dim newUsingRowsInMainPost, oldUsingRowsInMainPost As Integer
    newUsingRowsInMainPost = 1
    oldUsingRowsInMainPost = 1

    ' sourceTable��̓Ǎ��C���f�b�N�X�ʒu (r, c)
    Dim r, c As Long

    ' targetTable��̏����C���f�b�N�X�ʒu (target_r, tareget_c)
    Dim target_r, target_c As Long
    target_r = 1
    target_c = 1

    ' for ���[�v�� index
    Dim i As Long

    Dim sourceTableReadingCols As Long
    sourceTableReadingCols = IIf(sourceReadingType = "NEW_ONLY", COL_S_NEW_END, sourceTable.columns.Count)

    For r = 1 To sourceTable.Rows.Count
        For c = 1 To sourceTableReadingCols

            ' �X�L�b�v�Ώۂ̗�ł���Ή������Ȃ�
            ' FIXME ���������܂Ƃ��ȃX�L�b�v�̂������ŃX�L�b�v������
            ' �� source_table �̗񂪕ύX�ɂȂ����ꍇ�A�v�C��
            ' �V�̃X�L�b�v�Ώ�
            If COL_S_NEW_REPEAT1_0A < c And c < COL_S_NEW_REPEAT1_0B Or _
               COL_S_NEW_REPEAT1_1A < c And c < COL_S_NEW_REPEAT1_1B Or _
               COL_S_NEW_REPEAT1_2A < c And c < COL_S_NEW_REPEAT1_2B Or _
               COL_S_NEW_REPEAT1_3A < c And c < COL_S_NEW_REPEAT1_3B Or _
               c = 25 Or c = 26 Or c = 27 Or c = 29 Or c = 30 Or c = 31 Or c = 34 Or c = 35 Then

                'NOP

            ' ���̃X�L�b�v�Ώ�
            ElseIf sourceReadingType <> "NEW_ONLY" And _
               (COL_S_OLD_REPEAT1_0A < c And c < COL_S_OLD_REPEAT1_0B Or _
                COL_S_OLD_REPEAT1_1A < c And c < COL_S_OLD_REPEAT1_1B Or _
                COL_S_OLD_REPEAT1_2A < c And c < COL_S_OLD_REPEAT1_2B Or _
                COL_S_OLD_REPEAT1_3A < c And c < COL_S_OLD_REPEAT1_3B Or _
                c = 56 Or c = 57 Or c = 58 Or c = 60 Or c = 61 Or c = 62 Or c = 65 Or c = 66) Then

                'NOP

            ' �X�L�b�v�ΏۊO��
            Else
                '----------------------------------------------------
                ' common's field
                '----------------------------------------------------

                If COL_S_COMMON_PREFIX_START <= c And c <= COL_S_COMMON_PREFIX_END Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                    ' �u���R���̂Q�v��̃X�y�[�X���P�񕪊m��
                    If c = COL_T_NEW_REASON1 Then
                        target_c = target_c + 1 ' ��ړ�
                    End If

                '----------------------------------------------------
                ' new's field
                '----------------------------------------------------

                ' �V�{������
                ElseIf c = COL_S_NEW_REPEAT1_0A Then
                    ' �o�������O����
                    If sourceTable(r, COL_S_NEW_REPEAT2_1A) <> "" Then
                        newAssigneeCompanies = IIf(sourceTable(r, COL_S_NEW_REPEAT2_2A) <> "", 2, 1)
                        newUsingRowsInMainPost = IIf(newAssigneeCompanies >= 2, newAssigneeCompanies, 1)

                        ' target_table �͈̔͂������g�������邽�߂ɃZ���Ƀ_�~�[�l�iSPACE_INDENT_STR�j��ݒ�
                        If newUsingRowsInMainPost >= 2 Then
                            For i = 1 To newUsingRowsInMainPost - 1
                                targetTable(target_r + i, COL_T_NEW_UNIFY2_A) = SPACE_INDENT_STR
                            Next i
                        End If
                    End If

                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' �V�{����E
                ElseIf c = COL_S_NEW_REPEAT1_0B Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' �V����1����
                ElseIf c = COL_S_NEW_REPEAT1_1A Then
                    target_c = target_c - 2 ' ������ɖ߂�

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 1

                        target_r = target_r + (newConcurrentPosts + (newUsingRowsInMainPost - 1)) ' �����s�։��s����

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "�i�����P�j")
                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' �V����1��E
                ElseIf c = COL_S_NEW_REPEAT1_1B Then
                    If newConcurrentPosts >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (newConcurrentPosts + (newUsingRowsInMainPost - 1)) ' �����s�։��s�������߂�
                    End If

                ' �V����2����
                ElseIf c = COL_S_NEW_REPEAT1_2A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 2

                        target_r = target_r + (newConcurrentPosts + (newUsingRowsInMainPost - 1))

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "�i�����Q�j")
                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' �V����2��E
                ElseIf c = COL_S_NEW_REPEAT1_2B Then
                    If newConcurrentPosts >= 2 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (newConcurrentPosts + (newUsingRowsInMainPost - 1))
                    End If

                ' �V����3����
                ElseIf c = COL_S_NEW_REPEAT1_3A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        newConcurrentPosts = 3

                        target_r = target_r + (newConcurrentPosts + (newUsingRowsInMainPost - 1))

                        Call SetConcurrentPostLabel(targetTable(target_r, COL_T_NEW_CONCURRENT_POST_LABEL), "�i�����R�j")
                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' �V����3��E
                ElseIf c = COL_S_NEW_REPEAT1_3B Then
                    If newConcurrentPosts >= 3 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (newConcurrentPosts + (newUsingRowsInMainPost - 1))
                    End If

                ' �V�o����1
                ElseIf c = COL_S_NEW_REPEAT2_1A Then
                    If newAssigneeCompanies >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)
                    End If

                ' �V�o����2
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

                ' ���{������
                ElseIf c = COL_S_OLD_REPEAT1_0A Then
                    ' �o�������O����
                    If sourceTable(r, COL_S_OLD_REPEAT2_1A) <> "" Then
                        oldAssigneeCompanies = IIf(sourceTable(r, COL_S_OLD_REPEAT2_2A) <> "", 2, 1)
                        oldUsingRowsInMainPost = IIf(oldAssigneeCompanies >= 2, oldAssigneeCompanies, 1)

                        ' target_table �͈̔͂������g�������邽�߂ɃZ���Ƀ_�~�[�l�iSPACE_INDENT_STR�j��ݒ�
                        If oldUsingRowsInMainPost >= 2 Then
                            For i = 1 To oldUsingRowsInMainPost - 1
                                targetTable(target_r + i, COL_T_OLD_UNIFY2_A) = SPACE_INDENT_STR
                            Next i
                        End If
                    End If

                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' ���{����E
                ElseIf c = COL_S_OLD_REPEAT1_0B Then
                    targetTable(target_r, target_c) = sourceTable(r, c)

                ' ������1����
                ElseIf c = COL_S_OLD_REPEAT1_1A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        oldConcurrentPosts = 1

                        target_r = target_r + (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))

                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' ������1��E
                ElseIf c = COL_S_OLD_REPEAT1_1B Then
                    If oldConcurrentPosts >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))
                    End If

                ' ������2����
                ElseIf c = COL_S_OLD_REPEAT1_2A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        oldConcurrentPosts = 2

                        target_r = target_r + (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))

                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' ������2��E
                ElseIf c = COL_S_OLD_REPEAT1_2B Then
                    If oldConcurrentPosts >= 2 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))
                    End If

                '������3����
                ElseIf c = COL_S_OLD_REPEAT1_3A Then
                    target_c = target_c - 2

                    If sourceTable(r, c) <> "" Then
                        oldConcurrentPosts = 3

                        target_r = target_r + (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))

                        targetTable(target_r, target_c) = SPACE_INDENT_STR & sourceTable(r, c)
                    End If

                ' ������3��E
                ElseIf c = COL_S_OLD_REPEAT1_3B Then
                    If oldConcurrentPosts >= 3 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)

                        target_r = target_r - (oldConcurrentPosts + (oldUsingRowsInMainPost - 1))
                    End If

                ' ���o����1
                ElseIf c = COL_S_OLD_REPEAT2_1A Then
                    If oldAssigneeCompanies >= 1 Then
                        targetTable(target_r, target_c) = sourceTable(r, c)
                    End If

                ' ���o����2
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

                target_c = target_c + 1 ' ��ړ�
            End If

        Next

        ' ���s����
        usingRows = IIf(newUsingRows >= oldUsingRows, newUsingRows, oldUsingRows)
        target_c = 1                    ' ��ړ�
        target_r = target_r + usingRows ' �s�ړ�

        ' �s�p�����[�^�̃��Z�b�g
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
    ' target_table �̏����̐ݒ�
    '=======================================================================================================

    ' �y�����̏������z�����t�������N���A
    Set targetTable = Range("target_table") ' targetTable �������O���g������Ă���̂ŁA���߂Ĕ͈͂��Ē�`����
    With targetTable.ListObject.Range
        .FormatConditions.Delete      ' ���ɏ����t��������`����Ă�����A�����t�������N���A����i�����t����Ȃ������̓N���A���Ȃ��j
    End With

    ' �y�S��̏����ݒ�i���̐l���ɂ�����2�s�ڈȍ~���ΏۂƂȂ�j�z���̍s��"�Ј��ԍ�"�񂪁i�󔒂ł���΁j�����s�Ƃ݂Ȃ����̍s�̏㑤�̌r���𖳂���
    With targetTable.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(" & CELL_T_EMPLOYEE_NO_STR & ")")
        .Borders(xlTop).LineStyle = xlLineStyleNone
    End With

    ' �y���R���́A���ߓ���̏����ݒ�z���g�̃Z�����󔒂̏ꍇ�A�㑤�̌r���𖳂���
    Call SetFormatConditionsTopLineNone(targetTable.columns(COL_T_NEW_REASON1), CELL_T_NEW_REASON1_STR)
    Call SetFormatConditionsTopLineNone(targetTable.columns(COL_T_NEW_REASON2), CELL_T_NEW_REASON2_STR)
    Call SetFormatConditionsTopLineNone(targetTable.columns(COL_T_NEW_CHANGE_DATE), CELL_T_NEW_CHANGE_DATE_STR)

    ' �y������̏����ݒ�i���̐l���ɂ�����2�s�ڈȍ~���ΏۂƂȂ�j�z
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_NEW_POST_NAME_STR)
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_NEW_UNIFY1_B), CELL_T_EMPLOYEE_NO_STR, CELL_T_NEW_POST_NAME_STR)
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_OLD_UNIFY1_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_OLD_POST_NAME_STR)
    Call SetConcurrentPostFormatConditions(targetTable.columns(COL_T_OLD_UNIFY1_B), CELL_T_EMPLOYEE_NO_STR, CELL_T_OLD_POST_NAME_STR)

    ' �y�o�����̏����ݒ�i���̐l���ɂ�����2�s�ڈȍ~���ΏۂƂȂ�j�z
    Call SetAssigneeCompanyFormatConditions(targetTable.columns(COL_T_NEW_UNIFY2_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_NEW_POST_NAME_STR, CELL_T_NEW_ACOMPANY_NAME_STR)
    Call SetAssigneeCompanyFormatConditions(targetTable.columns(COL_T_OLD_UNIFY2_A), CELL_T_EMPLOYEE_NO_STR, CELL_T_OLD_POST_NAME_STR, CELL_T_OLD_ACOMPANY_NAME_STR)

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
' �w�肵��Columns (���R���́A���ߓ�) �ɏ����t�������Z�b�g����T�u���[�`��
'
Sub SetFormatConditionsTopLineNone(columns As Range, referenceCell1Str As String)

    With columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(NOT(ISBLANK(" & referenceCell1Str & ")))").Borders(xlTop).LineStyle = xlContinuous
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "))").Borders(xlTop).LineStyle = xlNone
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

'
' �w�肵��Columns (�o�����) �ɏ����t�������Z�b�g����T�u���[�`��
'
Sub SetAssigneeCompanyFormatConditions(columns As Range, referenceCell1Str As String, referenceCell2Str As String, selfCellStr As String)

    With columns
        .FormatConditions.Delete
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), NOT(ISBLANK(" & referenceCell2Str & ")))").Borders(xlTop).LineStyle = xlDash
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), ISBLANK(" & referenceCell2Str & "), NOT(ISBLANK(" & selfCellStr & ")))").Borders(xlTop).LineStyle = xlDot
        .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISBLANK(" & referenceCell1Str & "), ISBLANK(" & referenceCell2Str & "), ISBLANK(" & selfCellStr & "))").Borders(xlTop).LineStyle = xlNone
    End With

End Sub
