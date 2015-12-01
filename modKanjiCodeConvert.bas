Attribute VB_Name = "modKanjiCodeConvert"
'*********************************************************************************************
'*                            �����R�[�h�ϊ����W���[�� For VB6/VBA                           *
'*********************************************************************************************
'
'[�@�\/����]
'���̃��W���[����Shift-JIS/EUC/JIS/Unicode(UTF-8,UTF-16)�̑��ݕϊ����s���܂��B
'���܂��Ƃ��ĕ����R�[�h���ʊ֐�(��)���p�ӂ��܂����B
'��ʂ�e�X�g�͂��܂������ϊ����ꂽ���e���������Ƃ����ۏ؂͂ł��܂���B
'
'[�g����]
'�ϊ����f�[�^���o�C�g�^�̔z��String�^�ϐ��ɓǂݏo���Ă�����
'    KCConvert( �o�C�g�z��, �ϊ��������R�[�h, �ϊ��敶���R�[�h )
'�Ƃ��ČĂяo����OK�ł��D
'�����R�[�h�̔�����s�������Ƃ���
'    KCDetectCode( �o�C�g�z�� )
'�Ƃ��܂��D
'�Ȃ��G���[�����������Ƃ��́C�G���[�ԍ��ɒ萔KC_ERROR_INVALID���Z�b�g���܂��D
'�֐����Ăяo���O�ɃG���[�g���b�v�����Ă��������B
'
'[���쌠]
'���̃��W���[���̒��쌠�͕ێ����܂������ρA�Ĕz�z�A�����R�ɂ��g�����������B
'���̃��W���[�����g�p���Ă��Ȃ���������\�t�g�̒��쌠��A���̃��W���[�������ς������̂�
'���쌠�͂��Ȃ��ɂ���܂��B�Ĕz�z��]�ڂ̂Ƃ��̓��[���ňꌾ�A������������Ƃ��ꂵ���ł��B
'
'[�Ɛӎ���]
'���̃��W���[���̎g�p�ɂ�邢���Ȃ鑹�Q�ɑ΂��Ă��ӔC�𕉂��܂���B
'
'[���̑�]
'�ϊ��A���S���Y���ɊԈႢ��������܂�����A��҂܂ł��m�点���������B
'
'[�X�V����]
'Version 1.00   2003/10/08  �ꉞ����
'        1.20   2003/11/01  ����OUT�Ƃ����Ԉ�����T�O���g���Ă����̂Œ���
'        1.30   2004/03/05  ������
'        1.50   2005/03/03  UTF-8��Unicode�Ԃ̕ϊ��֐��AUTF-8�̐������`�F�b�N�֐�������
'        2.00   2008/07/13  KCConvert�֐��Ɉ�{��
'                           EUC��Shift-JIS�Ԃ̕ϊ��𒼐ڍs���CJIS��EUC�Ԃ̕ϊ��֐����폜
'                           UTF-16BE,UTF-16LE�ɑΉ�
'                           VBA�œ�����m�F
'                           EUC�̃V���O���V�t�g3�ɑΉ�
'                           ���ʊ֐�������
'
'                                   ==========================================================
'                                    330k http://www.330k.info/
'                                   ==========================================================
Option Explicit

'�����R�[�h������킷�񋓌^
Public Enum KCCode
    KC_UNKNOWN = 0              '�s��
    KC_SHIFTJIS = 10            'Shift-JIS
    KC_JIS = 20                 'JIS
    KC_EUC = 30                 'EUC-JP
    KC_UTF8N = 80               'UTF-8N
    KC_UTF8BOM = 81             'UTF-8(BOM��)
    KC_UNICODESTRING = 160      'VB��String�^,UCS2(UTF-16LE,BOM�Ȃ�,�T���Q�[�g�y�A��Ή�)
    KC_UTF16LE = 161            'UTF-16LE
    KC_UTF16BE = 162            'UTF-16BE
End Enum

'�G���[�R�[�h
Public Const KC_ERROR_INVALID As Long = 65000

'JIS�̃G�X�P�[�v�V�[�P���X
Private Const bytJISESC As Byte = &H1B      'JIS��ESC(�G�X�P�[�v�V�[�P���X)
Private Const bytKANJI1 As Byte = &H24      'JIS����IN1�o�C�g��"$"
Private Const bytKANJI2OLD As Byte = &H40   'JIS����IN2�o�C�g��(��JIS)"@"
Private Const bytKANJI2NEW As Byte = &H42   'JIS����IN2�o�C�g��(�VJIS)"B"
Private Const bytROME1 As Byte = &H28       'JIS���[�}��IN1�o�C�g��"("
Private Const bytROME2 As Byte = &H4A       'JIS���[�}��IN2�o�C�g��"J"
Private Const bytKATA1 As Byte = &H28       'JIS���p�J�^�J�iIN1�o�C�g��"("
Private Const bytKATA2 As Byte = &H49       'JIS���p�J�^�J�iIN2�o�C�g��"I"

'EUC�̃G�X�P�[�v�V�[�P���X
Private Const bytSS2 As Byte = &H8E         'Single Shift 2
Private Const bytSS3 As Byte = &H8F         'Single Shift 3

'UTF-8��BOM
Private Const bytUTF8BOM1 As Byte = &HEF
Private Const bytUTF8BOM2 As Byte = &HBB
Private Const bytUTF8BOM3 As Byte = &HBF

'UTF-16��BOM
Private Const bytUTF16BOM1 As Byte = &HFE
Private Const bytUTF16BOM2 As Byte = &HFF

'�ϊ��������ɓ����Ŏg�p����t���O
Private Enum EncodeMode
    mKanji = 1
    mRome = 2
    mKata = 3
End Enum

Public Function KCConvert(ByRef bytSource() As Byte, ByVal kcFrom As KCCode, ByVal kcTo As KCCode) As Byte()
    '*****************************************************************************************
    '���w�肵�������R�[�h�ɕϊ�����
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource()             �ϊ����̕����񂪓�����Byte�^�z��
    '           kcFrom                  �ϊ����̕����R�[�h(�񋓌^KCCode)
    '           kcTo                    �ϊ���̕����R�[�h(�񋓌^KCCode)
    '[�߂�l]   �ϊ����ꂽ�����񂪊i�[���ꂽByte�^�z��
    '-----------------------------------------------------------------------------------------
    
    '[�ϊ���������]
    '    JIS        ����  Shift-JIS   ���� EUC
    '                       ��
    '                       ��
    '   UTF-16LE    ����   UTF-16N    ���� UTF-16BE
    '                       ��
    '                       ��
    '   UTF-8BOM    ����   UTF-8N
    
    Dim bytResult() As Byte     '�ϊ����ꂽ��������i�[����Byte�^�z��
    Dim i As Long               '�J�E���^
    
    '�����ϊ����ƕϊ��悪���������R�[�h��������ϊ����Ȃ�
    If kcFrom = kcTo Then
        KCConvert = bytSource
        Exit Function
    End If
    
    '�ϊ��������񂪋�̂Ƃ��͒E�o
    If UBound(bytSource) < 0 Then
        KCConvert = bytSource
        Exit Function
    End If
    
    '�ϊ��悪KC_UNKNOWN��������ϊ����Ȃ�
    If kcTo = KC_UNKNOWN Then
        KCConvert = bytSource
        Exit Function
    End If
    
    '�ϊ�����KC_UNKNOWN�������玩���F��
    If kcFrom = KC_UNKNOWN Then
        kcFrom = KCDetectCode(bytSource)
        '�����F���ł��Ȃ���ΒE�o
        If kcFrom = KC_UNKNOWN Then
            Exit Function
        End If
    End If
    
    '�������ϊ����Ă���
    Select Case kcFrom
    Case KC_SHIFTJIS
        '�ϊ���ɏ]���ď���
        Select Case kcTo
        Case KC_JIS
            bytResult = KCConvertShiftJISIntoJIS(bytSource)
        Case KC_EUC
            bytResult = KCConvertShiftJISIntoEUC(bytSource)
        Case Else
            'UTF-16N�ɕϊ����Ă��炻�̐��
            bytResult = StrConv(bytSource, vbUnicode)
            bytResult = KCConvert(bytResult, KC_UNICODESTRING, kcTo)
        End Select
    
    Case KC_JIS
        'ShiftJIS�ɕϊ����Ă��炻�̐��
        bytResult = KCConvertJISIntoShiftJIS(bytSource)
        bytResult = KCConvert(bytResult, KC_SHIFTJIS, kcTo)
    
    Case KC_EUC
        'ShiftJIS�ɕϊ����Ă��炻�̐��
        bytResult = KCConvertEUCIntoShiftJIS(bytSource)
        bytResult = KCConvert(bytResult, KC_SHIFTJIS, kcTo)
    
    Case KC_UTF8N
        '�ϊ���ɏ]���ď���
        Select Case kcTo
        Case KC_UTF8BOM
            'BOM��t��
            ReDim bytResult(UBound(bytSource) + 3) As Byte
            bytResult(0) = bytUTF8BOM1
            bytResult(1) = bytUTF8BOM2
            bytResult(2) = bytUTF8BOM3
            For i = 3 To UBound(bytResult)
                bytResult(i) = bytSource(i - 3)
            Next
        Case Else
            'UTF-16N�ɕϊ����Ă��炻�̐��
            bytResult = KCConvertUTF8IntoUnicode(bytSource)
            bytResult = KCConvert(bytResult, KC_UNICODESTRING, kcTo)
        End Select
        
    Case KC_UTF8BOM
        'BOM���폜����UTF-8N�ɂ��Ă��炻�̐��
        If bytSource(0) = &HEF And bytSource(1) = &HBB And bytSource(2) = &HBF Then
            ReDim bytResult(UBound(bytSource) - 3) As Byte
            For i = 0 To UBound(bytResult)
                bytResult(i) = bytSource(i + 3)
            Next
            bytResult = KCConvert(bytResult, KC_UTF8N, kcTo)
        Else
            Err.Raise KC_ERROR_INVALID
        End If
    
    Case KC_UNICODESTRING
        '�ϊ���ɏ]���ď���
        Select Case kcTo
        Case KC_UTF8BOM
            bytResult = KCConvertUnicodeIntoUTF8(bytSource)
            bytResult = KCConvert(bytResult, KC_UTF8N, KC_UTF8BOM)
        Case KC_UTF8N
            bytResult = KCConvertUnicodeIntoUTF8(bytSource)
        Case KC_UTF16LE
            'BOM��t��
            ReDim bytResult(UBound(bytSource) + 2) As Byte
            bytResult(0) = bytUTF16BOM2
            bytResult(1) = bytUTF16BOM1
            For i = 2 To UBound(bytResult)
                bytResult(i) = bytSource(i - 2)
            Next
        Case KC_UTF16BE
            'BOM��t�����ăo�C�g�I�[�_���t��
            ReDim bytResult(UBound(bytSource) + 2) As Byte
            bytResult(0) = bytUTF16BOM1
            bytResult(1) = bytUTF16BOM2
            For i = 2 To UBound(bytResult) Step 2
                bytResult(i) = bytSource(i - 1)
                bytResult(i + 1) = bytSource(i - 2)
            Next
        Case Else
            'Shift-JIS�ɕϊ����Ă��炻�̐��
            bytResult = StrConv(bytSource, vbFromUnicode)
            bytResult = KCConvert(bytResult, KC_SHIFTJIS, kcTo)
        End Select
        
    Case KC_UTF16LE
        'BOM���폜����UTF-16N�ɂ��Ă��炻�̐��
        If bytSource(0) = bytUTF16BOM2 And bytSource(1) = bytUTF16BOM1 Then
            ReDim bytResult(UBound(bytSource) - 2) As Byte
            For i = 0 To UBound(bytResult)
                bytResult(i) = bytSource(i + 2)
            Next
            bytResult = KCConvert(bytResult, KC_UNICODESTRING, kcTo)
        Else
            Err.Raise KC_ERROR_INVALID
        End If
    
    Case KC_UTF16BE
        'BOM���폜���ăo�C�g�I�[�_���t�ɂ��CUTF-16N�ɂ��Ă��炻�̐��
        If bytSource(0) = bytUTF16BOM1 And bytSource(1) = bytUTF16BOM2 Then
            ReDim bytResult(UBound(bytSource) - 2) As Byte
            For i = 0 To UBound(bytResult) Step 2
                bytResult(i) = bytSource(i + 3)
                bytResult(i + 1) = bytSource(i + 2)
            Next
            bytResult = KCConvert(bytResult, KC_UNICODESTRING, kcTo)
        Else
            Err.Raise KC_ERROR_INVALID
        End If
    
    Case Else
        '�Ăяo�������s��
        Err.Raise KC_ERROR_INVALID
    End Select
    
    KCConvert = bytResult
    
End Function

Private Function KCConvertShiftJISIntoJIS(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '��Shift-JIS��JIS�ɕϊ�����(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource()             �ϊ�����Shift-JIS�����񂪓�����Byte�^�z��
    '[�߂�l]   �ϊ����ꂽ�����񂪊i�[���ꂽByte�^�z��
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '�ϊ����ꂽ��������i�[����Byte�^�z��
    Dim emMode As EncodeMode    '���������[�}�����J�^�J�i���̃t���O
    Dim i As Long               '�J�E���^
    Dim j As Long               '�J�E���^
    
    ReDim bytResult(UBound(bytSource) * 5) As Byte     '�傫�߂ɗp�ӂ��Ă���
    
    '�ϐ��̏�����
    i = 0
    j = 0
    emMode = mRome
    
    '�ϊ�����
    Do While i <= UBound(bytSource())
        If (bytSource(i) >= &H80& And bytSource(i) <= &HA0) Or (bytSource(i) >= &HE0) Then
            '�擪�r�b�g�������Ă��āA���p�J�i�̗̈�ɂȂ��Ƃ�
            '�S�p����
            If Not emMode = mKanji Then
                '����܂őS�p�����łȂ������Ƃ���JIS����IN������
                bytResult(j) = bytJISESC
                bytResult(j + 1) = bytKANJI1
                bytResult(j + 2) = bytKANJI2NEW
                
                '�J�E���^��i�߂�
                j = j + 3
            End If
            
            '��U���̂܂ܑ��
            bytResult(j) = bytSource(i)
            bytResult(j + 1) = bytSource(i + 1)
            
            '�ϊ�����
            If bytResult(j) >= &HE0& Then
                bytResult(j) = bytResult(j) - &H40&
            End If
            If bytResult(j + 1) >= &H80& Then
                bytResult(j + 1) = bytResult(j + 1) - 1
            End If
            If bytResult(j + 1) >= &H9E& Then
                bytResult(j) = (bytResult(j) - &H70) * 2
                bytResult(j + 1) = bytResult(j + 1) - &H7D
            Else
                bytResult(j) = (bytResult(j) - &H70) * 2 - 1
                bytResult(j + 1) = bytResult(j + 1) - &H1F
                
            End If
            
            '�����ł���t���O�𗧂Ă�
            emMode = mKanji
            
            '�J�E���^��i�߂�
            i = i + 2
            j = j + 2
        ElseIf bytSource(i) >= &HA1 And bytSource(i) <= &HDF Then
            '���p�J�^�J�i
            
            If Not emMode = mKata Then
                'JIS�J�^�J�iIN������
                bytResult(j) = bytJISESC
                bytResult(j + 1) = bytKATA1
                bytResult(j + 2) = bytKATA2
                
                '�J�E���^��i�߂�
                j = j + 3
            End If
            
            '�g�b�v�r�b�g�����낷
            bytResult(j) = bytSource(i) - &H80
            
            '�J�^�J�i�ł���t���O�𗧂Ă�
            emMode = mKata
            
            i = i + 1
            j = j + 1
            
        Else
            'ANSI����
            
            If Not emMode = mRome Then
                'JIS���[�}��IN������
                bytResult(j) = bytJISESC
                bytResult(j + 1) = bytROME1
                bytResult(j + 2) = bytROME2
                
                '�J�E���^��i�߂�
                j = j + 3
            End If
            
            '���̂܂�
            bytResult(j) = bytSource(i)
            
            '���[�}���ł���t���O�𗧂Ă�
            emMode = mRome
            
            i = i + 1
            j = j + 1
        End If
    Loop
    
    '�I���O�ɂ͕K�����[�}���ɖ߂�
    If Not emMode = mRome Then
        'JIS���[�}��IN������
        bytResult(j) = bytJISESC
        bytResult(j + 1) = bytROME1
        bytResult(j + 2) = bytROME2
        
        '�J�E���^��i�߂�
        j = j + 3
    End If
    
    'bytResult����K�v�ȕ������������o��
    ReDim Preserve bytResult(j - 1) As Byte
    
    '�߂�l�ɑ��
    KCConvertShiftJISIntoJIS = bytResult()
    
End Function

Private Function KCConvertJISIntoShiftJIS(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '��JIS��Shift-JIS�ɕϊ�����(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource()             �ϊ�����JIS�����񂪓�����Byte�^�z��
    '[�߂�l]   �ϊ����ꂽ�����񂪊i�[���ꂽByte�^�z��
    '-----------------------------------------------------------------------------------------
    'ESC ( J�Ŕ��p�J�^�J�i�ɓ�����̂ɑΉ����邽�߁A�t���O�͔��p�������S�p�������𕪂��邾���ɂ��Ă��܂��B
    
    Dim bytResult() As Byte     '�ϊ����ꂽ��������i�[����Byte�^�z��
    Dim emMode As EncodeMode    '���������[�}�����J�^�J�i���̃t���O
    Dim i As Long               '�J�E���^
    Dim j As Long               '�J�E���^
    
    ReDim bytResult(UBound(bytSource) * 5) As Byte     '�傫�߂ɗp�ӂ��Ă���
    
    '�ϐ��̏�����
    i = 0
    j = 0
    emMode = mRome
    
    '�ϊ�����
    Do While i <= UBound(bytSource())
        If bytSource(i) = bytJISESC Then
            'ESC�������Ƃ�
            If bytSource(i + 1) = bytKANJI1 And (bytSource(i + 2) = bytKANJI2OLD Or bytSource(i + 2) = bytKANJI2NEW) Then
                '����IN�ł������Ƃ�
                
                '�S�p�����ł���t���O�𗧂Ă�
                emMode = mKanji
                
                '�J�E���^��i�߂�
                i = i + 3
                
            ElseIf bytSource(i + 1) = bytROME1 And bytSource(i + 2) = bytROME2 Then
                '���[�}��IN�ł������Ƃ�
                
                '���[�}���ł���t���O�𗧂Ă�
                emMode = mRome
                
                '�J�E���^��i�߂�
                i = i + 3
                
            ElseIf bytSource(i + 1) = bytKATA1 And bytSource(i + 2) = bytKATA2 Then
                '�J�^�J�iIN�ł������Ƃ�
                
                '�J�^�J�i�ł���t���O�𗧂Ă�
                emMode = mKata
                
                '�J�E���^��i�߂�
                i = i + 3
                
            Else
                '�������ăJ�E���^��i�߂�(��������Ȃ��ƕs���ȃt�@�C���̎��Ɏ~�܂��Ă��܂�)
                i = i + 1
                
            End If
            
        Else
            'ESC�łȂ������Ƃ�
            Select Case emMode
            Case mKanji
                '�S�p�����ł���Ƃ�
                
                '��U���̂܂ܑ��
                bytResult(j) = bytSource(i)
                bytResult(j + 1) = bytSource(i + 1)
                
                '�ϊ�
                If (bytResult(j) Mod 2) = 1 Then
                    bytResult(j) = (bytResult(j) + 1) \ 2 + &H70
                    bytResult(j + 1) = bytResult(j + 1) + &H1F
                Else
                    bytResult(j) = bytResult(j) \ 2 + &H70
                    bytResult(j + 1) = bytResult(j + 1) + &H7D
                End If
                If bytResult(j) >= &HA0 Then
                    bytResult(j) = bytResult(j) + &H40
                End If
                If bytResult(j + 1) >= &H7F Then
                    bytResult(j + 1) = bytResult(j + 1) + &H1
                End If
                
                '�J�E���^��i�߂�
                i = i + 2
                j = j + 2
                
            Case mRome
                '���[�}���ł���Ƃ�
                
                '���̂܂�
                bytResult(j) = bytSource(i)
                
                '�J�E���^��i�߂�
                i = i + 1
                j = j + 1
            
            Case mKata
                '���p�J�i�ł���Ƃ�
                
                '�g�b�v�r�b�g�𗧂Ă�
                bytResult(j) = bytSource(i) Or &H80
                
                '�J�E���^��i�߂�
                i = i + 1
                j = j + 1
                
            End Select
        End If
        
        
    Loop
    
    'bytResult����K�v�ȕ������������o��
    ReDim Preserve bytResult(j - 1) As Byte
    
    '�߂�l�ɑ��
    KCConvertJISIntoShiftJIS = bytResult()
    
End Function

Private Function KCConvertEUCIntoShiftJIS(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '��EUC��ShiftJIS�ɕϊ�����(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource()             �ϊ�����EUC�����񂪓�����Byte�^�z��
    '[�߂�l]   �ϊ����ꂽ�����񂪊i�[���ꂽByte�^�z��
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '�ϊ����ꂽ��������i�[����Byte�^�z��
    Dim i As Long               '�J�E���^(�ϊ���)
    Dim j As Long               '�J�E���^(�ϊ���)
    
    ReDim bytResult(UBound(bytSource) * 2) As Byte     '�傫�߂ɗp�ӂ��Ă���
    
    '�ϐ��̏�����
    i = 0
    j = 0
    
    '�ϊ�����
    Do While i <= UBound(bytSource())
        
        If bytSource(i) = bytSS2 Then
            '���p�J�i
            
            '��1�o�C�g���΂�
            bytResult(j) = bytSource(i + 1)
            
            '�J�E���^��i�߂�
            i = i + 2
            j = j + 1
            
        ElseIf bytSource(i) = bytSS3 Then
            'JIS�⏕����(ShiftJIS�ł͕\���s��)
            
            bytResult(j) = &H3F     '�Ƃ肠����"?"�ɒu��������
            
            '�J�E���^��i�߂�
            i = i + 3
            j = j + 1
            
        ElseIf bytSource(i) >= &H80& Then
            '�ʏ�̑S�p����
            
            '�ϊ�����
            bytResult(j) = bytSource(i) - &H80
            bytResult(j + 1) = bytSource(i + 1) - &H80
            If (bytResult(j) Mod 2) = 1 Then
                bytResult(j) = (bytResult(j) + 1) \ 2 + &H70
                bytResult(j + 1) = bytResult(j + 1) + &H1F
            Else
                bytResult(j) = bytResult(j) \ 2 + &H70
                bytResult(j + 1) = bytResult(j + 1) + &H7D
            End If
            If bytResult(j) >= &HA0 Then
                bytResult(j) = bytResult(j) + &H40
            End If
            If bytResult(j + 1) >= &H7F Then
                bytResult(j + 1) = bytResult(j + 1) + &H1
            End If
            
            
            '�J�E���^��i�߂�
            i = i + 2
            j = j + 2
        Else
            'ASCII�Ɣ��f
            
            '���̂܂�
            bytResult(j) = bytSource(i)
            
            i = i + 1
            j = j + 1
        End If
    Loop
    
    'bytResult����K�v�ȕ������������o��
    ReDim Preserve bytResult(j - 1) As Byte
    
    '�߂�l�ɑ��
    KCConvertEUCIntoShiftJIS = bytResult()
    
End Function

Private Function KCConvertShiftJISIntoEUC(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '��ShiftJIS��EUC�ɕϊ�����(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource()             �ϊ�����JIS�����񂪓�����Byte�^�z��
    '[�߂�l]   �ϊ����ꂽ������̓�����Byte�^�z��
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '�ϊ����ꂽ��������i�[����Byte�^�z��
    Dim i As Long               '�J�E���^(�ϊ���)
    Dim j As Long               '�J�E���^(�ϊ���)
    
    ReDim bytResult(UBound(bytSource) * 2) As Byte     '�傫�߂ɗp�ӂ��Ă���
    
    '�ϐ��̏�����
    i = 0
    j = 0
    
    '�ϊ�����
    Do While i <= UBound(bytSource())
        If (bytSource(i) >= &H80& And bytSource(i) <= &HA0) Or (bytSource(i) >= &HE0) Then
            '�擪�r�b�g�������Ă��āA���p�J�i�̗̈�ɂȂ��Ƃ�
            '�S�p����
            
            '��U���̂܂ܑ��
            bytResult(j) = bytSource(i)
            bytResult(j + 1) = bytSource(i + 1)
            
            '�ϊ�����
            If bytResult(j) >= &HE0& Then
                bytResult(j) = bytResult(j) - &H40&
            End If
            If bytResult(j + 1) >= &H80& Then
                bytResult(j + 1) = bytResult(j + 1) - 1
            End If
            If bytResult(j + 1) >= &H9E& Then
                bytResult(j) = (bytResult(j) - &H70) * 2
                bytResult(j + 1) = bytResult(j + 1) - &H7D
            Else
                bytResult(j) = (bytResult(j) - &H70) * 2 - 1
                bytResult(j + 1) = bytResult(j + 1) - &H1F
            End If
            bytResult(j) = bytResult(j) + &H80
            bytResult(j + 1) = bytResult(j + 1) + &H80
            
            '�J�E���^��i�߂�
            i = i + 2
            j = j + 2
        
        ElseIf bytSource(i) >= &HA1 And bytSource(i) <= &HDF Then
            '���p�J�^�J�i
            
            'SS2��}��
            bytResult(j) = bytSS2
            
            bytResult(j + 1) = bytSource(i)
            
            i = i + 1
            j = j + 2
            
        Else
            'ASCII����
            
            '���̂܂�
            bytResult(j) = bytSource(i)
            
            i = i + 1
            j = j + 1
        End If
    Loop
    
    'bytResult����K�v�ȕ������������o��
    ReDim Preserve bytResult(j - 1) As Byte
    
    '�߂�l�ɑ��
    KCConvertShiftJISIntoEUC = bytResult()
    
End Function

Private Function KCConvertUnicodeIntoUTF8(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '��Unicode�������UTF-8�ɕϊ�����(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource               �ϊ�����UTF-16�����񂪓�����Byte�^�z��
    '[�߂�l]   �ϊ����ꂽ�����񂪊i�[���ꂽByte�^�z��
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '�ϊ����ꂽ��������i�[����Byte�^�z��
    Dim i As Long               '�J�E���^
    Dim j As Long               '�J�E���^
    Dim lngRet As Long
    
    ReDim bytResult(UBound(bytSource) * 5) As Byte     '�傫�߂ɗp�ӂ��Ă���
    
    '�ϐ��̏�����
    i = 0
    j = 0
    
    For i = 0 To UBound(bytSource) Step 2
        '1������Long�^�Ɋi�[
        lngRet = CLng(bytSource(i)) + CLng(bytSource(i + 1)) * 256
        
        Select Case lngRet
        Case 0 To &H7F&
            'ASCII�����̂Ƃ�
            bytResult(j) = CByte(lngRet)    '���̂܂�
            
            '�J�E���^��i�߂�
            j = j + 1
            
        Case &H80& To &H7FF&
            '2�o�C�g�����̂Ƃ�
            
            '��1�o�C�g
            bytResult(j) = CByte(&HC0 Or ((lngRet And &H7C0&) \ 64))
            '��2�o�C�g
            bytResult(j + 1) = CByte(&H80 Or (lngRet And &H3F))
            
            '�J�E���^��i�߂�
            j = j + 2
            
        Case &H800& To &HFFFF&
            '3�o�C�g�����̂Ƃ�
            
            '��1�o�C�g
            bytResult(j) = CByte(&HE0 Or ((lngRet And &HF000&) \ 4096))
            '��2�o�C�g
            bytResult(j + 1) = CByte(&H80 Or ((lngRet And &HFC0&) \ 64))
            '��3�o�C�g
            bytResult(j + 2) = CByte(&H80 Or (lngRet And &H3F))
            
            '�J�E���^��i�߂�
            j = j + 3
            
        End Select
    Next i
    
    'bytResult����K�v�ȕ������������o��
    ReDim Preserve bytResult(j - 1) As Byte
    
    '�߂�l�ɑ��
    KCConvertUnicodeIntoUTF8 = bytResult()
    
End Function

Private Function KCConvertUTF8IntoUnicode(ByRef bytSource() As Byte) As String
    '*****************************************************************************************
    '��UTF-8��Unicode������ɕϊ�����(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource               �ϊ�����UTF-8�����񂪓�����Byte�^�z��
    '[�߂�l]   �ϊ����ꂽ�����񂪊i�[���ꂽString
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '�ϊ����ꂽ��������i�[����Byte�^�z��
    Dim i As Long               '�J�E���^
    Dim j As Long               '�J�E���^
    Dim lngRet As Long
    
    ReDim bytResult(UBound(bytSource) * 3) As Byte      '�傫�߂ɗp�ӂ��Ă���
    
    '�ϐ��̏�����
    i = 0
    j = 0
    
    '�ϊ�����
    Do While i <= UBound(bytSource())
        Select Case bytSource(i)
        Case 0 To &H7F
            'ASCII�����������Ƃ�
            
            bytResult(j) = bytSource(i)
            bytResult(j + 1) = 0
            
            '�J�E���^��i�߂�
            i = i + 1
            j = j + 2
            
        Case &HC2 To &HDF
            '2�o�C�g����
            
            bytResult(j) = (bytSource(i) And &H3&) * 64 Or (bytSource(i + 1) And &H3F)
            bytResult(j + 1) = (bytSource(i) And &H1C) \ 4
            
            '�J�E���^��i�߂�
            i = i + 2
            j = j + 2
            
        Case &HE0 To &HEF
            '3�o�C�g����
        
            bytResult(j) = (bytSource(i + 1) And &H3) * 64 Or (bytSource(i + 2) And &H3F)
            bytResult(j + 1) = (bytSource(i) And &HF) * 16 Or (bytSource(i + 1) And &H3C) \ 4
            
            '�J�E���^��i�߂�
            i = i + 3
            j = j + 2
            
        Case &HF0 To &HF7
            '4�o�C�g����(VB��Unicode(UCS-2)�ł͕\���s��)
            
            bytResult(j) = (bytSource(i + 2) And &H3) * 64 Or (bytSource(i + 3) And &H3F)
            bytResult(j + 1) = (bytSource(i + 1) And &HF) * 16 Or (bytSource(i + 2) And &H3C) \ 4
            
            '�J�E���^��i�߂�
            i = i + 4
            j = j + 2
            
        Case &HF8 To &HFB
            '5�o�C�g����(VB��Unicode(UCS-2)�ł͕\���s��)
            
            bytResult(j) = (bytSource(i + 3) And &H3) * 64 Or (bytSource(i + 4) And &H3F)
            bytResult(j + 1) = (bytSource(i + 2) And &HF) * 16 Or (bytSource(i + 3) And &H3C) \ 4
            
            '�J�E���^��i�߂�
            i = i + 5
            j = j + 2
            
        Case &HFC To &HFD
            '6�o�C�g����(VB��Unicode(UCS-2)�ł͕\���s��)
            
            bytResult(j) = (bytSource(i + 4) And &H3) * 64 Or (bytSource(i + 5) And &H3F)
            bytResult(j + 1) = (bytSource(i + 3) And &HF) * 16 Or (bytSource(i + 4) And &H3C) \ 4
            
            '�J�E���^��i�߂�
            i = i + 6
            j = j + 2
            
        Case Else
            '�s���ȕ�����
            Err.Raise KC_ERROR_INVALID
            
        End Select
    Loop
    
    'bytResult����K�v�ȕ������������o��
    ReDim Preserve bytResult(j - 1) As Byte
    
    '�߂�l�ɑ��
    KCConvertUTF8IntoUnicode = bytResult()
    
End Function

Public Function KCDetectCode(ByRef bytSource() As Byte) As KCCode
    '*****************************************************************************************
    '�������R�[�h�𔻕ʂ���
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource               ���f�����R�[�h�̕����񂪊i�[���ꂽByte�^�z��
    '[�߂�l]   ���f���ꂽ�R�[�h�����߂��񋓌^KCCode(���W���[���擪���Q�Ƃ��Ă�������)
    '-----------------------------------------------------------------------------------------
    Dim i As Long                       '�J�E���^
    Dim lngCountStart As Long           '�J�E���^�̍ŏ��̓Y����
    Dim bolIsAscii As Boolean           'ASCII���ǂ����̃t���O
    
    '�ϐ�������
    lngCountStart = 0
    
    'BOM���`�F�b�N
    If UBound(bytSource()) >= 3 Then
        If bytSource(0) = bytUTF8BOM1 And bytSource(1) = bytUTF8BOM2 And bytSource(2) = bytUTF8BOM3 Then
            'UTF-8BOM�H
            If KCIsValidUTF8(bytSource, True) Then
                KCDetectCode = KC_UTF8BOM
                Exit Function
            End If
        End If
    End If
    If UBound(bytSource()) >= 2 Then
        If bytSource(0) = bytUTF16BOM2 And bytSource(1) = bytUTF16BOM1 Then
            'UTF-16LE�H
            If UBound(bytSource()) Mod 2 = 1 Then   '�o�C�g������łȂ����Ƃ������m�F
                KCDetectCode = KC_UTF16LE
                Exit Function
            End If
        ElseIf bytSource(0) = bytUTF16BOM1 And bytSource(1) = bytUTF16BOM2 Then
            'UTF-16BE�H
            If UBound(bytSource()) Mod 2 = 1 Then   '�o�C�g������łȂ����Ƃ������m�F
                KCDetectCode = KC_UTF16BE
                Exit Function
            End If
        End If
    End If
    
    '�܂�ASCII��JIS���𔻕�
    bolIsAscii = True
    For i = 0 To UBound(bytSource())
        If Not KCInRange(bytSource(i), &H0, &H7F) Then
            'ASCII�ł͂Ȃ�(�������JIS�ł��Ȃ�)
            bolIsAscii = False
            Exit For
        ElseIf bytSource(i) = bytJISESC Then
            'ESC���o�Ă��Ă���̂�JIS�̉\���A��
            If KCIsValidJIS(bytSource) Then
                KCDetectCode = KC_JIS
                Exit Function
            End If
        End If
    Next i
    If bolIsAscii Then
        'ASCII������(�}���`�o�C�g�������g���Ă��Ȃ�)�Ƃ���ShiftJIS�Ƃ��ĕԂ�
        KCDetectCode = KC_SHIFTJIS
        Exit Function
    End If
    
    'ShiftJIS������
    If KCIsValidShiftJIS(bytSource()) Then
        KCDetectCode = KC_SHIFTJIS
        Exit Function
    End If
    
    'EUC������
    If KCIsValidEUC(bytSource()) Then
        KCDetectCode = KC_EUC
        Exit Function
    End If
    
    'UTF-8N���𔻒�
    If KCIsValidUTF8(bytSource(), False) Then
        KCDetectCode = KC_UTF8N
        Exit Function
    End If
    
    '�����܂ł���Ĕ��ʂł��Ȃ����UCS2���H
    If UBound(bytSource()) Mod 2 = 1 Then   '�o�C�g������łȂ����Ƃ������m�F
        KCDetectCode = KC_UNICODESTRING
        Exit Function
    End If
    
    '���ʕs�\
    KCDetectCode = KC_UNKNOWN
    
End Function

Private Function KCIsValidJIS(ByRef bytSource() As Byte) As Boolean
    '*****************************************************************************************
    '��������JIS�����񂩂ǂ����𔻕ʂ���(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource               ���f����镶���񂪊i�[���ꂽByte�^�z��
    '[�߂�l]   �������Ȃ��True,�s���ȕ�����Ȃ��False
    '-----------------------------------------------------------------------------------------
    
    Dim i As Long               '�J�E���^
    Dim emMode As EncodeMode    '���������[�}�����J�^�J�i���̃t���O
    
    emMode = mRome
    KCIsValidJIS = False
    
    On Error GoTo Err_Handler
    
    For i = 0 To UBound(bytSource)
        If bytSource(i) = bytJISESC Then
            'ESC�������Ƃ�
            If bytSource(i + 1) = bytKANJI1 And (bytSource(i + 2) = bytKANJI2OLD Or bytSource(i + 2) = bytKANJI2NEW) Then
                '����IN�ł������Ƃ�
                emMode = mKanji
                i = i + 2
            ElseIf bytSource(i + 1) = bytROME1 And bytSource(i + 2) = bytROME2 Then
                '���[�}��IN�ł������Ƃ�
                emMode = mRome
                i = i + 2
            ElseIf bytSource(i + 1) = bytKATA1 And bytSource(i + 2) = bytKATA2 Then
                '�J�^�J�iIN�ł������Ƃ�
                emMode = mKata
                i = i + 2
            Else
                'ESC�������������Ƃ���JIS�ł���Ƃ͔��f���Ȃ�
                Exit Function
            End If
        Else
            Select Case emMode
            Case mRome
                If Not KCInRange(bytSource(i), &H0, &H7F) Then
                    '�s��
                    Exit Function
                End If
            Case Else
                If Not KCInRange(bytSource(i), &H20, &H7F) Then
                    '�s��
                    Exit Function
                End If
            End Select
        End If
    Next
    
    '����
    KCIsValidJIS = True
    
Err_Handler:
    '�G���[����������s��
    
End Function

Private Function KCIsValidShiftJIS(ByRef bytSource() As Byte) As Boolean
    '*****************************************************************************************
    '��������ShiftJIS�����񂩂ǂ����𔻕ʂ���(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource               ���f����镶���񂪊i�[���ꂽByte�^�z��
    '[�߂�l]   �������Ȃ��True,�s���ȕ�����Ȃ��False
    '-----------------------------------------------------------------------------------------
    Dim i As Long               '�J�E���^
    
    KCIsValidShiftJIS = False
    
    On Error GoTo Err_Handler
    
    For i = 0 To UBound(bytSource)
        If KCInRange(bytSource(i), &H81, &H9F) Or KCInRange(bytSource(i), &HE0, &HFC) Then
            '�S�p����
            If Not (KCInRange(bytSource(i + 1), &H40, &H7E) Or KCInRange(bytSource(i + 1), &H80, &HFC)) Then
                '�s��
                Exit Function
            End If
            i = i + 1
        ElseIf Not (KCInRange(bytSource(i), &HA1, &HDF) Or KCInRange(bytSource(i), &H0, &H7F)) Then
            'ASCII�┼�p�J�i�łȂ���Εs��
            Exit Function
        End If
    Next i
    
    '������������ł���
    KCIsValidShiftJIS = True
    
Err_Handler:
    '�G���[����������s��
    
End Function

Private Function KCIsValidEUC(ByRef bytSource() As Byte) As Boolean
    '*****************************************************************************************
    '��������EUC�����񂩂ǂ����𔻕ʂ���(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource               ���f����镶���񂪊i�[���ꂽByte�^�z��
    '[�߂�l]   �������Ȃ��True,�s���ȕ�����Ȃ��False
    '-----------------------------------------------------------------------------------------
    Dim i As Long               '�J�E���^
    
    KCIsValidEUC = False
    
    On Error GoTo Err_Handler
    
    For i = 0 To UBound(bytSource)
        If bytSource(i) = bytSS2 Then
            If Not KCInRange(bytSource(i + 1), &HA1, &HDF) Then
                '�s��
                Exit Function
            End If
            i = i + 1
        ElseIf bytSource(i) = bytSS3 Then
            If Not (KCInRange(bytSource(i + 1), &HA1, &HFE) And KCInRange(bytSource(i + 2), &HA1, &HFE)) Then
                '�s��
                Exit Function
            End If
            i = i + 2
        ElseIf KCInRange(bytSource(i), &HA1, &HFE) Then
            If Not KCInRange(bytSource(i + 1), &HA1, &HFE) Then
                '�s��
                Exit Function
            End If
            i = i + 1
        ElseIf Not KCInRange(bytSource(i), &H0, &H7F) Then
            '�s��
            Exit Function
        End If
    Next i
    
    '������������ł���
    KCIsValidEUC = True
    
Err_Handler:
    '�G���[����������s��
    
End Function

Private Function KCIsValidUTF8(ByRef bytSource() As Byte, ByVal bolBOM As Boolean) As Boolean
    '*****************************************************************************************
    '��������UTF-8�����񂩂ǂ����𔻕ʂ���(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   bytSource               ���f����镶���񂪊i�[���ꂽByte�^�z��
    '           bolBOM                  BOM����Ȃ��True,�Ȃ��Ȃ��False
    '[�߂�l]   �������Ȃ��True,�s���ȕ�����Ȃ��False
    '-----------------------------------------------------------------------------------------
    
    Dim i As Long               '�J�E���^
    
    KCIsValidUTF8 = False
    
    On Error GoTo Err_Handler
    
    For i = 0 To UBound(bytSource)
        If KCInRange(bytSource(i), &H0, &H7F) Then
        ElseIf KCInRange(bytSource(i), &HC2, &HDF) And _
                KCInRange(bytSource(i + 1), &H80, &HBF) Then
            i = i + 1
        ElseIf KCInRange(bytSource(i), &HE0, &HE0) And _
                KCInRange(bytSource(i + 1), &HA0, &HBF) And _
                KCInRange(bytSource(i + 2), &H80, &HBF) Then
            i = i + 2
        ElseIf KCInRange(bytSource(i), &HE1, &HEC) And _
                KCInRange(bytSource(i + 1), &H80, &HBF) And _
                KCInRange(bytSource(i + 2), &H80, &HBF) Then
            i = i + 2
        ElseIf KCInRange(bytSource(i), &HED, &HED) And _
                KCInRange(bytSource(i + 1), &H80, &H9F) And _
                KCInRange(bytSource(i + 2), &H80, &HBF) Then
            i = i + 2
        ElseIf KCInRange(bytSource(i), &HEE, &HEF) And _
                KCInRange(bytSource(i + 1), &H80, &HBF) And _
                KCInRange(bytSource(i + 2), &H80, &HBF) Then
            i = i + 2
        ElseIf KCInRange(bytSource(i), &HF0, &HF0) And _
                KCInRange(bytSource(i + 1), &H90, &HBF) And _
                KCInRange(bytSource(i + 2), &H80, &HBF) And _
                KCInRange(bytSource(i + 3), &H80, &HBF) Then
            i = i + 3
        ElseIf KCInRange(bytSource(i), &HF1, &HF3) And _
                KCInRange(bytSource(i + 1), &H80, &HBF) And _
                KCInRange(bytSource(i + 2), &H80, &HBF) And _
                KCInRange(bytSource(i + 3), &H80, &HBF) Then
            i = i + 3
        ElseIf KCInRange(bytSource(i), &HF4, &HF4) And _
                KCInRange(bytSource(i + 1), &H80, &H8F) And _
                KCInRange(bytSource(i + 2), &H80, &HBF) And _
                KCInRange(bytSource(i + 3), &H80, &HBF) Then
            i = i + 3
        Else
            '�s���ȕ�����
            Exit Function
        End If
        
    Next i
    
    '������������ł���
    KCIsValidUTF8 = True
    
Err_Handler:
    '�G���[����������s��
    
End Function

Private Function KCInRange(ByVal lngTarget As Long, ByVal lngLowerBound As Long, ByVal lngUpperBound As Long) As Boolean
    '*****************************************************************************************
    '���͈͓��ɓ����Ă��邩���肷��(�����֐�)
    '-----------------------------------------------------------------------------------------
    '[ ���� ]   lngTarget               ���肳��鐔�l
    '           lngLowerBound           ����
    '           lngUpperBound           ���
    '[�߂�l]   �͈͓��Ȃ��True,�͈͊O�Ȃ��False
    '-----------------------------------------------------------------------------------------

    KCInRange = (lngTarget >= lngLowerBound) And (lngTarget <= lngUpperBound)
End Function
