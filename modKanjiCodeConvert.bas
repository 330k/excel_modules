Attribute VB_Name = "modKanjiCodeConvert"
'*********************************************************************************************
'*                            漢字コード変換モジュール For VB6/VBA                           *
'*********************************************************************************************
'
'[機能/制約]
'このモジュールはShift-JIS/EUC/JIS/Unicode(UTF-8,UTF-16)の相互変換を行います。
'おまけとして文字コード識別関数(β)も用意しました。
'一通りテストはしましたが変換された内容が正しいという保証はできません。
'
'[使い方]
'変換元データをバイト型の配列かString型変数に読み出しておいて
'    KCConvert( バイト配列, 変換元文字コード, 変換先文字コード )
'として呼び出せばOKです．
'文字コードの判定を行いたいときは
'    KCDetectCode( バイト配列 )
'とします．
'なおエラーが発生したときは，エラー番号に定数KC_ERROR_INVALIDをセットします．
'関数を呼び出す前にエラートラップをしてください。
'
'[著作権]
'このモジュールの著作権は保持しますが改変、再配布、ご自由にお使いください。
'このモジュールを使用してあなたが作ったソフトの著作権や、このモジュールを改変したものの
'著作権はあなたにあります。再配布や転載のときはメールで一言連絡いただけるとうれしいです。
'
'[免責事項]
'このモジュールの使用によるいかなる損害に対しても責任を負いません。
'
'[その他]
'変換アルゴリズムに間違い等がありましたら、作者までお知らせください。
'
'[更新履歴]
'Version 1.00   2003/10/08  一応完成
'        1.20   2003/11/01  漢字OUTという間違った概念を使っていたので訂正
'        1.30   2004/03/05  高速化
'        1.50   2005/03/03  UTF-8とUnicode間の変換関数、UTF-8の正当性チェック関数を実装
'        2.00   2008/07/13  KCConvert関数に一本化
'                           EUCとShift-JIS間の変換を直接行い，JIS⇔EUC間の変換関数を削除
'                           UTF-16BE,UTF-16LEに対応
'                           VBAで動作を確認
'                           EUCのシングルシフト3に対応
'                           判別関数を改良
'
'                                   ==========================================================
'                                    330k http://www.330k.info/
'                                   ==========================================================
Option Explicit

'文字コードをあらわす列挙型
Public Enum KCCode
    KC_UNKNOWN = 0              '不明
    KC_SHIFTJIS = 10            'Shift-JIS
    KC_JIS = 20                 'JIS
    KC_EUC = 30                 'EUC-JP
    KC_UTF8N = 80               'UTF-8N
    KC_UTF8BOM = 81             'UTF-8(BOMつき)
    KC_UNICODESTRING = 160      'VBのString型,UCS2(UTF-16LE,BOMなし,サロゲートペア非対応)
    KC_UTF16LE = 161            'UTF-16LE
    KC_UTF16BE = 162            'UTF-16BE
End Enum

'エラーコード
Public Const KC_ERROR_INVALID As Long = 65000

'JISのエスケープシーケンス
Private Const bytJISESC As Byte = &H1B      'JISのESC(エスケープシーケンス)
Private Const bytKANJI1 As Byte = &H24      'JIS漢字IN1バイト目"$"
Private Const bytKANJI2OLD As Byte = &H40   'JIS漢字IN2バイト目(旧JIS)"@"
Private Const bytKANJI2NEW As Byte = &H42   'JIS漢字IN2バイト目(新JIS)"B"
Private Const bytROME1 As Byte = &H28       'JISローマ字IN1バイト目"("
Private Const bytROME2 As Byte = &H4A       'JISローマ字IN2バイト目"J"
Private Const bytKATA1 As Byte = &H28       'JIS半角カタカナIN1バイト目"("
Private Const bytKATA2 As Byte = &H49       'JIS半角カタカナIN2バイト目"I"

'EUCのエスケープシーケンス
Private Const bytSS2 As Byte = &H8E         'Single Shift 2
Private Const bytSS3 As Byte = &H8F         'Single Shift 3

'UTF-8のBOM
Private Const bytUTF8BOM1 As Byte = &HEF
Private Const bytUTF8BOM2 As Byte = &HBB
Private Const bytUTF8BOM3 As Byte = &HBF

'UTF-16のBOM
Private Const bytUTF16BOM1 As Byte = &HFE
Private Const bytUTF16BOM2 As Byte = &HFF

'変換処理中に内部で使用するフラグ
Private Enum EncodeMode
    mKanji = 1
    mRome = 2
    mKata = 3
End Enum

Public Function KCConvert(ByRef bytSource() As Byte, ByVal kcFrom As KCCode, ByVal kcTo As KCCode) As Byte()
    '*****************************************************************************************
    '■指定した文字コードに変換する
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource()             変換元の文字列が入ったByte型配列
    '           kcFrom                  変換元の文字コード(列挙型KCCode)
    '           kcTo                    変換先の文字コード(列挙型KCCode)
    '[戻り値]   変換された文字列が格納されたByte型配列
    '-----------------------------------------------------------------------------------------
    
    '[変換処理流れ]
    '    JIS        ←→  Shift-JIS   ←→ EUC
    '                       ↑
    '                       ↓
    '   UTF-16LE    ←→   UTF-16N    ←→ UTF-16BE
    '                       ↑
    '                       ↓
    '   UTF-8BOM    ←→   UTF-8N
    
    Dim bytResult() As Byte     '変換された文字列を格納するByte型配列
    Dim i As Long               'カウンタ
    
    'もし変換元と変換先が同じ文字コードだったら変換しない
    If kcFrom = kcTo Then
        KCConvert = bytSource
        Exit Function
    End If
    
    '変換元文字列が空のときは脱出
    If UBound(bytSource) < 0 Then
        KCConvert = bytSource
        Exit Function
    End If
    
    '変換先がKC_UNKNOWNだったら変換しない
    If kcTo = KC_UNKNOWN Then
        KCConvert = bytSource
        Exit Function
    End If
    
    '変換元がKC_UNKNOWNだったら自動認識
    If kcFrom = KC_UNKNOWN Then
        kcFrom = KCDetectCode(bytSource)
        '自動認識できなければ脱出
        If kcFrom = KC_UNKNOWN Then
            Exit Function
        End If
    End If
    
    '文字列を変換していく
    Select Case kcFrom
    Case KC_SHIFTJIS
        '変換先に従って処理
        Select Case kcTo
        Case KC_JIS
            bytResult = KCConvertShiftJISIntoJIS(bytSource)
        Case KC_EUC
            bytResult = KCConvertShiftJISIntoEUC(bytSource)
        Case Else
            'UTF-16Nに変換してからその先へ
            bytResult = StrConv(bytSource, vbUnicode)
            bytResult = KCConvert(bytResult, KC_UNICODESTRING, kcTo)
        End Select
    
    Case KC_JIS
        'ShiftJISに変換してからその先へ
        bytResult = KCConvertJISIntoShiftJIS(bytSource)
        bytResult = KCConvert(bytResult, KC_SHIFTJIS, kcTo)
    
    Case KC_EUC
        'ShiftJISに変換してからその先へ
        bytResult = KCConvertEUCIntoShiftJIS(bytSource)
        bytResult = KCConvert(bytResult, KC_SHIFTJIS, kcTo)
    
    Case KC_UTF8N
        '変換先に従って処理
        Select Case kcTo
        Case KC_UTF8BOM
            'BOMを付加
            ReDim bytResult(UBound(bytSource) + 3) As Byte
            bytResult(0) = bytUTF8BOM1
            bytResult(1) = bytUTF8BOM2
            bytResult(2) = bytUTF8BOM3
            For i = 3 To UBound(bytResult)
                bytResult(i) = bytSource(i - 3)
            Next
        Case Else
            'UTF-16Nに変換してからその先へ
            bytResult = KCConvertUTF8IntoUnicode(bytSource)
            bytResult = KCConvert(bytResult, KC_UNICODESTRING, kcTo)
        End Select
        
    Case KC_UTF8BOM
        'BOMを削除してUTF-8Nにしてからその先へ
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
        '変換先に従って処理
        Select Case kcTo
        Case KC_UTF8BOM
            bytResult = KCConvertUnicodeIntoUTF8(bytSource)
            bytResult = KCConvert(bytResult, KC_UTF8N, KC_UTF8BOM)
        Case KC_UTF8N
            bytResult = KCConvertUnicodeIntoUTF8(bytSource)
        Case KC_UTF16LE
            'BOMを付加
            ReDim bytResult(UBound(bytSource) + 2) As Byte
            bytResult(0) = bytUTF16BOM2
            bytResult(1) = bytUTF16BOM1
            For i = 2 To UBound(bytResult)
                bytResult(i) = bytSource(i - 2)
            Next
        Case KC_UTF16BE
            'BOMを付加してバイトオーダを逆に
            ReDim bytResult(UBound(bytSource) + 2) As Byte
            bytResult(0) = bytUTF16BOM1
            bytResult(1) = bytUTF16BOM2
            For i = 2 To UBound(bytResult) Step 2
                bytResult(i) = bytSource(i - 1)
                bytResult(i + 1) = bytSource(i - 2)
            Next
        Case Else
            'Shift-JISに変換してからその先へ
            bytResult = StrConv(bytSource, vbFromUnicode)
            bytResult = KCConvert(bytResult, KC_SHIFTJIS, kcTo)
        End Select
        
    Case KC_UTF16LE
        'BOMを削除してUTF-16Nにしてからその先へ
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
        'BOMを削除してバイトオーダを逆にし，UTF-16Nにしてからその先へ
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
        '呼び出し方が不正
        Err.Raise KC_ERROR_INVALID
    End Select
    
    KCConvert = bytResult
    
End Function

Private Function KCConvertShiftJISIntoJIS(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '■Shift-JISをJISに変換する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource()             変換元のShift-JIS文字列が入ったByte型配列
    '[戻り値]   変換された文字列が格納されたByte型配列
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '変換された文字列を格納するByte型配列
    Dim emMode As EncodeMode    '漢字かローマ字かカタカナかのフラグ
    Dim i As Long               'カウンタ
    Dim j As Long               'カウンタ
    
    ReDim bytResult(UBound(bytSource) * 5) As Byte     '大きめに用意しておく
    
    '変数の初期化
    i = 0
    j = 0
    emMode = mRome
    
    '変換処理
    Do While i <= UBound(bytSource())
        If (bytSource(i) >= &H80& And bytSource(i) <= &HA0) Or (bytSource(i) >= &HE0) Then
            '先頭ビットがたっていて、半角カナの領域にないとき
            '全角文字
            If Not emMode = mKanji Then
                'それまで全角文字でなかったときはJIS漢字INを入れる
                bytResult(j) = bytJISESC
                bytResult(j + 1) = bytKANJI1
                bytResult(j + 2) = bytKANJI2NEW
                
                'カウンタを進める
                j = j + 3
            End If
            
            '一旦そのまま代入
            bytResult(j) = bytSource(i)
            bytResult(j + 1) = bytSource(i + 1)
            
            '変換する
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
            
            '漢字であるフラグを立てる
            emMode = mKanji
            
            'カウンタを進める
            i = i + 2
            j = j + 2
        ElseIf bytSource(i) >= &HA1 And bytSource(i) <= &HDF Then
            '半角カタカナ
            
            If Not emMode = mKata Then
                'JISカタカナINを入れる
                bytResult(j) = bytJISESC
                bytResult(j + 1) = bytKATA1
                bytResult(j + 2) = bytKATA2
                
                'カウンタを進める
                j = j + 3
            End If
            
            'トップビットを下ろす
            bytResult(j) = bytSource(i) - &H80
            
            'カタカナであるフラグを立てる
            emMode = mKata
            
            i = i + 1
            j = j + 1
            
        Else
            'ANSI文字
            
            If Not emMode = mRome Then
                'JISローマ字INを入れる
                bytResult(j) = bytJISESC
                bytResult(j + 1) = bytROME1
                bytResult(j + 2) = bytROME2
                
                'カウンタを進める
                j = j + 3
            End If
            
            'そのまま
            bytResult(j) = bytSource(i)
            
            'ローマ字であるフラグを立てる
            emMode = mRome
            
            i = i + 1
            j = j + 1
        End If
    Loop
    
    '終了前には必ずローマ字に戻す
    If Not emMode = mRome Then
        'JISローマ字INを入れる
        bytResult(j) = bytJISESC
        bytResult(j + 1) = bytROME1
        bytResult(j + 2) = bytROME2
        
        'カウンタを進める
        j = j + 3
    End If
    
    'bytResultから必要な部分だけ抜き出す
    ReDim Preserve bytResult(j - 1) As Byte
    
    '戻り値に代入
    KCConvertShiftJISIntoJIS = bytResult()
    
End Function

Private Function KCConvertJISIntoShiftJIS(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '■JISをShift-JISに変換する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource()             変換元のJIS文字列が入ったByte型配列
    '[戻り値]   変換された文字列が格納されたByte型配列
    '-----------------------------------------------------------------------------------------
    'ESC ( Jで半角カタカナに入るものに対応するため、フラグは半角文字か全角文字かを分けるだけにしています。
    
    Dim bytResult() As Byte     '変換された文字列を格納するByte型配列
    Dim emMode As EncodeMode    '漢字かローマ字かカタカナかのフラグ
    Dim i As Long               'カウンタ
    Dim j As Long               'カウンタ
    
    ReDim bytResult(UBound(bytSource) * 5) As Byte     '大きめに用意しておく
    
    '変数の初期化
    i = 0
    j = 0
    emMode = mRome
    
    '変換処理
    Do While i <= UBound(bytSource())
        If bytSource(i) = bytJISESC Then
            'ESCだったとき
            If bytSource(i + 1) = bytKANJI1 And (bytSource(i + 2) = bytKANJI2OLD Or bytSource(i + 2) = bytKANJI2NEW) Then
                '漢字INであったとき
                
                '全角文字であるフラグを立てる
                emMode = mKanji
                
                'カウンタを進める
                i = i + 3
                
            ElseIf bytSource(i + 1) = bytROME1 And bytSource(i + 2) = bytROME2 Then
                'ローマ字INであったとき
                
                'ローマ字であるフラグを立てる
                emMode = mRome
                
                'カウンタを進める
                i = i + 3
                
            ElseIf bytSource(i + 1) = bytKATA1 And bytSource(i + 2) = bytKATA2 Then
                'カタカナINであったとき
                
                'カタカナであるフラグを立てる
                emMode = mKata
                
                'カウンタを進める
                i = i + 3
                
            Else
                '無視してカウンタを進める(これをしないと不正なファイルの時に止まってしまう)
                i = i + 1
                
            End If
            
        Else
            'ESCでなかったとき
            Select Case emMode
            Case mKanji
                '全角文字であるとき
                
                '一旦そのまま代入
                bytResult(j) = bytSource(i)
                bytResult(j + 1) = bytSource(i + 1)
                
                '変換
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
                
                'カウンタを進める
                i = i + 2
                j = j + 2
                
            Case mRome
                'ローマ字であるとき
                
                'そのまま
                bytResult(j) = bytSource(i)
                
                'カウンタを進める
                i = i + 1
                j = j + 1
            
            Case mKata
                '半角カナであるとき
                
                'トップビットを立てる
                bytResult(j) = bytSource(i) Or &H80
                
                'カウンタを進める
                i = i + 1
                j = j + 1
                
            End Select
        End If
        
        
    Loop
    
    'bytResultから必要な部分だけ抜き出す
    ReDim Preserve bytResult(j - 1) As Byte
    
    '戻り値に代入
    KCConvertJISIntoShiftJIS = bytResult()
    
End Function

Private Function KCConvertEUCIntoShiftJIS(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '■EUCをShiftJISに変換する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource()             変換元のEUC文字列が入ったByte型配列
    '[戻り値]   変換された文字列が格納されたByte型配列
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '変換された文字列を格納するByte型配列
    Dim i As Long               'カウンタ(変換元)
    Dim j As Long               'カウンタ(変換先)
    
    ReDim bytResult(UBound(bytSource) * 2) As Byte     '大きめに用意しておく
    
    '変数の初期化
    i = 0
    j = 0
    
    '変換処理
    Do While i <= UBound(bytSource())
        
        If bytSource(i) = bytSS2 Then
            '半角カナ
            
            '第1バイトを飛ばす
            bytResult(j) = bytSource(i + 1)
            
            'カウンタを進める
            i = i + 2
            j = j + 1
            
        ElseIf bytSource(i) = bytSS3 Then
            'JIS補助漢字(ShiftJISでは表現不可)
            
            bytResult(j) = &H3F     'とりあえず"?"に置き換える
            
            'カウンタを進める
            i = i + 3
            j = j + 1
            
        ElseIf bytSource(i) >= &H80& Then
            '通常の全角文字
            
            '変換する
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
            
            
            'カウンタを進める
            i = i + 2
            j = j + 2
        Else
            'ASCIIと判断
            
            'そのまま
            bytResult(j) = bytSource(i)
            
            i = i + 1
            j = j + 1
        End If
    Loop
    
    'bytResultから必要な部分だけ抜き出す
    ReDim Preserve bytResult(j - 1) As Byte
    
    '戻り値に代入
    KCConvertEUCIntoShiftJIS = bytResult()
    
End Function

Private Function KCConvertShiftJISIntoEUC(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '■ShiftJISをEUCに変換する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource()             変換元のJIS文字列が入ったByte型配列
    '[戻り値]   変換された文字列の入ったByte型配列
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '変換された文字列を格納するByte型配列
    Dim i As Long               'カウンタ(変換元)
    Dim j As Long               'カウンタ(変換先)
    
    ReDim bytResult(UBound(bytSource) * 2) As Byte     '大きめに用意しておく
    
    '変数の初期化
    i = 0
    j = 0
    
    '変換処理
    Do While i <= UBound(bytSource())
        If (bytSource(i) >= &H80& And bytSource(i) <= &HA0) Or (bytSource(i) >= &HE0) Then
            '先頭ビットがたっていて、半角カナの領域にないとき
            '全角文字
            
            '一旦そのまま代入
            bytResult(j) = bytSource(i)
            bytResult(j + 1) = bytSource(i + 1)
            
            '変換する
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
            
            'カウンタを進める
            i = i + 2
            j = j + 2
        
        ElseIf bytSource(i) >= &HA1 And bytSource(i) <= &HDF Then
            '半角カタカナ
            
            'SS2を挿入
            bytResult(j) = bytSS2
            
            bytResult(j + 1) = bytSource(i)
            
            i = i + 1
            j = j + 2
            
        Else
            'ASCII文字
            
            'そのまま
            bytResult(j) = bytSource(i)
            
            i = i + 1
            j = j + 1
        End If
    Loop
    
    'bytResultから必要な部分だけ抜き出す
    ReDim Preserve bytResult(j - 1) As Byte
    
    '戻り値に代入
    KCConvertShiftJISIntoEUC = bytResult()
    
End Function

Private Function KCConvertUnicodeIntoUTF8(ByRef bytSource() As Byte) As Byte()
    '*****************************************************************************************
    '■Unicode文字列をUTF-8に変換する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource               変換元のUTF-16文字列が入ったByte型配列
    '[戻り値]   変換された文字列が格納されたByte型配列
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '変換された文字列を格納するByte型配列
    Dim i As Long               'カウンタ
    Dim j As Long               'カウンタ
    Dim lngRet As Long
    
    ReDim bytResult(UBound(bytSource) * 5) As Byte     '大きめに用意しておく
    
    '変数の初期化
    i = 0
    j = 0
    
    For i = 0 To UBound(bytSource) Step 2
        '1文字をLong型に格納
        lngRet = CLng(bytSource(i)) + CLng(bytSource(i + 1)) * 256
        
        Select Case lngRet
        Case 0 To &H7F&
            'ASCII文字のとき
            bytResult(j) = CByte(lngRet)    'そのまま
            
            'カウンタを進める
            j = j + 1
            
        Case &H80& To &H7FF&
            '2バイト文字のとき
            
            '第1バイト
            bytResult(j) = CByte(&HC0 Or ((lngRet And &H7C0&) \ 64))
            '第2バイト
            bytResult(j + 1) = CByte(&H80 Or (lngRet And &H3F))
            
            'カウンタを進める
            j = j + 2
            
        Case &H800& To &HFFFF&
            '3バイト文字のとき
            
            '第1バイト
            bytResult(j) = CByte(&HE0 Or ((lngRet And &HF000&) \ 4096))
            '第2バイト
            bytResult(j + 1) = CByte(&H80 Or ((lngRet And &HFC0&) \ 64))
            '第3バイト
            bytResult(j + 2) = CByte(&H80 Or (lngRet And &H3F))
            
            'カウンタを進める
            j = j + 3
            
        End Select
    Next i
    
    'bytResultから必要な部分だけ抜き出す
    ReDim Preserve bytResult(j - 1) As Byte
    
    '戻り値に代入
    KCConvertUnicodeIntoUTF8 = bytResult()
    
End Function

Private Function KCConvertUTF8IntoUnicode(ByRef bytSource() As Byte) As String
    '*****************************************************************************************
    '■UTF-8をUnicode文字列に変換する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource               変換元のUTF-8文字列が入ったByte型配列
    '[戻り値]   変換された文字列が格納されたString
    '-----------------------------------------------------------------------------------------
    
    Dim bytResult() As Byte     '変換された文字列を格納するByte型配列
    Dim i As Long               'カウンタ
    Dim j As Long               'カウンタ
    Dim lngRet As Long
    
    ReDim bytResult(UBound(bytSource) * 3) As Byte      '大きめに用意しておく
    
    '変数の初期化
    i = 0
    j = 0
    
    '変換処理
    Do While i <= UBound(bytSource())
        Select Case bytSource(i)
        Case 0 To &H7F
            'ASCII文字だったとき
            
            bytResult(j) = bytSource(i)
            bytResult(j + 1) = 0
            
            'カウンタを進める
            i = i + 1
            j = j + 2
            
        Case &HC2 To &HDF
            '2バイト文字
            
            bytResult(j) = (bytSource(i) And &H3&) * 64 Or (bytSource(i + 1) And &H3F)
            bytResult(j + 1) = (bytSource(i) And &H1C) \ 4
            
            'カウンタを進める
            i = i + 2
            j = j + 2
            
        Case &HE0 To &HEF
            '3バイト文字
        
            bytResult(j) = (bytSource(i + 1) And &H3) * 64 Or (bytSource(i + 2) And &H3F)
            bytResult(j + 1) = (bytSource(i) And &HF) * 16 Or (bytSource(i + 1) And &H3C) \ 4
            
            'カウンタを進める
            i = i + 3
            j = j + 2
            
        Case &HF0 To &HF7
            '4バイト文字(VBのUnicode(UCS-2)では表現不可)
            
            bytResult(j) = (bytSource(i + 2) And &H3) * 64 Or (bytSource(i + 3) And &H3F)
            bytResult(j + 1) = (bytSource(i + 1) And &HF) * 16 Or (bytSource(i + 2) And &H3C) \ 4
            
            'カウンタを進める
            i = i + 4
            j = j + 2
            
        Case &HF8 To &HFB
            '5バイト文字(VBのUnicode(UCS-2)では表現不可)
            
            bytResult(j) = (bytSource(i + 3) And &H3) * 64 Or (bytSource(i + 4) And &H3F)
            bytResult(j + 1) = (bytSource(i + 2) And &HF) * 16 Or (bytSource(i + 3) And &H3C) \ 4
            
            'カウンタを進める
            i = i + 5
            j = j + 2
            
        Case &HFC To &HFD
            '6バイト文字(VBのUnicode(UCS-2)では表現不可)
            
            bytResult(j) = (bytSource(i + 4) And &H3) * 64 Or (bytSource(i + 5) And &H3F)
            bytResult(j + 1) = (bytSource(i + 3) And &HF) * 16 Or (bytSource(i + 4) And &H3C) \ 4
            
            'カウンタを進める
            i = i + 6
            j = j + 2
            
        Case Else
            '不正な文字列
            Err.Raise KC_ERROR_INVALID
            
        End Select
    Loop
    
    'bytResultから必要な部分だけ抜き出す
    ReDim Preserve bytResult(j - 1) As Byte
    
    '戻り値に代入
    KCConvertUTF8IntoUnicode = bytResult()
    
End Function

Public Function KCDetectCode(ByRef bytSource() As Byte) As KCCode
    '*****************************************************************************************
    '■文字コードを判別する
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource               判断されるコードの文字列が格納されたByte型配列
    '[戻り値]   判断されたコードをしめす列挙型KCCode(モジュール先頭を参照してください)
    '-----------------------------------------------------------------------------------------
    Dim i As Long                       'カウンタ
    Dim lngCountStart As Long           'カウンタの最初の添え字
    Dim bolIsAscii As Boolean           'ASCIIかどうかのフラグ
    
    '変数初期化
    lngCountStart = 0
    
    'BOMをチェック
    If UBound(bytSource()) >= 3 Then
        If bytSource(0) = bytUTF8BOM1 And bytSource(1) = bytUTF8BOM2 And bytSource(2) = bytUTF8BOM3 Then
            'UTF-8BOM？
            If KCIsValidUTF8(bytSource, True) Then
                KCDetectCode = KC_UTF8BOM
                Exit Function
            End If
        End If
    End If
    If UBound(bytSource()) >= 2 Then
        If bytSource(0) = bytUTF16BOM2 And bytSource(1) = bytUTF16BOM1 Then
            'UTF-16LE？
            If UBound(bytSource()) Mod 2 = 1 Then   'バイト数が奇数でないことだけを確認
                KCDetectCode = KC_UTF16LE
                Exit Function
            End If
        ElseIf bytSource(0) = bytUTF16BOM1 And bytSource(1) = bytUTF16BOM2 Then
            'UTF-16BE？
            If UBound(bytSource()) Mod 2 = 1 Then   'バイト数が奇数でないことだけを確認
                KCDetectCode = KC_UTF16BE
                Exit Function
            End If
        End If
    End If
    
    'まずASCIIかJISかを判別
    bolIsAscii = True
    For i = 0 To UBound(bytSource())
        If Not KCInRange(bytSource(i), &H0, &H7F) Then
            'ASCIIではない(もちろんJISでもない)
            bolIsAscii = False
            Exit For
        ElseIf bytSource(i) = bytJISESC Then
            'ESCが出てきているのでJISの可能性アリ
            If KCIsValidJIS(bytSource) Then
                KCDetectCode = KC_JIS
                Exit Function
            End If
        End If
    Next i
    If bolIsAscii Then
        'ASCIIだった(マルチバイト文字が使われていない)ときはShiftJISとして返す
        KCDetectCode = KC_SHIFTJIS
        Exit Function
    End If
    
    'ShiftJISか判定
    If KCIsValidShiftJIS(bytSource()) Then
        KCDetectCode = KC_SHIFTJIS
        Exit Function
    End If
    
    'EUCか判定
    If KCIsValidEUC(bytSource()) Then
        KCDetectCode = KC_EUC
        Exit Function
    End If
    
    'UTF-8Nかを判定
    If KCIsValidUTF8(bytSource(), False) Then
        KCDetectCode = KC_UTF8N
        Exit Function
    End If
    
    'ここまでやって判別できなければUCS2か？
    If UBound(bytSource()) Mod 2 = 1 Then   'バイト数が奇数でないことだけを確認
        KCDetectCode = KC_UNICODESTRING
        Exit Function
    End If
    
    '判別不能
    KCDetectCode = KC_UNKNOWN
    
End Function

Private Function KCIsValidJIS(ByRef bytSource() As Byte) As Boolean
    '*****************************************************************************************
    '■正しいJIS文字列かどうかを判別する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource               判断される文字列が格納されたByte型配列
    '[戻り値]   正しいならばTrue,不正な文字列ならばFalse
    '-----------------------------------------------------------------------------------------
    
    Dim i As Long               'カウンタ
    Dim emMode As EncodeMode    '漢字かローマ字かカタカナかのフラグ
    
    emMode = mRome
    KCIsValidJIS = False
    
    On Error GoTo Err_Handler
    
    For i = 0 To UBound(bytSource)
        If bytSource(i) = bytJISESC Then
            'ESCだったとき
            If bytSource(i + 1) = bytKANJI1 And (bytSource(i + 2) = bytKANJI2OLD Or bytSource(i + 2) = bytKANJI2NEW) Then
                '漢字INであったとき
                emMode = mKanji
                i = i + 2
            ElseIf bytSource(i + 1) = bytROME1 And bytSource(i + 2) = bytROME2 Then
                'ローマ字INであったとき
                emMode = mRome
                i = i + 2
            ElseIf bytSource(i + 1) = bytKATA1 And bytSource(i + 2) = bytKATA2 Then
                'カタカナINであったとき
                emMode = mKata
                i = i + 2
            Else
                'ESCだけがあったときはJISであるとは判断しない
                Exit Function
            End If
        Else
            Select Case emMode
            Case mRome
                If Not KCInRange(bytSource(i), &H0, &H7F) Then
                    '不正
                    Exit Function
                End If
            Case Else
                If Not KCInRange(bytSource(i), &H20, &H7F) Then
                    '不正
                    Exit Function
                End If
            End Select
        End If
    Next
    
    '正当
    KCIsValidJIS = True
    
Err_Handler:
    'エラーが生じたら不正
    
End Function

Private Function KCIsValidShiftJIS(ByRef bytSource() As Byte) As Boolean
    '*****************************************************************************************
    '■正しいShiftJIS文字列かどうかを判別する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource               判断される文字列が格納されたByte型配列
    '[戻り値]   正しいならばTrue,不正な文字列ならばFalse
    '-----------------------------------------------------------------------------------------
    Dim i As Long               'カウンタ
    
    KCIsValidShiftJIS = False
    
    On Error GoTo Err_Handler
    
    For i = 0 To UBound(bytSource)
        If KCInRange(bytSource(i), &H81, &H9F) Or KCInRange(bytSource(i), &HE0, &HFC) Then
            '全角文字
            If Not (KCInRange(bytSource(i + 1), &H40, &H7E) Or KCInRange(bytSource(i + 1), &H80, &HFC)) Then
                '不正
                Exit Function
            End If
            i = i + 1
        ElseIf Not (KCInRange(bytSource(i), &HA1, &HDF) Or KCInRange(bytSource(i), &H0, &H7F)) Then
            'ASCIIや半角カナでなければ不正
            Exit Function
        End If
    Next i
    
    '正しい文字列である
    KCIsValidShiftJIS = True
    
Err_Handler:
    'エラーが生じたら不正
    
End Function

Private Function KCIsValidEUC(ByRef bytSource() As Byte) As Boolean
    '*****************************************************************************************
    '■正しいEUC文字列かどうかを判別する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource               判断される文字列が格納されたByte型配列
    '[戻り値]   正しいならばTrue,不正な文字列ならばFalse
    '-----------------------------------------------------------------------------------------
    Dim i As Long               'カウンタ
    
    KCIsValidEUC = False
    
    On Error GoTo Err_Handler
    
    For i = 0 To UBound(bytSource)
        If bytSource(i) = bytSS2 Then
            If Not KCInRange(bytSource(i + 1), &HA1, &HDF) Then
                '不正
                Exit Function
            End If
            i = i + 1
        ElseIf bytSource(i) = bytSS3 Then
            If Not (KCInRange(bytSource(i + 1), &HA1, &HFE) And KCInRange(bytSource(i + 2), &HA1, &HFE)) Then
                '不正
                Exit Function
            End If
            i = i + 2
        ElseIf KCInRange(bytSource(i), &HA1, &HFE) Then
            If Not KCInRange(bytSource(i + 1), &HA1, &HFE) Then
                '不正
                Exit Function
            End If
            i = i + 1
        ElseIf Not KCInRange(bytSource(i), &H0, &H7F) Then
            '不正
            Exit Function
        End If
    Next i
    
    '正しい文字列である
    KCIsValidEUC = True
    
Err_Handler:
    'エラーが生じたら不正
    
End Function

Private Function KCIsValidUTF8(ByRef bytSource() As Byte, ByVal bolBOM As Boolean) As Boolean
    '*****************************************************************************************
    '■正しいUTF-8文字列かどうかを判別する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   bytSource               判断される文字列が格納されたByte型配列
    '           bolBOM                  BOMありならばTrue,なしならばFalse
    '[戻り値]   正しいならばTrue,不正な文字列ならばFalse
    '-----------------------------------------------------------------------------------------
    
    Dim i As Long               'カウンタ
    
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
            '不正な文字列
            Exit Function
        End If
        
    Next i
    
    '正しい文字列である
    KCIsValidUTF8 = True
    
Err_Handler:
    'エラーが生じたら不正
    
End Function

Private Function KCInRange(ByVal lngTarget As Long, ByVal lngLowerBound As Long, ByVal lngUpperBound As Long) As Boolean
    '*****************************************************************************************
    '■範囲内に入っているか判定する(内部関数)
    '-----------------------------------------------------------------------------------------
    '[ 引数 ]   lngTarget               判定される数値
    '           lngLowerBound           下限
    '           lngUpperBound           上限
    '[戻り値]   範囲内ならばTrue,範囲外ならばFalse
    '-----------------------------------------------------------------------------------------

    KCInRange = (lngTarget >= lngLowerBound) And (lngTarget <= lngUpperBound)
End Function
