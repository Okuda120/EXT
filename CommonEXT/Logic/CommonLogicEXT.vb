Imports Common
Imports System.Windows.Forms
Imports FarPoint.Win.Spread
Imports System.Text.RegularExpressions
Imports System.Configuration
Imports System.IO
Imports Npgsql

''' <summary>
''' CommonLogicEXT
''' </summary>
''' <remarks>EXT内の共通プロシージャ、ファンクションを定義したクラス
''' <para>作成情報：
''' <p>改定情報：</p>
''' </para></remarks>
Public Class CommonLogicEXT

    'インスタンス作成
    Private commonLogic As New CommonLogic          '共通ロジッククラス

    ''' <summary>
    ''' DB登録値変換処理（数値）
    ''' </summary>
    ''' <param name="strVal">対象文字列</param>
    ''' <returns>boolean  true OK false NG </returns>
    ''' <remarks>Decimal、DBNull.Valueを返却する
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改定情報：</p>
    ''' </para></remarks>
    Public Function convInsNumStr(ByVal strVal As String) As Object

        If (strVal = String.Empty) = True Then
            Return DBNull.Value
        End If
        Return Decimal.Parse(strVal)
    End Function

    ''' <summary>
    ''' DB登録値変換処理（数値）
    ''' </summary>
    ''' <param name="intVal"></param>
    ''' <returns>boolean  true OK false NG </returns>
    ''' <remarks>Decimal、DBNull.Valueを返却する
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改定情報：</p>
    ''' </para></remarks>
    Public Function convInsNum(ByVal intVal As Integer) As Object
        If (intVal = Nothing) = True Then
            Return DBNull.Value
        End If
        Return intVal
    End Function

    ''' <summary>
    ''' DB登録値変換処理（文字列）
    ''' </summary>
    ''' <param name="strVal">対象文字列</param>
    ''' <returns>boolean  true OK false NG </returns>
    ''' <remarks>文字列、DBNull.Valueを返却する
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改定情報：</p>
    ''' </para></remarks>
    Public Function convInsStr(ByVal strVal As String) As Object

        If (strVal = String.Empty) = True Then
            Return DBNull.Value
        End If
        Return strVal
    End Function

    ''' <summary>
    ''' 数値への変換処理
    ''' </summary>
    ''' <param name="Val">数値変換対象の文字列</param>
    ''' <returns>Double 変換後の値 </returns>
    ''' <remarks>ObjectからDoubleへの返還を行う。Exceptionが発生した場合は0を返却する。
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function convInteger(ByVal Val As Object) As Integer
        If String.IsNullOrEmpty(Val) Or Val Is DBNull.Value Then
            Return Nothing
        End If
        Dim iVal As Integer
        Try
            iVal = Val.Replace(",", "")
            iVal = Integer.Parse(iVal)
        Catch ex As Exception
            Return Nothing
        End Try
        Return iVal
    End Function

    ''' <summary>
    ''' 文字列から数値の変換処理
    ''' </summary>
    ''' <param name="Val">数値変換対象の文字列</param>
    ''' <returns>Double 変換後の値 </returns>
    ''' <remarks>StringからDoubleへの返還を行う。Exceptionが発生した場合は0を返却する。
    ''' <para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function convDouble(ByVal Val As String) As Double
        Dim dVal As Double
        Try
            dVal = Double.Parse(Val)
        Catch ex As Exception
            Return 0
        End Try
        Return dVal
    End Function

    ''' <summary>
    ''' ファイル選択ダイアログ表示処理
    ''' </summary>
    ''' <param name="strInitFile">[IN]初期表示ファイル名</param>
    ''' <param name="strInitPath">[IN]初期表示ディレクトリ</param>
    ''' <param name="strFileType">[IN]選択可能なファイル形式</param>
    ''' <param name="intSelFileType">[IN]選択ファイル形式</param>
    ''' <param name="txtFilePath">[OUT]選択したファイルパス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ファイル選択ダイアログ表示し、ファイルを選択する
    ''' <para>作成情報：2013/10/18 m.kato
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function OpenFileDialog(ByVal strInitFile As String, _
                                   ByVal strInitPath As String, _
                                   ByVal strFileType As String, _
                                   ByVal intSelFileType As Integer, _
                                   ByRef txtFilePath As String) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim ofdTorikomiFile As New OpenFileDialog

        Try
            '初期表示ファイル名設定
            ofdTorikomiFile.FileName = strInitFile

            '初期表示ディレクトリ設定
            ofdTorikomiFile.InitialDirectory = strInitPath

            '選択ファイル形式設定
            ofdTorikomiFile.Filter = strFileType

            '選択されているファイル形式を設定
            ofdTorikomiFile.FilterIndex = intSelFileType

            'ダイアログを閉じる前に現在のディレクトリを復元
            ofdTorikomiFile.RestoreDirectory = True

            'ファイル選択ダイアログの名前を設定
            ofdTorikomiFile.Title = "ファイルを開く"

            'ダイアログを表示
            If ofdTorikomiFile.ShowDialog() = DialogResult.OK Then
                'データをセット
                txtFilePath = ofdTorikomiFile.FileName
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 日付チェック処理
    ''' </summary>
    ''' <param name="strDate">[IN]チェックする日付yyyy/mm/dd形式</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>正しい日付かどうかチェックする
    ''' <para>作成情報：2013/10/18 m.kato
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function IsDate(ByVal strDate As String) As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim intYear As Integer
        Dim intMonth As Integer
        Dim intDay As Integer

        Try
            'ブランクの場合、チェックしない
            If Trim(strDate) = "" Then
                Return True
            End If

            intYear = CInt(Mid(strDate, 1, 4))
            intMonth = CInt(Mid(strDate, 6, 2))
            intDay = CInt(Mid(strDate, 9, 2))

            If (DateTime.MinValue.Year > intYear) OrElse (intYear > DateTime.MaxValue.Year) Then
                Return False
            End If

            If (DateTime.MinValue.Month > intMonth) OrElse (intMonth > DateTime.MaxValue.Month) Then
                Return False
            End If

            Dim iLastDay As Integer = DateTime.DaysInMonth(intYear, intMonth)

            If (DateTime.MinValue.Day > intDay) OrElse (intDay > iLastDay) Then
                Return False
            End If

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function


    ''' <summary>
    ''' メールアドレス書式チェック処理
    ''' </summary>
    ''' <param name="strChkValue">[IN]チェック対象文字列</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>メールアドレスとして正しい書式かチェックする
    ''' <para>作成情報：2012/06/20 fukuo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function IsMailAddress(ByVal strChkValue As String) As Boolean

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim r As New System.Text.RegularExpressions.Regex( _
        "\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b", _
        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        Dim strReturn As Boolean

        Try
            'メールアドレスに一致する対象があるか検索
            Dim m As System.Text.RegularExpressions.Match = r.Match(strChkValue)

            If m.Success = True Then
                m = m.NextMatch()
                If m.Success = False Then
                    strReturn = True
                Else
                    strReturn = False
                End If
            Else
                strReturn = False
            End If

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理
            Return strReturn

        Catch ex As Exception
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            strReturn = Nothing
            Return strReturn
        Finally

        End Try

    End Function


    ''' <summary>
    ''' 一覧行フォーカス設定処理
    ''' </summary>
    ''' <param name="vwTarget">[IN/OUT]対象スプレッド</param>
    ''' <param name="intTargetSheet">[IN]対象スプレッドシート番号</param>
    ''' <param name="intTargetRow">[IN]フォーカス設定開始行番号</param>
    ''' <param name="intTargetCol">[IN]フォーカス設定開始列番号</param>
    ''' <param name="intFocusColCnt">[IN]フォーカス設定行数</param>
    ''' <param name="intFocusRowCnt">[IN]フォーカス設定列数</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>一覧において入力チェックに引っかかった行にフォーカスを設定する
    ''' <para>作成情報：2012/06/21 t.fukuo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function SetFocusOnVwRow(ByRef vwTarget As FpSpread, _
                                    ByVal intTargetSheet As Integer, _
                                    ByVal intTargetRow As Integer, _
                                    ByVal intTargetCol As Integer, _
                                    ByVal intFocusRowCnt As Integer, _
                                    ByVal intFocusColCnt As Integer) As Boolean

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            'フォーカス対象セルを設定
            vwTarget.Sheets(intTargetSheet).SetActiveCell(intTargetRow, intTargetCol)

            'フォーカス対象行に選択範囲を設定
            vwTarget.Sheets(intTargetSheet).AddSelection(intTargetRow, intTargetCol, intFocusRowCnt, intFocusColCnt)

            'フォーカス対象行を表示する
            vwTarget.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Center, FarPoint.Win.Spread.HorizontalPosition.Center)

            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '処理正常終了
            Return True

        Catch ex As Exception
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try
    End Function


    ''' <summary>
    ''' 全角カナチェック
    ''' </summary>
    ''' <param name="argTxt">[IN]チェック対象文字列</param>
    ''' <returns>Boolean  チェック結果    true  ＯＫ  false  ＮＧ</returns>
    ''' <remarks>対象文字列が全角カナ文字のみかチェックする
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改定情報：</p>
    ''' </para></remarks>
    Public Function IsFullKana(ByVal argTxt As String) As Boolean

        Dim i As Integer
        Dim length As Integer

        length = Len(argTxt)

        If length = 0 Then
            Return True
        End If

        For i = 1 To length
            If Not Mid(argTxt, i, 1) Like "[ァ-ー　]" Then
                Return False
            End If
        Next i

        Return True

    End Function

    ''' <summary>
    ''' 全角英数字を半角へ変換
    ''' </summary>
    ''' <param name="strTarget">変換したい文字列</param>
    ''' <returns>半角変換後の文字列</returns>
    ''' <remarks>全角英数字を半角英数字へ変換する
    ''' <para>作成情報：2012/06/06 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function ChangeToHankakuStr(ByVal strTarget As String) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim re As Regex = New Regex("[０-９Ａ-Ｚａ-ｚ：＿＊＋？／・（）「」｛｝＜＞＝～｜￥，－　]+")
        Dim output As String = re.Replace(strTarget, AddressOf MyReplacerHankaku)

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return output

    End Function

    ''' <summary>
    ''' 半角変換実行
    ''' </summary>
    ''' <param name="m">変換したい文字列</param>
    ''' <returns>変換後の文字列</returns>
    ''' <remarks>対象文字列を半角に変換する
    ''' <para>作成情報：2012/06/06 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Private Function MyReplacerHankaku(ByVal m As Match) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return Strings.StrConv(m.Value, VbStrConv.Narrow, 0)
    End Function

    ''' <summary>
    ''' 大文字を小文字へ変換
    ''' </summary>
    ''' <param name="strTarget">変換したい文字列</param>
    ''' <returns>小文字変換後の文字列</returns>
    ''' <remarks>大文字を小文字へ変換する
    ''' <para>作成情報：2012/06/08 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function ChangeToLowerCase(ByVal strTarget As String) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim re As Regex = New Regex("[Ａ-ＺA-Z]+")
        Dim output As String = re.Replace(strTarget, AddressOf MyReplacerLowerCase)

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return output

    End Function

    ''' <summary>
    ''' 小文字変換実行
    ''' </summary>
    ''' <param name="m">変換したい文字列</param>
    ''' <returns>小文字変換後の文字列</returns>
    ''' <remarks>対象文字列を特定の文字列に変換する
    ''' <para>作成情報：2012/06/08 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Private Function MyReplacerLowerCase(ByVal m As Match) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return Strings.StrConv(m.Value, VbStrConv.Lowercase, 0)
    End Function

    ''' <summary>
    ''' ひらがなをカナへ変換
    ''' </summary>
    ''' <param name="strTarget">変換したい文字列</param>
    ''' <returns>カナ変換後の文字列</returns>
    ''' <remarks>ひらがなを全角カナへ変換する
    ''' <para>作成情報：2012/06/06 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function ChangeToKatakana(ByVal strTarget As String) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return StrConv(strTarget, Microsoft.VisualBasic.VbStrConv.Katakana, &H411)

    End Function

    ''' <summary>
    ''' 半角カナを全角カナに変換
    ''' </summary>
    ''' <param name="strTarget">変換したい文字列</param>
    ''' <returns>全角カナ変換後の文字列</returns>
    ''' <remarks>半角カナを全角カナへ変換する
    ''' <para>作成情報：2012/06/06 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function ChangeToZenkakuKana(ByVal strTarget As String) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim i As Long
        Dim strTemp As String = ""
        Dim strKana As String = ""
        Dim chrKana As String = ""

        For i = 1& To Len(strTarget)
            chrKana = Mid$(strTarget, i, 1&)
            Select Case Asc(chrKana)
                Case 166 To 223
                    '半角が続いたら文字をつなぐ
                    strKana = strKana & chrKana
                Case Else
                    '全角文字になったら半角の未処理文字を全部全角に変換（これにより濁点処理等が不要）
                    If Len(strKana) > 0& Then
                        strTemp = strTemp & StrConv(strKana, vbWide)
                        strKana = vbNullString
                    End If
                    strTemp = strTemp & chrKana
            End Select
        Next i

        '最後の文字が半角の場合の処理
        If Len(strKana) > 0& Then
            strTemp = strTemp & StrConv(strKana, vbWide)
        End If

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return strTemp

    End Function

    ''' <summary>
    ''' 半角全角スペース除去
    ''' </summary>
    ''' <param name="strTarget">スペース除去したい文字列</param>
    ''' <returns>スペース除去後の文字列</returns>
    ''' <remarks>対象文字列から半角全角のスペースを除去する
    ''' <para>作成情報：2012/06/08 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function RemoveSpace(ByVal strTarget As String) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)


        '半角全角スペース、およびタブ文字を除去した文字列を返す
        Return strTarget.Replace(" ", "").Replace("　", "").Replace(ControlChars.Tab, "")

    End Function

    ''' <summary>
    ''' 改行コード除去
    ''' </summary>
    ''' <param name="strTarget">改行コードを除去したい文字列</param>
    ''' <returns>改行コード除去後の文字列</returns>
    ''' <remarks>対象文字列から改行コードを除去する
    ''' <para>作成情報：2012/06/08 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function RemoveVbCr(ByVal strTarget As String) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)


        '改行コードを除去した文字列を返す
        Return strTarget.Replace(System.Environment.NewLine, "")

    End Function

    ''' <summary>
    ''' スプレッドセル：NothingもしくはDBNullの文字列変換
    ''' </summary>
    ''' <param name="cellTarget">[IN]変換対象したいセル</param>
    ''' <param name="strChangeVal">[IN]セルがNothingもしくはDBNullだった場合の変換文字列</param>
    ''' <returns>DBNullだった場合の変換の文字列</returns>
    ''' <remarks>対象文字列がNothingもしくはDBNullだった場合、指定された文字列に変換して返す
    ''' <para>作成情報：2012/06/21 t.fukuo
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function ChangeNothingToStr(ByVal cellTarget As Cell, _
                                       ByVal strChangeVal As String) As String

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim strReturn As String = strChangeVal

        'NothingもしくはDBNullの場合は指定された文字列に変換し、そうでない場合はセルの値を文字列変換して返す
        If cellTarget IsNot Nothing AndAlso _
           cellTarget.Value IsNot Nothing AndAlso _
           cellTarget.Value IsNot DBNull.Value Then

            strReturn = cellTarget.Value.ToString()

        End If

        'ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return strReturn

    End Function

    ''' <summary>
    ''' 都道府県リストの設定
    ''' </summary>
    ''' <param name="cmbBox"></param>
    ''' <remarks></remarks>
    Public Sub TodohukenLst(ByRef cmbBox As ComboBox)
        cmbBox.Items.Add("北海道")
        cmbBox.Items.Add("青森県")
        cmbBox.Items.Add("岩手県")
        cmbBox.Items.Add("宮城県")
        cmbBox.Items.Add("秋田県")
        cmbBox.Items.Add("山形県")
        cmbBox.Items.Add("福島県")
        cmbBox.Items.Add("茨城県")
        cmbBox.Items.Add("栃木県")
        cmbBox.Items.Add("群馬県")
        cmbBox.Items.Add("埼玉県")
        cmbBox.Items.Add("千葉県")
        cmbBox.Items.Add("東京都")
        cmbBox.Items.Add("神奈川県")
        cmbBox.Items.Add("新潟県")
        cmbBox.Items.Add("富山県")
        cmbBox.Items.Add("石川県")
        cmbBox.Items.Add("福井県")
        cmbBox.Items.Add("山梨県")
        cmbBox.Items.Add("長野県")
        cmbBox.Items.Add("岐阜県")
        cmbBox.Items.Add("静岡県")
        cmbBox.Items.Add("愛知県")
        cmbBox.Items.Add("三重県")
        cmbBox.Items.Add("滋賀県")
        cmbBox.Items.Add("京都府")
        cmbBox.Items.Add("大阪府")
        cmbBox.Items.Add("兵庫県")
        cmbBox.Items.Add("奈良県")
        cmbBox.Items.Add("和歌山県")
        cmbBox.Items.Add("鳥取県")
        cmbBox.Items.Add("島根県")
        cmbBox.Items.Add("岡山県")
        cmbBox.Items.Add("広島県")
        cmbBox.Items.Add("山口県")
        cmbBox.Items.Add("徳島県")
        cmbBox.Items.Add("香川県")
        cmbBox.Items.Add("愛媛県")
        cmbBox.Items.Add("高知県")
        cmbBox.Items.Add("福岡県")
        cmbBox.Items.Add("佐賀県")
        cmbBox.Items.Add("長崎県")
        cmbBox.Items.Add("熊本県")
        cmbBox.Items.Add("大分県")
        cmbBox.Items.Add("宮崎県")
        cmbBox.Items.Add("鹿児島県")
        cmbBox.Items.Add("沖縄県")
    End Sub

    ''' <summary>
    ''' DBNULLの値をNothingで返す
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="colNm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DbNullToNothing(ByVal row As DataRow, ByVal colNm As String) As Object
        If row(colNm) Is DBNull.Value Then
            Return Nothing
        End If
        Return row(colNm)
    End Function

    ''' <summary>
    ''' Nothingの値をDBNullで返す
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DbNothingToNull(ByVal val As Object) As Object
        If val Is Nothing Then
            Return DBNull.Value
        End If
        Return val
    End Function

    ''' <summary>
    ''' カンマ編集数値への変換処理
    ''' </summary>
    ''' <param name="Val">数値変換対象の文字列</param>
    ''' <returns>Double 変換後の値 </returns>
    ''' <remarks>ObjectからDoubleへの返還を行う。Exceptionが発生した場合は0を返却する。
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function convKingaku(ByVal Val As Object) As String
        If Val Is Nothing Then
            Return Nothing
        End If
        Dim iVal As String
        Try
            iVal = Integer.Parse(Val).ToString("#,##0")
        Catch ex As Exception
            Return ""
        End Try
        Return iVal
    End Function

    ''' <summary>
    ''' 年月への変換処理
    ''' </summary>
    ''' <param name="Val">数値変換対象の文字列</param>
    ''' <returns>Double 変換後の値 </returns>
    ''' <remarks>ObjectからDoubleへの返還を行う。Exceptionが発生した場合は0を返却する。
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function convYmDateStr(ByVal Val As String) As String
        If String.IsNullOrEmpty(Val) Or Val Is DBNull.Value Then
            Return Nothing
        End If
        Dim iVal As String
        Try
            Val = Val.Replace("/", "")
            iVal = Val.Substring(0, 4) + "年" + Val.Substring(4, 2) + "月"
        Catch ex As Exception
            Return Nothing
        End Try
        Return iVal
    End Function

    ''' <summary>
    ''' 利用者レベル名(1-一般 2-要注意 3-利用不可)
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRiyoshalvlNm(ByVal val As String) As String
        If val Is Nothing Then
            Return ""
        End If
        If val = RIYOSHA_LV1 Then
            Return "一般"
        ElseIf val = RIYOSHA_LV2 Then
            Return "要注意"
        ElseIf val = RIYOSHA_LV3 Then
            Return "利用不可"
        End If
        Return ""
    End Function

    ''' <summary>
    ''' フォーム背景色設定
    ''' </summary>
    ''' <param name="strSystemConfFlg"></param>
    ''' <returns>フォーム背景色
    '''          ・環境設定フラグが0（本番機）：白
    '''          ・環境設定フラグが1（検証機）：青
    ''' </returns>
    ''' <remarks>引数（環境設定フラグ）に応じたフォーム背景色を返す
    ''' <para>作成情報： 2015.11.18 h.hagiwara
    ''' <p>   改定情報：</p>
    '''</para> </remarks>
    Public Function SetFormBackColor(ByVal strSystemConfFlg As String) As Color

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim clrBack As Color = Nothing

        If strSystemConfFlg IsNot DBNull.Value Then
            If strSystemConfFlg = "0" Then
                ' 本番環境用の背景色（白）を返す
                clrBack = PropBackColorHONBAN
            Else
                clrBack = PropBackColorKENSHOU
            End If
        End If

        '終了ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return clrBack

    End Function

#Region "レプリケーション対応"

    ''' <summary>
    ''' DB接続確認処理
    ''' </summary>
    ''' <returns>Boolean True:正常終了 False:異常終了(接続不可)</returns>
    ''' <remarks>DB接続の確認と設定を行う
    ''' <para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CheckDBCondition() As Boolean

        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'Net Useコマンド用変数宣言
        Dim strCmd As String = ""                 'コマンド文字列
        Dim strDriveName1 As String = ""          '使用論理ドライブ名(メイン)
        Dim strDriveName2 As String = ""          '使用論理ドライブ名(サブ)
        Dim strDummyFilePath As String = ""       'ダミーファイルパス
        Dim strNetUseUserID As String = ""        'NetUseユーザID
        Dim strNetUsePassword As String = ""      'NetUseパスワード

        '接続確認
        Dim strSql As String = "SELECT * FROM pg_stat_replication" 'マスタＤＢ確認用ＳＱＬ
        Dim bolConErr As Boolean = False

        Try
            'ＤＢ接続先確認用設定取得
            If CommonDeclareEXT.PropDBCheck Is Nothing Then
                CommonDeclareEXT.PropMainDB = System.Configuration.ConfigurationManager.AppSettings("MainDB")
                CommonDeclareEXT.PropDummyFilePath1 = System.Configuration.ConfigurationManager.AppSettings("DummyFilePath1")
                CommonDeclareEXT.PropDummyFilePath2 = System.Configuration.ConfigurationManager.AppSettings("DummyFilePath2")
                CommonDeclareEXT.PropDummyFileName = System.Configuration.ConfigurationManager.AppSettings("DummyFileName")
                CommonDeclareEXT.PropDBString1 = System.Configuration.ConfigurationManager.AppSettings("DbString1")
                CommonDeclareEXT.PropDBString2 = System.Configuration.ConfigurationManager.AppSettings("DbString2")
                CommonDeclareEXT.PropNetUseUserID1 = System.Configuration.ConfigurationManager.AppSettings("NetUseUserID1")
                CommonDeclareEXT.PropNetUsePassword1 = System.Configuration.ConfigurationManager.AppSettings("NetUsePassword1")
                CommonDeclareEXT.PropNetUseUserID2 = System.Configuration.ConfigurationManager.AppSettings("NetUseUserID2")
                CommonDeclareEXT.PropNetUsePassword2 = System.Configuration.ConfigurationManager.AppSettings("NetUsePassword2")
                CommonDeclareEXT.PropDBCheck = System.Configuration.ConfigurationManager.AppSettings("DBCheck")
                CommonDeclareEXT.PropDebugDB = System.Configuration.ConfigurationManager.AppSettings("DebugDB")
                If CommonDeclareEXT.PropNetUseUserID1.Equals("") Then
                    CommonDeclareEXT.PropNetUseUserID1 = NET_USE_USERID1
                End If
                If CommonDeclareEXT.PropNetUsePassword1.Equals("") Then
                    CommonDeclareEXT.PropNetUsePassword1 = NET_USE_PASSWORD1
                End If
                If CommonDeclareEXT.PropNetUseUserID2.Equals("") Then
                    CommonDeclareEXT.PropNetUseUserID2 = NET_USE_USERID2
                End If
                If CommonDeclareEXT.PropNetUsePassword2.Equals("") Then
                    CommonDeclareEXT.PropNetUsePassword2 = NET_USE_PASSWORD2
                End If
            End If

            ' ============================================================================

            'DB接続先チェック有無確認
            If CommonDeclareEXT.PropDBCheck.Equals("0") Then
                'チェックなしの場合は、常に「DbString」へ接続
                Return True
            End If

            'メイン接続先設定
            If CommonDeclareEXT.PropMainDB.Equals("1") Then
                strDummyFilePath = CommonDeclareEXT.PropDummyFilePath1
                strNetUseUserID = CommonDeclareEXT.PropNetUseUserID1
                strNetUsePassword = CommonDeclareEXT.PropNetUsePassword1
            Else
                strDummyFilePath = CommonDeclareEXT.PropDummyFilePath2
                strNetUseUserID = CommonDeclareEXT.PropNetUseUserID2
                strNetUsePassword = CommonDeclareEXT.PropNetUsePassword2
            End If

            'NetUse設定
            If NetUseConect(strDummyFilePath, strNetUseUserID, strNetUsePassword) = False Then
                Return False
            End If

            'ダミーファイル存在チェック(メインDB)
            If File.Exists(Path.Combine(CommonDeclareEXT.PropDummyFilePath1, _
                                        CommonDeclareEXT.PropDummyFileName)) = True Then
                'メインＤＢの接続先を設定
                If CommonDeclareEXT.PropMainDB.Equals("1") Then
                    DbString = CommonDeclareEXT.PropDBString1
                Else
                    DbString = CommonDeclareEXT.PropDBString2
                End If
                'DB接続チェック
                Dim Cn As New NpgsqlConnection(DbString)
                Dim Adapter As New NpgsqlDataAdapter
                Dim Table As New DataTable()
                Try
                    Cn.Open()
                    Adapter.SelectCommand = New NpgsqlCommand(strSql, Cn)
                    Adapter.Fill(Table)
                    If Table.Rows.Count < 1 Then
                        bolConErr = True
                    End If
#If DEBUG Then
                    'DEBUGモードの場合、「DebugDB」をマスタＤＢとしてチェックする
                    If (CommonDeclareEXT.PropDebugDB IsNot Nothing AndAlso _
                        CommonDeclareEXT.PropDebugDB.Equals("") = False) And _
                        CommonDeclareEXT.PropDebugDB.Equals(CommonDeclareEXT.PropMainDB) Then
                        bolConErr = False
                    End If
#End If
                Catch ex As Exception
                    'ログ出力
                    commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
                    bolConErr = True
                Finally
                    '終了処理
                    Cn.Dispose()
                    Adapter.Dispose()
                    Table.Dispose()
                End Try
            Else
                bolConErr = True
            End If

            'メインＤＢ接続判定
            If bolConErr = True Then
                'メインＤＢで接続不可の場合、接続先を変更して再度確認する
                If CommonDeclareEXT.PropMainDB.Equals("1") Then
                    strDummyFilePath = CommonDeclareEXT.PropDummyFilePath2
                    strNetUseUserID = CommonDeclareEXT.PropNetUseUserID2
                    strNetUsePassword = CommonDeclareEXT.PropNetUsePassword2
                Else
                    strDummyFilePath = CommonDeclareEXT.PropDummyFilePath1
                    strNetUseUserID = CommonDeclareEXT.PropNetUseUserID1
                    strNetUsePassword = CommonDeclareEXT.PropNetUsePassword1
                End If

                'NetUse設定
                If NetUseConect(strDummyFilePath, strNetUseUserID, strNetUsePassword) = False Then
                    Return False
                End If

                bolConErr = False

                'ダミーファイル存在チェック(サブ)
                If File.Exists(Path.Combine(CommonDeclareEXT.PropDummyFilePath2, _
                                            CommonDeclareEXT.PropDummyFileName)) = True Then
                    'DB接続文字列を切り替え
                    If CommonDeclareEXT.PropMainDB.Equals("1") Then
                        DbString = CommonDeclareEXT.PropDBString2
                    Else
                        DbString = CommonDeclareEXT.PropDBString1
                    End If
                    'DB接続チェック
                    Dim Cn As New NpgsqlConnection(DbString)
                    Dim Adapter As New NpgsqlDataAdapter
                    Dim Table As New DataTable()
                    Try
                        Cn.Open()
                        Adapter.SelectCommand = New NpgsqlCommand(strSql, Cn)
                        Adapter.Fill(Table)
                        If Table.Rows.Count < 1 Then
                            bolConErr = True
                        End If
#If DEBUG Then
                        'DEBUGモードの場合、「DebugDB」をマスタＤＢとしてチェックする
                        If CommonDeclareEXT.PropDebugDB IsNot Nothing AndAlso _
                           CommonDeclareEXT.PropDebugDB.Equals("") = False Then
                            If CommonDeclareEXT.PropMainDB.Equals("1") Then
                                If CommonDeclareEXT.PropDebugDB.Equals("2") Then
                                    bolConErr = False
                                End If
                            Else
                                If CommonDeclareEXT.PropDebugDB.Equals("1") Then
                                    bolConErr = False
                                End If
                            End If
                        End If
#End If
                    Catch ex As Exception
                        'ログ出力
                        commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
                        bolConErr = True
                    Finally
                        '終了処理
                        Cn.Dispose()
                        Adapter.Dispose()
                        Table.Dispose()
                    End Try
                Else
                    bolConErr = True
                End If

                'サブＤＢ接続判定
                If bolConErr = True Then
                    'エラーの場合
                    'メッセージ変数にエラーメッセージを格納
                    puErrMsg = EXT_E001 & E9999
                    Return False
                Else
                    '接続できた場合
                    Dim config As System.Configuration.Configuration = _
                                  System.Configuration.ConfigurationManager.OpenExeConfiguration( _
                                  System.Configuration.ConfigurationUserLevel.None)
                    'メインとサブを切り替えて保持する
                    If CommonDeclareEXT.PropMainDB.Equals("1") Then
                        CommonDeclareEXT.PropMainDB = "2"
                    Else
                        CommonDeclareEXT.PropMainDB = "1"
                    End If
                    'ファイル設定書換え
                    config.AppSettings.Settings("MainDB").Value = CommonDeclareEXT.PropMainDB
                    config.Save()
                End If
            End If

            ' ============================================================================

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False
        Finally
            '接続した論理ドライブの削除
            NetUseConectDel(CommonDeclareEXT.PropDummyFilePath1)
            NetUseConectDel(CommonDeclareEXT.PropDummyFilePath2)
            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
        End Try

        '正常処理終了
        Return True

    End Function


    ''' <summary>
    ''' NetUse接続処理
    ''' </summary>
    ''' <param name="strDummyFilePath">[IN]ダミーファイルパス</param>
    ''' <param name="strUserID">[IN]ユーザID</param>
    ''' <param name="strPassword">[IN]パスワード</param>
    ''' <returns>boolean 終了状況    True:正常  False:異常</returns>
    ''' <remarks>NetUseで接続する文字列を設定し接続を行う
    ''' <para>作成情報：2015/11/18 e.okamura
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function NetUseConect(ByVal strDummyFilePath As String, _
                                 ByVal strUserID As String, ByVal strPassword As String) As Boolean
        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strCmd As String = ""
        'プロセスクラスの宣言
        Dim p As Process = Nothing                              'プロセスクラス
        Dim psi As New System.Diagnostics.ProcessStartInfo()    'プロセススタートインフォクラス

        Try

            psi.FileName = System.Environment.GetEnvironmentVariable("ComSpec")

            '出力を読み取れるようにする
            psi.RedirectStandardInput = False
            psi.RedirectStandardOutput = True
            psi.UseShellExecute = False
            'ウィンドウを非表示にする
            psi.CreateNoWindow = True

            'コマンドの設定
            strCmd = "/C net use " & strDummyFilePath & " " & strPassword & " /user:" & strUserID & " /persistent:no"

            psi.Arguments = strCmd
            p = Process.Start(psi)
            p.WaitForExit()

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

    ''' <summary>
    ''' NetUse接続削除処理
    ''' </summary>
    ''' <param name="strDummyFilePath">[IN]ダミーファイルパス</param>
    ''' <returns>boolean 終了状況    True:正常  False:異常</returns>
    ''' <remarks>NetUseで接続した論理ドライブを削除する
    ''' <para>作成情報：2015/11/18 e.okamura
    ''' <p>改定情報：</p>
    ''' </para></remarks> 
    Public Function NetUseConectDel(ByVal strDummyFilePath As String) As Boolean
        '開始ログ出力
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strCmd As String = ""
        'プロセスクラスの宣言
        Dim p As Process = Nothing                              'プロセスクラス
        Dim psi As New System.Diagnostics.ProcessStartInfo()    'プロセススタートインフォクラス

        Try

            psi.FileName = System.Environment.GetEnvironmentVariable("ComSpec")

            '出力を読み取れるようにする
            psi.RedirectStandardInput = False
            psi.RedirectStandardOutput = True
            psi.UseShellExecute = False
            'ウィンドウを非表示にする
            psi.CreateNoWindow = True

            '接続の解除
            strCmd = "/C net use " & strDummyFilePath & " /delete /y"
            psi.Arguments = strCmd
            p = Process.Start(psi)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

#End Region

    ''' <summary>
    ''' 年月への変換処理
    ''' </summary>
    ''' <param name="Val">数値変換対象の文字列</param>
    ''' <returns>Double 変換後の値 </returns>
    ''' <remarks>ObjectからDoubleへの返還を行う。Exceptionが発生した場合は0を返却する。
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function convYmDateStr2(ByVal Val As String) As String
        If String.IsNullOrEmpty(Val) Or Val Is DBNull.Value Then
            Return Nothing
        End If
        Dim iVal As String
        Try
            Val = Val.Replace("/", "")
            iVal = Val.Substring(0, 4) & "年" & Integer.Parse(Val.Substring(4, 2)) & "月"
        Catch ex As Exception
            Return Nothing
        End Try
        Return iVal
    End Function

    ''' <summary>
    ''' 数値への変換処理
    ''' </summary>
    ''' <param name="Val">数値変換対象の文字列</param>
    ''' <returns>Long 変換後の値 </returns>
    ''' <remarks>ObjectからLongへの返還を行う。Exceptionが発生した場合は0を返却する。
    ''' <para>作成情報：2015.12.21 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function convLong(ByVal Val As Object) As Long
        If String.IsNullOrEmpty(Val) Or Val Is DBNull.Value Then
            Return Nothing
        End If
        Dim iVal As Long
        Try
            iVal = Val.ToString.Replace(",", "")
            iVal = Long.Parse(iVal)
        Catch ex As Exception
            Return Nothing
        End Try
        Return iVal
    End Function

End Class
