Imports Common
Imports CommonEXT
Public Class EXTZ0202

    Private ppStrTorokuKbn As String   '登録区分
    Private ppStrYoyakuNo As String   '予約No
    Private ppStrShisetuKbn As String   '施設区分
    Private ppStrStudioKbn As String   'スタジオ区分
    Private ppLstRiyobi As ArrayList   '利用日一覧

    '変数宣言
    Public dataCommon As New CommonDataEXT      '共通データクラス
    Public logicEXTZ0202 As New LogicEXTZ0202   'ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT    '共通ロジッククラス

#Region "「利用日画面」ロード時処理 "

    ''' <summary>
    ''' 「利用日画面」ロード時処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「利用日画面」ロード時の処理を行う 
    ''' <para>作成情報： 2015.11.18 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Private Sub EXTZ0202_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

#End Region

    Private Sub ButtonAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim dtFrom As Date
        Dim dtTo As Date
        Dim aryStrRiyobi As New ArrayList
        Dim aryStrRiyobiDisp As New ArrayList
        Try

            ' DBレプリケーション
            If commonLogicEXT.CheckDBCondition() = False Then
                'メッセージを出力 
                MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
                Exit Sub
            End If

            If Me.dtpRiyobiFrom.txtDate.Text = "" And Me.dtpRiyobiTo.txtDate.Text = "" Then
                MsgBox(String.Format(CommonDeclareEXT.E0001, "利用日"), MsgBoxStyle.Exclamation, "エラー")
                Return
            ElseIf Me.dtpRiyobiFrom.txtDate.Text = "" Then
                dtTo = Me.dtpRiyobiTo.txtDate.Text
                dtFrom = Me.dtpRiyobiTo.txtDate.Text
            ElseIf Me.dtpRiyobiTo.txtDate.Text = "" Then
                dtFrom = Me.dtpRiyobiFrom.txtDate.Text
                dtTo = Me.dtpRiyobiFrom.txtDate.Text
            Else
                dtFrom = Me.dtpRiyobiFrom.txtDate.Text
                dtTo = Me.dtpRiyobiTo.txtDate.Text
            End If

            If Date.Compare(dtFrom, dtTo) > 0 Then
                MsgBox(String.Format(CommonDeclareEXT.E2010, "利用日From", "利用日To"), MsgBoxStyle.Exclamation, "エラー")
                Return
            End If

            'ループ
            Do While Date.Compare(dtFrom, dtTo) < 1
                aryStrRiyobi.Add(dtFrom.ToString(CommonDeclareEXT.FMT_DATE))
                dtFrom = dtFrom.AddDays(1)
            Loop

            '利用日追加処理
            Dim dataRiyobi As New CommonDataRiyobi
            Dim dataCancelWait As New CommonDataCancel
            Dim aryAddRiyobi As New ArrayList '登録用リスト
            Dim dtDisp As Date '表示用
            If PropStrTorokuKbn = TOUROKU_KBN_KARI Or PropStrTorokuKbn = TOUROKU_KBN_SEISHIKI Then
                '仮予約／正式予約
                For Each strRiyobi As String In aryStrRiyobi
                    With dataRiyobi
                        .PropStrYoyakuNo = PropStrYoyakuNo
                        .PropStrYoyakuDt = strRiyobi
                        .PropStrShisetuKbn = PropStrShisetuKbn
                        .PropStrStudioKbn = PropStrStudioKbn
                        .PropStrRegistFlg = "0"
                    End With
                    If logicEXTZ0202.CheckRiyobi(PropStrTorokuKbn, dataRiyobi, PropLstRiyobi) = False Then
                        MsgBox(String.Format(CommonDeclareEXT.E2017, ""), MsgBoxStyle.Exclamation, "エラー")
                        Return
                    End If
                    dtDisp = strRiyobi
                    dataRiyobi.PropStrYoyakuDtDisp = dtDisp.ToString(CommonDeclareEXT.FMT_DATE_DISP)
                    PropLstRiyobi.Add(dataRiyobi)
                    aryAddRiyobi.Add(dataRiyobi)
                    dataRiyobi = New CommonDataRiyobi
                Next
                '登録処理
                logicEXTZ0202.InsertYoyakuCtlData(aryAddRiyobi)
            Else
                'キャンセル待ち登録
                For Each strRiyobi As String In aryStrRiyobi
                    With dataCancelWait
                        .PropStrCancelWaitNo = PropStrYoyakuNo
                        .PropStrCancelDt = strRiyobi
                        .PropStrShisetuKbn = PropStrShisetuKbn
                        .PropStrStudioKbn = PropStrStudioKbn
                        .PropStrRegistFlg = "0"
                    End With
                    If logicEXTZ0202.CheckCancelWait(PropStrTorokuKbn, dataCancelWait, PropLstRiyobi) = False Then
                        MsgBox(String.Format(CommonDeclareEXT.E2017, ""), MsgBoxStyle.Exclamation, "エラー")
                        Return
                    End If
                    dtDisp = strRiyobi
                    dataCancelWait.PropStrCancelDtDisp = dtDisp.ToString(CommonDeclareEXT.FMT_DATE_DISP)
                    dataCancelWait.PropStrCancelYm = dtDisp.ToString(CommonDeclareEXT.FMT_DATE_YM)
                    dataCancelWait.PropStrKibobiKbn = "1"
                    PropLstRiyobi.Add(dataCancelWait)
                    aryAddRiyobi.Add(dataCancelWait)
                    dataCancelWait = New CommonDataCancel
                Next
                '登録処理
                logicEXTZ0202.InsertCancelCtlData(aryAddRiyobi)
            End If

        Catch ex As Exception
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
        End Try

        '画面を閉じる
        Me.Close()

    End Sub

    ''' <summary>
    ''' 戻るボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        '画面を閉じる
        Me.Close()
    End Sub

    ''' <summary>
    ''' プロパティセット【登録区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTorokuKbn</returns>
    Public Property PropStrTorokuKbn()
        Get
            Return ppStrTorokuKbn
        End Get
        Set(ByVal value)
            ppStrTorokuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuNo</returns>
    Public Property PropStrYoyakuNo()
        Get
            Return ppStrYoyakuNo
        End Get
        Set(ByVal value)
            ppStrYoyakuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetuKbn</returns>
    Public Property PropStrShisetuKbn()
        Get
            Return ppStrShisetuKbn
        End Get
        Set(ByVal value)
            ppStrShisetuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStudioKbn</returns>
    Public Property PropStrStudioKbn()
        Get
            Return ppStrStudioKbn
        End Get
        Set(ByVal value)
            ppStrStudioKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLstRiyobi</returns>
    Public Property PropLstRiyobi()
        Get
            Return ppLstRiyobi
        End Get
        Set(ByVal value)
            ppLstRiyobi = value
        End Set
    End Property

End Class
