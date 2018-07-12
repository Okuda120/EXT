Imports CommonEXT

''' <summary>
''' EXTZ0209
''' </summary>
''' <remarks>付帯設備コピー
''' <para>作成情報：2015/08/31 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTZ0210

    Private ppStrSelectedDate As String
    Private ppStrSelectedDateDisp As String
    Private ppAryRiyobi As Array
    Private ppLstSelectedRiyobi As ArrayList
    Private ppBlnChangeFlg As Boolean
    Private commonLogicEXT As New CommonLogicEXT      '共通ロジッククラス

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0210_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.lblRiyobi.Text = ppStrSelectedDateDisp
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbCopyDate.ActiveSheet
        sheet.RowCount = ppAryRiyobi.Length
        sheet.ColumnCount = 3
        'カラム設定(コード列の非表示)
        sheet.Columns(2).Visible = False
        Dim index As Integer = 0
        Dim dt As Date
        For Each riyobiDisp As String In ppAryRiyobi
            dt = riyobiDisp
            sheet.Cells(index, 0).Value = False
            sheet.Cells(index, 1).Value = dt.ToString(CommonDeclareEXT.FMT_DATE_DISP)
            sheet.Cells(index, 2).Value = riyobiDisp
            index = index + 1
        Next

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 戻るボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        ppBlnChangeFlg = False
        Me.Close()
    End Sub

    ''' <summary>
    ''' 入力完了処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnComplate_Click(sender As Object, e As EventArgs) Handles btnComplate.Click
        Dim blnSelect As Boolean = False
        ppLstSelectedRiyobi = New ArrayList
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbCopyDate.ActiveSheet
        Dim index As New Integer
        Do While index < sheet.Rows.Count
            If sheet.Cells(index, 0).Value = True Then
                ppLstSelectedRiyobi.Add(sheet.Cells(index, 2).Value)
                blnSelect = True
            End If
            index = index + 1
        Loop

        If blnSelect = False Then
            MsgBox(String.Format(CommonEXT.E0002, "コピー先の利用日"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ppBlnChangeFlg = True
        Me.Close()
    End Sub


    ''' <summary>
    ''' プロパティセット【コピー元の曜日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSelectedDate</returns>
    ''' <remarks><para>作成情報：2015/08/31 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSelectedDate()
        Get
            Return ppStrSelectedDate
        End Get
        Set(ByVal value)
            ppStrSelectedDate = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【コピー元の曜日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSelectedDateDisp</returns>
    ''' <remarks><para>作成情報：2015/08/31 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSelectedDateDisp()
        Get
            Return ppStrSelectedDateDisp
        End Get
        Set(ByVal value)
            ppStrSelectedDateDisp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日一覧】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppAryRiyobi</returns>
    ''' <remarks><para>作成情報：2015/08/31 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropAryRiyobi()
        Get
            Return ppAryRiyobi
        End Get
        Set(ByVal value)
            ppAryRiyobi = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択した利用日一覧】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLstSelectedRiyobi</returns>
    ''' <remarks><para>作成情報：2015/08/31 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLstSelectedRiyobi()
        Get
            Return ppLstSelectedRiyobi
        End Get
        Set(ByVal value)
            ppLstSelectedRiyobi = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsApproval</returns>
    ''' <remarks><para>作成情報：2015/08/31 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnChangeFlg()
        Get
            Return ppBlnChangeFlg
        End Get
        Set(ByVal value)
            ppBlnChangeFlg = value
        End Set
    End Property

End Class
