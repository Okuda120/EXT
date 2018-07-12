Imports Common
Imports CommonEXT
Imports EXTZ
Imports FarPoint.Win.Spread
Imports FarPoint.Win.Spread.Model

''' <summary>
''' EXTZ0206
''' </summary>
''' <remarks>
''' <para>作成情報：2015/09/01 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTZ0206

    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Public dataEXTZ0206 As New DataEXTZ0206             'データクラス
    Private commonLogicEXT As New CommonLogicEXT        '共通クラス

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0206_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dataEXTZ0206.PropIntWariai1 = 3
        dataEXTZ0206.PropIntWariai2 = 7
        Me.txtWariai1.Text = dataEXTZ0206.PropIntWariai1
        Me.txtWariai2.Text = dataEXTZ0206.PropIntWariai2
        'Spread
        Dim sheetFrom As FarPoint.Win.Spread.SheetView = Me.fbFrom.ActiveSheet
        sheetFrom.RowCount = 1
        sheetFrom.ColumnCount = 3
        sheetFrom.Cells(0, 0).Value = dataEXTZ0206.PropStrFromName
        sheetFrom.Cells(0, 0).Locked = True
        sheetFrom.Cells(0, 1).Value = dataEXTZ0206.PropIntFromKakutei
        sheetFrom.Cells(0, 1).Locked = True
        sheetFrom.Cells(0, 2).Value = dataEXTZ0206.PropStrFromNaiyo
        sheetFrom.Cells(0, 2).Locked = True

        Dim sheetTo As FarPoint.Win.Spread.SheetView = Me.fbTo.ActiveSheet
        sheetTo.RowCount = 2
        sheetTo.ColumnCount = 3
        sheetTo.Cells(0, 0).Value = dataEXTZ0206.PropStrFromName
        sheetTo.Cells(0, 0).Locked = True
        sheetTo.Cells(1, 0).Value = dataEXTZ0206.PropStrFromName
        sheetTo.Cells(1, 0).Locked = True
        sheetTo.Cells(0, 2).Value = dataEXTZ0206.PropStrFromNaiyo
        sheetTo.Cells(0, 2).Locked = True
        sheetTo.Cells(1, 2).Value = dataEXTZ0206.PropStrFromNaiyo
        sheetTo.Cells(1, 2).Locked = True
        calc()

        ' 背景色設定
        Me.BackColor = CommonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 分割割合による計算処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub calc()
        Dim wari1 As Integer = dataEXTZ0206.PropIntWariai1
        Dim wari2 As Integer = dataEXTZ0206.PropIntWariai2
        Dim wariTtl As New Integer
        'Dim fromKakutei As Integer = dataEXTZ0206.PropIntFromKakutei                ' 2015.12.21 UPD h.hagiwara
        Dim fromKakutei As Long = dataEXTZ0206.PropIntFromKakutei                    ' 2015.12.21 UPD h.hagiwara
        Dim wari1calc As Long = 0                                                    ' 2015.12.21 ADD h.hagiwara 
        Dim wari2calc As Long = 0                                                    ' 2015.12.21 ADD h.hagiwara 

        '計算
        wariTtl = wari1 + wari2
        'wari1 = Math.Round((fromKakutei / wariTtl) * wari1, MidpointRounding.AwayFromZero)                                                    ' 2015.12.21 UPD h.hagiwara 
        'wari2 = Math.Round((fromKakutei / wariTtl) * wari2, MidpointRounding.AwayFromZero)                                                    ' 2015.12.21 UPD h.hagiwara 
        wari1calc = Math.Round((fromKakutei / wariTtl) * wari1, MidpointRounding.AwayFromZero)                                                 ' 2015.12.21 UPD h.hagiwara 
        wari2calc = Math.Round((fromKakutei / wariTtl) * wari2, MidpointRounding.AwayFromZero)                                                 ' 2015.12.21 UPD h.hagiwara 

        'Me.lblSagaku.Text = fromKakutei - (wari1 + wari2)                                                                                     ' 2015.12.21 UPD h.hagiwara 
        Me.lblSagaku.Text = fromKakutei - (wari1calc + wari2calc)                                                                              ' 2015.12.21 UPD h.hagiwara 
        Dim sheetTo As FarPoint.Win.Spread.SheetView = Me.fbTo.ActiveSheet
        ' 2015.12.21 ADD START↓ h.hagiwara
        Dim numCellKeijyo As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellKeijyo.MaximumValue = 99999999999
        numCellKeijyo.ShowSeparator = True
        numCellKeijyo.DecimalPlaces = 0
        ' 2015.12.21 ADD END↑ h.hagiwara

        sheetTo.Cells(0, 1).CellType = numCellKeijyo                                 ' 2015.12.21 ADD h.hagiwara
        'sheetTo.Cells(0, 1).Value = wari1                                           ' 2015.12.21 UPD h.hagiwara 
        sheetTo.Cells(0, 1).Value = wari1calc                                        ' 2015.12.21 UPD h.hagiwara 
        sheetTo.Cells(1, 1).CellType = numCellKeijyo                                 ' 2015.12.21 ADD h.hagiwara
        'sheetTo.Cells(1, 1).Value = wari2                                           ' 2015.12.21 UPD h.hagiwara 
        sheetTo.Cells(1, 1).Value = wari2calc                                        ' 2015.12.21 UPD h.hagiwara 
    End Sub

    ''' <summary>
    ''' 再表示ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        If String.IsNullOrEmpty(Me.txtWariai1.Text) = True Then
            MsgBox(Format(CommonEXT.E0001, "分割割合(左)"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If String.IsNullOrEmpty(Me.txtWariai2.Text) = True Then
            MsgBox(Format(CommonEXT.E0001, "分割割合(右)"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        dataEXTZ0206.PropIntWariai1 = Me.txtWariai1.Text
        dataEXTZ0206.PropIntWariai2 = Me.txtWariai2.Text
        calc()
    End Sub

    ''' <summary>
    ''' 分割差額の自動計算
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub ChangeEventHandler(ByVal sender As Object, ByVal e As ChangeEventArgs) Handles fbTo.Change
        If 1 <> e.Column Then
            Exit Sub
        End If
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTo.ActiveSheet
        'Dim kingaku1 As Integer = sheet.Cells(0, 1).Value                                   ' 2015.12.21 UPD h.hagiwara
        'Dim kingaku2 As Integer = sheet.Cells(1, 1).Value                                   ' 2015.12.21 UPD h.hagiwara
        Dim kingaku1 As Long = sheet.Cells(0, 1).Value                                       ' 2015.12.21 UPD h.hagiwara
        Dim kingaku2 As Long = sheet.Cells(1, 1).Value                                       ' 2015.12.21 UPD h.hagiwara
        Me.lblSagaku.Text = dataEXTZ0206.PropIntFromKakutei - (kingaku1 + kingaku2)
    End Sub

    ''' <summary>
    ''' 戻るボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        dataEXTZ0206.PropBlnChangeFlg = False
        Me.Close()
    End Sub

    ''' <summary>
    ''' 入力完了処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnComplate_Click(sender As Object, e As EventArgs) Handles btnComplate.Click
        If Me.lblSagaku.Text > 0 Then
            MsgBox(CommonEXT.E2037, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbTo.ActiveSheet
        dataEXTZ0206.PropIntToKakutei1 = sheet.Cells(0, 1).Value
        dataEXTZ0206.PropIntToKakutei2 = sheet.Cells(1, 1).Value
        dataEXTZ0206.PropBlnChangeFlg = True
        Me.Close()
    End Sub
End Class
