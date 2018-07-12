<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0208
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim CheckBoxCellType1 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0208))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dtpNyukinTo = New Common.DateTimePickerEx()
        Me.dtpNyukinFrom = New Common.DateTimePickerEx()
        Me.dtpNyukinYoteiTo = New Common.DateTimePickerEx()
        Me.dtpNyukinYoteiFrom = New Common.DateTimePickerEx()
        Me.txtSeikyuIraiNo = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtSeikyuNo = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtAiteNm = New System.Windows.Forms.TextBox()
        Me.txtAiteCd = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.btnKakutei = New System.Windows.Forms.Button()
        Me.fbResult = New FarPoint.Win.Spread.FpSpread()
        Me.fbResult_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.fbResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbResult_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(12, 18)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(150, 21)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "分割請求とする請求書"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(70, 525)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(0, 13)
        Me.Label2.TabIndex = 79
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(542, 659)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 35)
        Me.btnBack.TabIndex = 89
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.dtpNyukinTo)
        Me.GroupBox1.Controls.Add(Me.dtpNyukinFrom)
        Me.GroupBox1.Controls.Add(Me.dtpNyukinYoteiTo)
        Me.GroupBox1.Controls.Add(Me.dtpNyukinYoteiFrom)
        Me.GroupBox1.Controls.Add(Me.txtSeikyuIraiNo)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtSeikyuNo)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.btnSearch)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtAiteNm)
        Me.GroupBox1.Controls.Add(Me.txtAiteCd)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(916, 134)
        Me.GroupBox1.TabIndex = 111
        Me.GroupBox1.TabStop = False
        '
        'dtpNyukinTo
        '
        Me.dtpNyukinTo.Location = New System.Drawing.Point(701, 45)
        Me.dtpNyukinTo.Name = "dtpNyukinTo"
        Me.dtpNyukinTo.Size = New System.Drawing.Size(144, 25)
        Me.dtpNyukinTo.TabIndex = 140
        '
        'dtpNyukinFrom
        '
        Me.dtpNyukinFrom.Location = New System.Drawing.Point(527, 45)
        Me.dtpNyukinFrom.Name = "dtpNyukinFrom"
        Me.dtpNyukinFrom.Size = New System.Drawing.Size(144, 25)
        Me.dtpNyukinFrom.TabIndex = 139
        '
        'dtpNyukinYoteiTo
        '
        Me.dtpNyukinYoteiTo.Location = New System.Drawing.Point(290, 45)
        Me.dtpNyukinYoteiTo.Name = "dtpNyukinYoteiTo"
        Me.dtpNyukinYoteiTo.Size = New System.Drawing.Size(144, 25)
        Me.dtpNyukinYoteiTo.TabIndex = 138
        '
        'dtpNyukinYoteiFrom
        '
        Me.dtpNyukinYoteiFrom.Location = New System.Drawing.Point(113, 45)
        Me.dtpNyukinYoteiFrom.Name = "dtpNyukinYoteiFrom"
        Me.dtpNyukinYoteiFrom.Size = New System.Drawing.Size(144, 25)
        Me.dtpNyukinYoteiFrom.TabIndex = 137
        '
        'txtSeikyuIraiNo
        '
        Me.txtSeikyuIraiNo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtSeikyuIraiNo.Location = New System.Drawing.Point(346, 81)
        Me.txtSeikyuIraiNo.Name = "txtSeikyuIraiNo"
        Me.txtSeikyuIraiNo.Size = New System.Drawing.Size(88, 20)
        Me.txtSeikyuIraiNo.TabIndex = 136
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.Location = New System.Drawing.Point(250, 84)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(103, 21)
        Me.Label7.TabIndex = 135
        Me.Label7.Text = "請求依頼番号"
        '
        'txtSeikyuNo
        '
        Me.txtSeikyuNo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtSeikyuNo.Location = New System.Drawing.Point(113, 81)
        Me.txtSeikyuNo.Name = "txtSeikyuNo"
        Me.txtSeikyuNo.Size = New System.Drawing.Size(88, 20)
        Me.txtSeikyuNo.TabIndex = 134
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(17, 84)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(103, 21)
        Me.Label5.TabIndex = 133
        Me.Label5.Text = "EXAS請求NO"
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(771, 93)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(97, 26)
        Me.btnSearch.TabIndex = 112
        Me.btnSearch.Text = "検索"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(681, 49)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 21)
        Me.Label4.TabIndex = 127
        Me.Label4.Text = "～"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(263, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 16)
        Me.Label3.TabIndex = 126
        Me.Label3.Text = "～"
        '
        'txtAiteNm
        '
        Me.txtAiteNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtAiteNm.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.txtAiteNm.Location = New System.Drawing.Point(357, 13)
        Me.txtAiteNm.Name = "txtAiteNm"
        Me.txtAiteNm.Size = New System.Drawing.Size(232, 20)
        Me.txtAiteNm.TabIndex = 125
        '
        'txtAiteCd
        '
        Me.txtAiteCd.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtAiteCd.Location = New System.Drawing.Point(113, 13)
        Me.txtAiteCd.Name = "txtAiteCd"
        Me.txtAiteCd.Size = New System.Drawing.Size(88, 20)
        Me.txtAiteCd.TabIndex = 124
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(453, 49)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 21)
        Me.Label1.TabIndex = 121
        Me.Label1.Text = "入金日"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(17, 50)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(103, 21)
        Me.Label6.TabIndex = 118
        Me.Label6.Text = "入金予定日"
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label19.Location = New System.Drawing.Point(17, 16)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(103, 21)
        Me.Label19.TabIndex = 117
        Me.Label19.Text = "相手先コード"
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label18.Location = New System.Drawing.Point(261, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(103, 21)
        Me.Label18.TabIndex = 116
        Me.Label18.Text = "相手先名"
        '
        'btnKakutei
        '
        Me.btnKakutei.Location = New System.Drawing.Point(852, 659)
        Me.btnKakutei.Name = "btnKakutei"
        Me.btnKakutei.Size = New System.Drawing.Size(96, 35)
        Me.btnKakutei.TabIndex = 112
        Me.btnKakutei.Text = "確定"
        Me.btnKakutei.UseVisualStyleBackColor = True
        '
        'fbResult
        '
        Me.fbResult.AccessibleDescription = "fbResult, Sheet1, Row 0, Column 0, "
        Me.fbResult.Location = New System.Drawing.Point(12, 173)
        Me.fbResult.Name = "fbResult"
        Me.fbResult.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbResult_Sheet1})
        Me.fbResult.Size = New System.Drawing.Size(1638, 456)
        Me.fbResult.TabIndex = 136
        '
        'fbResult_Sheet1
        '
        Me.fbResult_Sheet1.Reset()
        Me.fbResult_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbResult_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbResult_Sheet1.ColumnCount = 16
        Me.fbResult_Sheet1.RowCount = 0
        Me.fbResult_Sheet1.ActiveColumnIndex = -1
        Me.fbResult_Sheet1.ActiveRowIndex = -1
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "相手先ｺｰﾄﾞ"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "利用開始日"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用終了日"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "相手先名"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "予約番号"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "催事名"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "入金予定日"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "入金日"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "入金額"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "入金確認"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "請求日"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "入力日"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "EXAS請求NO"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "請求依頼番号（入力摘要）"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "モニタNO"
        Me.fbResult_Sheet1.ColumnHeader.Rows.Get(0).Height = 34.0!
        Me.fbResult_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.fbResult_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbResult_Sheet1.Columns.Get(0).Label = "選択"
        Me.fbResult_Sheet1.Columns.Get(0).Width = 40.0!
        Me.fbResult_Sheet1.Columns.Get(1).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(1).Label = "相手先ｺｰﾄﾞ"
        Me.fbResult_Sheet1.Columns.Get(1).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(2).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(2).Label = "利用開始日"
        Me.fbResult_Sheet1.Columns.Get(2).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(3).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(3).Label = "利用終了日"
        Me.fbResult_Sheet1.Columns.Get(3).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(4).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(4).Label = "相手先名"
        Me.fbResult_Sheet1.Columns.Get(4).Width = 181.0!
        Me.fbResult_Sheet1.Columns.Get(5).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(5).Label = "予約番号"
        Me.fbResult_Sheet1.Columns.Get(5).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(6).Label = "催事名"
        Me.fbResult_Sheet1.Columns.Get(6).Width = 207.0!
        Me.fbResult_Sheet1.Columns.Get(7).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(7).Label = "入金予定日"
        Me.fbResult_Sheet1.Columns.Get(7).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(8).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(8).Label = "入金日"
        Me.fbResult_Sheet1.Columns.Get(8).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(9).Label = "入金額"
        Me.fbResult_Sheet1.Columns.Get(9).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(10).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbResult_Sheet1.Columns.Get(10).Label = "入金確認"
        Me.fbResult_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbResult_Sheet1.Columns.Get(10).Width = 42.0!
        Me.fbResult_Sheet1.Columns.Get(11).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(11).Label = "請求日"
        Me.fbResult_Sheet1.Columns.Get(11).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(12).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(12).Label = "入力日"
        Me.fbResult_Sheet1.Columns.Get(12).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(13).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(13).Label = "EXAS請求NO"
        Me.fbResult_Sheet1.Columns.Get(13).Width = 95.0!
        Me.fbResult_Sheet1.Columns.Get(14).AllowAutoSort = True
        Me.fbResult_Sheet1.Columns.Get(14).Label = "請求依頼番号（入力摘要）"
        Me.fbResult_Sheet1.Columns.Get(14).Width = 93.0!
        Me.fbResult_Sheet1.Columns.Get(15).Label = "モニタNO"
        Me.fbResult_Sheet1.Columns.Get(15).Width = 68.0!
        Me.fbResult_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbResult_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.Location = New System.Drawing.Point(21, 636)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(512, 21)
        Me.Label8.TabIndex = 137
        Me.Label8.Text = "※表のデータは、コピーしたい部分を選択し、キーボードの「Ctrl」＋「C」を押すとコピーできます。"
        '
        'EXTZ0208
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1670, 705)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.fbResult)
        Me.Controls.Add(Me.btnKakutei)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label65)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0208"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　EXAS入金一覧"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.fbResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbResult_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtSeikyuIraiNo As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtSeikyuNo As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtAiteNm As System.Windows.Forms.TextBox
    Friend WithEvents txtAiteCd As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnKakutei As System.Windows.Forms.Button
    Friend WithEvents fbResult As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbResult_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents dtpNyukinTo As Common.DateTimePickerEx
    Friend WithEvents dtpNyukinFrom As Common.DateTimePickerEx
    Friend WithEvents dtpNyukinYoteiTo As Common.DateTimePickerEx
    Friend WithEvents dtpNyukinYoteiFrom As Common.DateTimePickerEx
    Friend WithEvents Label8 As System.Windows.Forms.Label

End Class
