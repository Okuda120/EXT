<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTY0101
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
        Dim CheckBoxCellType2 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTY0101))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnDisplay = New System.Windows.Forms.Button()
        Me.rdoStudio = New System.Windows.Forms.RadioButton()
        Me.rdoTheater = New System.Windows.Forms.RadioButton()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnOutput = New System.Windows.Forms.Button()
        Me.vwBillpay = New FarPoint.Win.Spread.FpSpread()
        Me.vwBillpay_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwBillpay_Sheet2 = New FarPoint.Win.Spread.SheetView()
        Me.GroupBox1.SuspendLayout()
        CType(Me.vwBillpay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwBillpay_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwBillpay_Sheet2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.btnDisplay)
        Me.GroupBox1.Controls.Add(Me.rdoStudio)
        Me.GroupBox1.Controls.Add(Me.rdoTheater)
        Me.GroupBox1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(34, 27)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(473, 64)
        Me.GroupBox1.TabIndex = 127
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "検索対象"
        '
        'btnDisplay
        '
        Me.btnDisplay.Location = New System.Drawing.Point(321, 19)
        Me.btnDisplay.Name = "btnDisplay"
        Me.btnDisplay.Size = New System.Drawing.Size(124, 29)
        Me.btnDisplay.TabIndex = 75
        Me.btnDisplay.Text = "表示"
        Me.btnDisplay.UseVisualStyleBackColor = True
        '
        'rdoStudio
        '
        Me.rdoStudio.AutoSize = True
        Me.rdoStudio.Location = New System.Drawing.Point(147, 25)
        Me.rdoStudio.Name = "rdoStudio"
        Me.rdoStudio.Size = New System.Drawing.Size(65, 17)
        Me.rdoStudio.TabIndex = 32
        Me.rdoStudio.TabStop = True
        Me.rdoStudio.Text = "スタジオ"
        Me.rdoStudio.UseVisualStyleBackColor = True
        '
        'rdoTheater
        '
        Me.rdoTheater.AutoSize = True
        Me.rdoTheater.Location = New System.Drawing.Point(40, 25)
        Me.rdoTheater.Name = "rdoTheater"
        Me.rdoTheater.Size = New System.Drawing.Size(66, 17)
        Me.rdoTheater.TabIndex = 31
        Me.rdoTheater.TabStop = True
        Me.rdoTheater.Text = "シアター"
        Me.rdoTheater.UseVisualStyleBackColor = True
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.BackColor = System.Drawing.Color.Transparent
        Me.Label32.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label32.Location = New System.Drawing.Point(43, 130)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(85, 13)
        Me.Label32.TabIndex = 128
        Me.Label32.Text = "作成対象請求"
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(475, 850)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(124, 35)
        Me.btnBack.TabIndex = 129
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnOutput
        '
        Me.btnOutput.Location = New System.Drawing.Point(896, 850)
        Me.btnOutput.Name = "btnOutput"
        Me.btnOutput.Size = New System.Drawing.Size(162, 35)
        Me.btnOutput.TabIndex = 130
        Me.btnOutput.Text = "請求依頼データ出力"
        Me.btnOutput.UseVisualStyleBackColor = True
        '
        'vwBillpay
        '
        Me.vwBillpay.AccessibleDescription = "vwBillpay, Sheet1, Row 0, Column 0, "
        Me.vwBillpay.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwBillpay.Location = New System.Drawing.Point(34, 161)
        Me.vwBillpay.Name = "vwBillpay"
        Me.vwBillpay.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwBillpay_Sheet1, Me.vwBillpay_Sheet2})
        Me.vwBillpay.Size = New System.Drawing.Size(1518, 644)
        Me.vwBillpay.TabIndex = 131
        Me.vwBillpay.TabStripPolicy = FarPoint.Win.Spread.TabStripPolicy.Never
        '
        'vwBillpay_Sheet1
        '
        Me.vwBillpay_Sheet1.Reset()
        Me.vwBillpay_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwBillpay_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwBillpay_Sheet1.ColumnCount = 13
        Me.vwBillpay_Sheet1.RowCount = 30
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "請求依頼番号"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "請求日"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用開始日"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "利用終了日"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "催事名"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "相手先名"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "請求金額"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "請求内容"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "入金予定日"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "予約番号"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "SEQ"
        Me.vwBillpay_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "相手先コード"
        Me.vwBillpay_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.vwBillpay_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(0).Label = "選択"
        Me.vwBillpay_Sheet1.Columns.Get(0).Width = 35.0!
        Me.vwBillpay_Sheet1.Columns.Get(1).Label = "請求依頼番号"
        Me.vwBillpay_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(1).Width = 104.0!
        Me.vwBillpay_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(2).Label = "請求日"
        Me.vwBillpay_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(2).Width = 117.0!
        Me.vwBillpay_Sheet1.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(3).Label = "利用開始日"
        Me.vwBillpay_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(3).Width = 103.0!
        Me.vwBillpay_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(4).Label = "利用終了日"
        Me.vwBillpay_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(4).Width = 103.0!
        Me.vwBillpay_Sheet1.Columns.Get(5).Label = "催事名"
        Me.vwBillpay_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(5).Width = 261.0!
        Me.vwBillpay_Sheet1.Columns.Get(6).Label = "相手先名"
        Me.vwBillpay_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(6).Width = 193.0!
        Me.vwBillpay_Sheet1.Columns.Get(7).Label = "請求金額"
        Me.vwBillpay_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(7).Width = 104.0!
        Me.vwBillpay_Sheet1.Columns.Get(8).Label = "請求内容"
        Me.vwBillpay_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(8).Width = 158.0!
        Me.vwBillpay_Sheet1.Columns.Get(9).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(9).Label = "入金予定日"
        Me.vwBillpay_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(9).Width = 107.0!
        Me.vwBillpay_Sheet1.Columns.Get(10).Label = "予約番号"
        Me.vwBillpay_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(10).Width = 90.0!
        Me.vwBillpay_Sheet1.Columns.Get(11).Label = "SEQ"
        Me.vwBillpay_Sheet1.Columns.Get(11).Visible = False
        Me.vwBillpay_Sheet1.Columns.Get(12).Label = "相手先コード"
        Me.vwBillpay_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet1.Columns.Get(12).Visible = False
        Me.vwBillpay_Sheet1.Columns.Get(12).Width = 90.0!
        Me.vwBillpay_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwBillpay_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwBillpay_Sheet2
        '
        Me.vwBillpay_Sheet2.Reset()
        Me.vwBillpay_Sheet2.SheetName = "Sheet2"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwBillpay_Sheet2.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwBillpay_Sheet2.ColumnCount = 14
        Me.vwBillpay_Sheet2.RowCount = 30
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 1).Value = "請求依頼番号"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 2).Value = "請求日"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 3).Value = "利用開始日"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 4).Value = "利用終了日"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 5).Value = "スタジオ"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 6).Value = "アーティスト名"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 7).Value = "相手先名"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 8).Value = "請求金額"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 9).Value = "請求内容"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 10).Value = "入金予定日"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 11).Value = "予約番号"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 12).Value = "SEQ"
        Me.vwBillpay_Sheet2.ColumnHeader.Cells.Get(0, 13).Value = "相手先コード"
        Me.vwBillpay_Sheet2.Columns.Get(0).CellType = CheckBoxCellType2
        Me.vwBillpay_Sheet2.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(0).Label = "選択"
        Me.vwBillpay_Sheet2.Columns.Get(0).Width = 35.0!
        Me.vwBillpay_Sheet2.Columns.Get(1).Label = "請求依頼番号"
        Me.vwBillpay_Sheet2.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(1).Width = 103.0!
        Me.vwBillpay_Sheet2.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(2).Label = "請求日"
        Me.vwBillpay_Sheet2.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(2).Width = 117.0!
        Me.vwBillpay_Sheet2.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(3).Label = "利用開始日"
        Me.vwBillpay_Sheet2.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(3).Width = 103.0!
        Me.vwBillpay_Sheet2.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(4).Label = "利用終了日"
        Me.vwBillpay_Sheet2.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(4).Width = 103.0!
        Me.vwBillpay_Sheet2.Columns.Get(5).Label = "スタジオ"
        Me.vwBillpay_Sheet2.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(5).Width = 88.0!
        Me.vwBillpay_Sheet2.Columns.Get(6).Label = "アーティスト名"
        Me.vwBillpay_Sheet2.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(6).Width = 261.0!
        Me.vwBillpay_Sheet2.Columns.Get(7).Label = "相手先名"
        Me.vwBillpay_Sheet2.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(7).Width = 193.0!
        Me.vwBillpay_Sheet2.Columns.Get(8).Label = "請求金額"
        Me.vwBillpay_Sheet2.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(8).Width = 104.0!
        Me.vwBillpay_Sheet2.Columns.Get(9).Label = "請求内容"
        Me.vwBillpay_Sheet2.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(9).Width = 158.0!
        Me.vwBillpay_Sheet2.Columns.Get(10).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(10).Label = "入金予定日"
        Me.vwBillpay_Sheet2.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwBillpay_Sheet2.Columns.Get(10).Width = 107.0!
        Me.vwBillpay_Sheet2.Columns.Get(11).Label = "予約番号"
        Me.vwBillpay_Sheet2.Columns.Get(11).Locked = True
        Me.vwBillpay_Sheet2.Columns.Get(11).Width = 158.0!
        Me.vwBillpay_Sheet2.Columns.Get(12).Label = "SEQ"
        Me.vwBillpay_Sheet2.Columns.Get(12).Visible = False
        Me.vwBillpay_Sheet2.Columns.Get(13).Label = "相手先コード"
        Me.vwBillpay_Sheet2.Columns.Get(13).Visible = False
        Me.vwBillpay_Sheet2.Columns.Get(13).Width = 77.0!
        Me.vwBillpay_Sheet2.RowHeader.Columns.Default.Resizable = False
        Me.vwBillpay_Sheet2.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTY0101
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTY.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1584, 913)
        Me.Controls.Add(Me.vwBillpay)
        Me.Controls.Add(Me.btnOutput)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.GroupBox1)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTY0101"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　EXAS請求依頼データ作成"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.vwBillpay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwBillpay_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwBillpay_Sheet2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoStudio As System.Windows.Forms.RadioButton
    Friend WithEvents rdoTheater As System.Windows.Forms.RadioButton
    Friend WithEvents btnDisplay As System.Windows.Forms.Button
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnOutput As System.Windows.Forms.Button
    Friend WithEvents vwBillpay As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwBillpay_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwBillpay_Sheet2 As FarPoint.Win.Spread.SheetView
End Class
