<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0209
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
        Dim NumberCellType1 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0209))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnComplate = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblYoyakuNo = New System.Windows.Forms.Label()
        Me.lblSaiji = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblRiyobi = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmbBunrui = New System.Windows.Forms.ComboBox()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.fbFutai = New FarPoint.Win.Spread.FpSpread()
        Me.fbFutai_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnAddAll = New System.Windows.Forms.Button()
        Me.btnDelAll = New System.Windows.Forms.Button()
        Me.btnDel = New System.Windows.Forms.Button()
        Me.fbFutaiSelected = New FarPoint.Win.Spread.FpSpread()
        Me.fbFutaiSelected_Sheet1 = New FarPoint.Win.Spread.SheetView()
        CType(Me.fbFutai, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbFutai_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbFutaiSelected, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbFutaiSelected_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.BackColor = System.Drawing.Color.Transparent
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(12, 18)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(124, 21)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "予約番号"
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
        Me.btnBack.Location = New System.Drawing.Point(547, 688)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 89
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnComplate
        '
        Me.btnComplate.Location = New System.Drawing.Point(698, 688)
        Me.btnComplate.Name = "btnComplate"
        Me.btnComplate.Size = New System.Drawing.Size(122, 23)
        Me.btnComplate.TabIndex = 112
        Me.btnComplate.Text = "入力完了"
        Me.btnComplate.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 21)
        Me.Label1.TabIndex = 113
        Me.Label1.Text = "催事名(ｱｰﾃｨｽﾄ名)"
        '
        'lblYoyakuNo
        '
        Me.lblYoyakuNo.BackColor = System.Drawing.Color.Silver
        Me.lblYoyakuNo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblYoyakuNo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblYoyakuNo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblYoyakuNo.Location = New System.Drawing.Point(142, 14)
        Me.lblYoyakuNo.Margin = New System.Windows.Forms.Padding(3)
        Me.lblYoyakuNo.Name = "lblYoyakuNo"
        Me.lblYoyakuNo.Size = New System.Drawing.Size(88, 20)
        Me.lblYoyakuNo.TabIndex = 114
        Me.lblYoyakuNo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSaiji
        '
        Me.lblSaiji.BackColor = System.Drawing.Color.Silver
        Me.lblSaiji.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSaiji.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSaiji.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSaiji.Location = New System.Drawing.Point(142, 44)
        Me.lblSaiji.Margin = New System.Windows.Forms.Padding(3)
        Me.lblSaiji.Name = "lblSaiji"
        Me.lblSaiji.Size = New System.Drawing.Size(327, 20)
        Me.lblSaiji.TabIndex = 115
        Me.lblSaiji.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(265, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 21)
        Me.Label4.TabIndex = 116
        Me.Label4.Text = "利用日"
        '
        'lblRiyobi
        '
        Me.lblRiyobi.BackColor = System.Drawing.Color.Silver
        Me.lblRiyobi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRiyobi.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblRiyobi.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblRiyobi.Location = New System.Drawing.Point(324, 14)
        Me.lblRiyobi.Margin = New System.Windows.Forms.Padding(3)
        Me.lblRiyobi.Name = "lblRiyobi"
        Me.lblRiyobi.Size = New System.Drawing.Size(145, 20)
        Me.lblRiyobi.TabIndex = 117
        Me.lblRiyobi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(12, 94)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 21)
        Me.Label6.TabIndex = 118
        Me.Label6.Text = "分類"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.Location = New System.Drawing.Point(338, 94)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 21)
        Me.Label7.TabIndex = 119
        Me.Label7.Text = "選択済み設備"
        '
        'cmbBunrui
        '
        Me.cmbBunrui.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbBunrui.FormattingEnabled = True
        Me.cmbBunrui.Location = New System.Drawing.Point(12, 118)
        Me.cmbBunrui.Name = "cmbBunrui"
        Me.cmbBunrui.Size = New System.Drawing.Size(198, 21)
        Me.cmbBunrui.TabIndex = 120
        '
        'lblTotal
        '
        Me.lblTotal.BackColor = System.Drawing.Color.Silver
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblTotal.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTotal.Location = New System.Drawing.Point(995, 606)
        Me.lblTotal.Margin = New System.Windows.Forms.Padding(3)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(110, 20)
        Me.lblTotal.TabIndex = 122
        Me.lblTotal.Text = "36,000"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label9.Location = New System.Drawing.Point(865, 610)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(124, 21)
        Me.Label9.TabIndex = 121
        Me.Label9.Text = "付帯設備合計"
        '
        'fbFutai
        '
        Me.fbFutai.AccessibleDescription = "fbFutai, Sheet1, Row 0, Column 0, マイク（追加分）"
        Me.fbFutai.Location = New System.Drawing.Point(15, 160)
        Me.fbFutai.Name = "fbFutai"
        Me.fbFutai.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbFutai_Sheet1})
        Me.fbFutai.Size = New System.Drawing.Size(204, 440)
        Me.fbFutai.TabIndex = 123
        '
        'fbFutai_Sheet1
        '
        Me.fbFutai_Sheet1.Reset()
        Me.fbFutai_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbFutai_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbFutai_Sheet1.ColumnCount = 6
        Me.fbFutai_Sheet1.RowCount = 0
        Me.fbFutai_Sheet1.ActiveColumnIndex = -1
        Me.fbFutai_Sheet1.ActiveRowIndex = -1
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "設備名"
        Me.fbFutai_Sheet1.Columns.Get(0).Label = "設備名"
        Me.fbFutai_Sheet1.Columns.Get(0).Width = 153.0!
        Me.fbFutai_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbFutai_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(253, 203)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(41, 32)
        Me.btnAdd.TabIndex = 124
        Me.btnAdd.Text = "＞"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnAddAll
        '
        Me.btnAddAll.Location = New System.Drawing.Point(253, 241)
        Me.btnAddAll.Name = "btnAddAll"
        Me.btnAddAll.Size = New System.Drawing.Size(41, 32)
        Me.btnAddAll.TabIndex = 125
        Me.btnAddAll.Text = "＞＞"
        Me.btnAddAll.UseVisualStyleBackColor = True
        '
        'btnDelAll
        '
        Me.btnDelAll.Location = New System.Drawing.Point(253, 367)
        Me.btnDelAll.Name = "btnDelAll"
        Me.btnDelAll.Size = New System.Drawing.Size(41, 32)
        Me.btnDelAll.TabIndex = 127
        Me.btnDelAll.Text = "＜＜"
        Me.btnDelAll.UseVisualStyleBackColor = True
        '
        'btnDel
        '
        Me.btnDel.Location = New System.Drawing.Point(253, 329)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(41, 32)
        Me.btnDel.TabIndex = 126
        Me.btnDel.Text = "＜"
        Me.btnDel.UseVisualStyleBackColor = True
        '
        'fbFutaiSelected
        '
        Me.fbFutaiSelected.AccessibleDescription = "FpSpread2, Sheet1, Row 0, Column 0, 音響"
        Me.fbFutaiSelected.Location = New System.Drawing.Point(341, 118)
        Me.fbFutaiSelected.Name = "fbFutaiSelected"
        Me.fbFutaiSelected.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbFutaiSelected_Sheet1})
        Me.fbFutaiSelected.Size = New System.Drawing.Size(971, 482)
        Me.fbFutaiSelected.TabIndex = 128
        '
        'fbFutaiSelected_Sheet1
        '
        Me.fbFutaiSelected_Sheet1.Reset()
        Me.fbFutaiSelected_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbFutaiSelected_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbFutaiSelected_Sheet1.ColumnCount = 10
        Me.fbFutaiSelected_Sheet1.RowCount = 0
        Me.fbFutaiSelected_Sheet1.ActiveColumnIndex = -1
        Me.fbFutaiSelected_Sheet1.ActiveRowIndex = -1
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "分類"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "設備名"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "単価"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "数量"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "単位"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "小計"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "調整額"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "合計"
        Me.fbFutaiSelected_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "備考"
        Me.fbFutaiSelected_Sheet1.Columns.Get(0).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutaiSelected_Sheet1.Columns.Get(0).ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.fbFutaiSelected_Sheet1.Columns.Get(0).Label = "分類"
        Me.fbFutaiSelected_Sheet1.Columns.Get(0).Locked = False
        Me.fbFutaiSelected_Sheet1.Columns.Get(0).Width = 103.0!
        Me.fbFutaiSelected_Sheet1.Columns.Get(1).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutaiSelected_Sheet1.Columns.Get(1).ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.fbFutaiSelected_Sheet1.Columns.Get(1).Label = "設備名"
        Me.fbFutaiSelected_Sheet1.Columns.Get(1).Width = 184.0!
        Me.fbFutaiSelected_Sheet1.Columns.Get(2).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutaiSelected_Sheet1.Columns.Get(2).Label = "単価"
        Me.fbFutaiSelected_Sheet1.Columns.Get(4).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutaiSelected_Sheet1.Columns.Get(4).Label = "単位"
        Me.fbFutaiSelected_Sheet1.Columns.Get(5).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        NumberCellType1.AllowEditorVerticalAlign = True
        NumberCellType1.DecimalPlaces = 0
        NumberCellType1.FixedPoint = False
        NumberCellType1.MaximumValue = 99999999999.0R
        NumberCellType1.MinimumValue = -99999999999.0R
        NumberCellType1.ShowSeparator = True
        Me.fbFutaiSelected_Sheet1.Columns.Get(5).CellType = NumberCellType1
        Me.fbFutaiSelected_Sheet1.Columns.Get(5).Label = "小計"
        Me.fbFutaiSelected_Sheet1.Columns.Get(5).Width = 76.0!
        Me.fbFutaiSelected_Sheet1.Columns.Get(6).Label = "調整額"
        Me.fbFutaiSelected_Sheet1.Columns.Get(6).Width = 76.0!
        Me.fbFutaiSelected_Sheet1.Columns.Get(7).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutaiSelected_Sheet1.Columns.Get(7).Label = "合計"
        Me.fbFutaiSelected_Sheet1.Columns.Get(7).Width = 76.0!
        Me.fbFutaiSelected_Sheet1.Columns.Get(8).Label = "備考"
        Me.fbFutaiSelected_Sheet1.Columns.Get(8).Width = 220.0!
        Me.fbFutaiSelected_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbFutaiSelected_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTZ0209
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1438, 736)
        Me.Controls.Add(Me.fbFutaiSelected)
        Me.Controls.Add(Me.btnDelAll)
        Me.Controls.Add(Me.btnDel)
        Me.Controls.Add(Me.btnAddAll)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.fbFutai)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.cmbBunrui)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblRiyobi)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblSaiji)
        Me.Controls.Add(Me.lblYoyakuNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnComplate)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label65)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0209"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　付帯設備登録／詳細"
        CType(Me.fbFutai, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbFutai_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbFutaiSelected, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbFutaiSelected_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnComplate As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblYoyakuNo As System.Windows.Forms.Label
    Friend WithEvents lblSaiji As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblRiyobi As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmbBunrui As System.Windows.Forms.ComboBox
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents fbFutai As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbFutai_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnAddAll As System.Windows.Forms.Button
    Friend WithEvents btnDelAll As System.Windows.Forms.Button
    Friend WithEvents btnDel As System.Windows.Forms.Button
    Friend WithEvents fbFutaiSelected As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbFutaiSelected_Sheet1 As FarPoint.Win.Spread.SheetView

End Class
