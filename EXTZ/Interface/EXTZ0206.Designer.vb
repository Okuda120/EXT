<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0206
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0206))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblSagaku = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.btnComplate = New System.Windows.Forms.Button()
        Me.txtWariai1 = New System.Windows.Forms.TextBox()
        Me.txtWariai2 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnCalc = New System.Windows.Forms.Button()
        Me.fbFrom = New FarPoint.Win.Spread.FpSpread()
        Me.fbFrom_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.fbTo = New FarPoint.Win.Spread.FpSpread()
        Me.fbTo_Sheet1 = New FarPoint.Win.Spread.SheetView()
        CType(Me.fbFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbFrom_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbTo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbTo_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.BackColor = System.Drawing.Color.Transparent
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(12, 18)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(150, 21)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "分割請求とする請求書"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(1222, 857)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(175, 36)
        Me.Button5.TabIndex = 22
        Me.Button5.Text = "戻る"
        Me.Button5.UseVisualStyleBackColor = True
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
        Me.btnBack.Location = New System.Drawing.Point(184, 349)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 89
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.Location = New System.Drawing.Point(66, 166)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(27, 21)
        Me.Label10.TabIndex = 104
        Me.Label10.Text = "："
        '
        'lblSagaku
        '
        Me.lblSagaku.BackColor = System.Drawing.Color.Silver
        Me.lblSagaku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSagaku.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSagaku.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSagaku.Location = New System.Drawing.Point(287, 287)
        Me.lblSagaku.Margin = New System.Windows.Forms.Padding(3)
        Me.lblSagaku.Name = "lblSagaku"
        Me.lblSagaku.Size = New System.Drawing.Size(118, 20)
        Me.lblSagaku.TabIndex = 110
        Me.lblSagaku.Text = "0"
        Me.lblSagaku.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label15.Location = New System.Drawing.Point(206, 291)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(75, 21)
        Me.Label15.TabIndex = 109
        Me.Label15.Text = "分割差額"
        '
        'btnComplate
        '
        Me.btnComplate.Location = New System.Drawing.Point(361, 349)
        Me.btnComplate.Name = "btnComplate"
        Me.btnComplate.Size = New System.Drawing.Size(140, 23)
        Me.btnComplate.TabIndex = 86
        Me.btnComplate.Text = "入力完了"
        Me.btnComplate.UseVisualStyleBackColor = True
        '
        'txtWariai1
        '
        Me.txtWariai1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtWariai1.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.txtWariai1.Location = New System.Drawing.Point(15, 163)
        Me.txtWariai1.MaxLength = 1
        Me.txtWariai1.Name = "txtWariai1"
        Me.txtWariai1.Size = New System.Drawing.Size(45, 20)
        Me.txtWariai1.TabIndex = 113
        '
        'txtWariai2
        '
        Me.txtWariai2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtWariai2.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.txtWariai2.Location = New System.Drawing.Point(84, 163)
        Me.txtWariai2.MaxLength = 1
        Me.txtWariai2.Name = "txtWariai2"
        Me.txtWariai2.Size = New System.Drawing.Size(45, 20)
        Me.txtWariai2.TabIndex = 114
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(144, 166)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(150, 21)
        Me.Label1.TabIndex = 115
        Me.Label1.Text = "で分割する"
        '
        'btnCalc
        '
        Me.btnCalc.Location = New System.Drawing.Point(223, 161)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(96, 23)
        Me.btnCalc.TabIndex = 116
        Me.btnCalc.Text = "再表示"
        Me.btnCalc.UseVisualStyleBackColor = True
        '
        'fbFrom
        '
        Me.fbFrom.AccessibleDescription = "FpSpread1, Sheet1, Row 0, Column 0, 株式会社イベント企画"
        Me.fbFrom.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbFrom.Location = New System.Drawing.Point(15, 52)
        Me.fbFrom.Name = "fbFrom"
        Me.fbFrom.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbFrom_Sheet1})
        Me.fbFrom.Size = New System.Drawing.Size(567, 43)
        Me.fbFrom.TabIndex = 117
        Me.fbFrom.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        '
        'fbFrom_Sheet1
        '
        Me.fbFrom_Sheet1.Reset()
        Me.fbFrom_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbFrom_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbFrom_Sheet1.ColumnCount = 3
        Me.fbFrom_Sheet1.RowCount = 0
        Me.fbFrom_Sheet1.ActiveColumnIndex = -1
        Me.fbFrom_Sheet1.ActiveRowIndex = -1
        Me.fbFrom_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "相手先名"
        Me.fbFrom_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "確定額"
        Me.fbFrom_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "請求内容"
        Me.fbFrom_Sheet1.Columns.Get(0).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFrom_Sheet1.Columns.Get(0).Label = "相手先名"
        Me.fbFrom_Sheet1.Columns.Get(0).Width = 226.0!
        Me.fbFrom_Sheet1.Columns.Get(1).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFrom_Sheet1.Columns.Get(1).Label = "確定額"
        Me.fbFrom_Sheet1.Columns.Get(1).Width = 116.0!
        Me.fbFrom_Sheet1.Columns.Get(2).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFrom_Sheet1.Columns.Get(2).Label = "請求内容"
        Me.fbFrom_Sheet1.Columns.Get(2).Width = 176.0!
        Me.fbFrom_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbFrom_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'fbTo
        '
        Me.fbTo.AccessibleDescription = "FpSpread2, Sheet1, Row 0, Column 0, 株式会社イベント企画"
        Me.fbTo.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbTo.Location = New System.Drawing.Point(15, 199)
        Me.fbTo.Name = "fbTo"
        Me.fbTo.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbTo_Sheet1})
        Me.fbTo.Size = New System.Drawing.Size(567, 64)
        Me.fbTo.TabIndex = 118
        Me.fbTo.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        '
        'fbTo_Sheet1
        '
        Me.fbTo_Sheet1.Reset()
        Me.fbTo_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbTo_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbTo_Sheet1.ColumnCount = 3
        Me.fbTo_Sheet1.RowCount = 0
        Me.fbTo_Sheet1.ActiveColumnIndex = -1
        Me.fbTo_Sheet1.ActiveRowIndex = -1
        Me.fbTo_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "相手先名"
        Me.fbTo_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "確定額"
        Me.fbTo_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "請求内容"
        Me.fbTo_Sheet1.Columns.Get(0).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbTo_Sheet1.Columns.Get(0).Label = "相手先名"
        Me.fbTo_Sheet1.Columns.Get(0).Width = 230.0!
        Me.fbTo_Sheet1.Columns.Get(1).Label = "確定額"
        Me.fbTo_Sheet1.Columns.Get(1).Width = 120.0!
        Me.fbTo_Sheet1.Columns.Get(2).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbTo_Sheet1.Columns.Get(2).Label = "請求内容"
        Me.fbTo_Sheet1.Columns.Get(2).Width = 166.0!
        Me.fbTo_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbTo_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTZ0206
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(623, 391)
        Me.Controls.Add(Me.fbTo)
        Me.Controls.Add(Me.fbFrom)
        Me.Controls.Add(Me.btnCalc)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtWariai2)
        Me.Controls.Add(Me.txtWariai1)
        Me.Controls.Add(Me.lblSagaku)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnComplate)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label65)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0206"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　請求分割"
        CType(Me.fbFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbFrom_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbTo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbTo_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblSagaku As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents btnComplate As System.Windows.Forms.Button
    Friend WithEvents txtWariai1 As System.Windows.Forms.TextBox
    Friend WithEvents txtWariai2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCalc As System.Windows.Forms.Button
    Friend WithEvents fbFrom As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbFrom_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents fbTo As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbTo_Sheet1 As FarPoint.Win.Spread.SheetView

End Class
