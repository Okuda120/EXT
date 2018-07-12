<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0210
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0210))
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.btnComplate = New System.Windows.Forms.Button()
        Me.lblRiyobi = New System.Windows.Forms.Label()
        Me.fbCopyDate = New FarPoint.Win.Spread.FpSpread()
        Me.fbCopyDate_Sheet1 = New FarPoint.Win.Spread.SheetView()
        CType(Me.fbCopyDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbCopyDate_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.btnBack.Location = New System.Drawing.Point(22, 236)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 89
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label19.Location = New System.Drawing.Point(17, 21)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(133, 21)
        Me.Label19.TabIndex = 117
        Me.Label19.Text = "選択している利用日"
        '
        'btnComplate
        '
        Me.btnComplate.Location = New System.Drawing.Point(156, 236)
        Me.btnComplate.Name = "btnComplate"
        Me.btnComplate.Size = New System.Drawing.Size(127, 23)
        Me.btnComplate.TabIndex = 112
        Me.btnComplate.Text = "入力完了"
        Me.btnComplate.UseVisualStyleBackColor = True
        '
        'lblRiyobi
        '
        Me.lblRiyobi.BackColor = System.Drawing.Color.Silver
        Me.lblRiyobi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRiyobi.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblRiyobi.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblRiyobi.Location = New System.Drawing.Point(156, 17)
        Me.lblRiyobi.Margin = New System.Windows.Forms.Padding(3)
        Me.lblRiyobi.Name = "lblRiyobi"
        Me.lblRiyobi.Size = New System.Drawing.Size(119, 21)
        Me.lblRiyobi.TabIndex = 137
        Me.lblRiyobi.Text = "2015/5/2(土)"
        Me.lblRiyobi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'fbCopyDate
        '
        Me.fbCopyDate.AccessibleDescription = "FpSpread1, Sheet1, Row 0, Column 0, "
        Me.fbCopyDate.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbCopyDate.Location = New System.Drawing.Point(20, 56)
        Me.fbCopyDate.Name = "fbCopyDate"
        Me.fbCopyDate.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbCopyDate_Sheet1})
        Me.fbCopyDate.Size = New System.Drawing.Size(255, 162)
        Me.fbCopyDate.TabIndex = 138
        '
        'fbCopyDate_Sheet1
        '
        Me.fbCopyDate_Sheet1.Reset()
        Me.fbCopyDate_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbCopyDate_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbCopyDate_Sheet1.ColumnCount = 2
        Me.fbCopyDate_Sheet1.RowCount = 0
        Me.fbCopyDate_Sheet1.ActiveColumnIndex = -1
        Me.fbCopyDate_Sheet1.ActiveRowIndex = -1
        Me.fbCopyDate_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.fbCopyDate_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "コピー先の利用日"
        Me.fbCopyDate_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.fbCopyDate_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbCopyDate_Sheet1.Columns.Get(0).Label = "選択"
        Me.fbCopyDate_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbCopyDate_Sheet1.Columns.Get(0).Width = 37.0!
        Me.fbCopyDate_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbCopyDate_Sheet1.Columns.Get(1).Label = "コピー先の利用日"
        Me.fbCopyDate_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbCopyDate_Sheet1.Columns.Get(1).Width = 158.0!
        Me.fbCopyDate_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbCopyDate_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTZ0210
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(339, 287)
        Me.Controls.Add(Me.fbCopyDate)
        Me.Controls.Add(Me.lblRiyobi)
        Me.Controls.Add(Me.btnComplate)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label19)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0210"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　付帯設備登録コピー"
        CType(Me.fbCopyDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbCopyDate_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents btnComplate As System.Windows.Forms.Button
    Friend WithEvents lblRiyobi As System.Windows.Forms.Label
    Friend WithEvents fbCopyDate As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbCopyDate_Sheet1 As FarPoint.Win.Spread.SheetView

End Class
