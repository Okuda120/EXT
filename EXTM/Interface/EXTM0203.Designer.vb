<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTM0203
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTM0203))
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnConfirm = New System.Windows.Forms.Button()
        Me.grpSearch = New System.Windows.Forms.GroupBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.txtExasName = New System.Windows.Forms.TextBox()
        Me.txtExasCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.vwResult = New FarPoint.Win.Spread.FpSpread()
        Me.vwResult_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.grpSearch.SuspendLayout()
        CType(Me.vwResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwResult_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(273, 830)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(124, 35)
        Me.btnBack.TabIndex = 132
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnConfirm
        '
        Me.btnConfirm.Location = New System.Drawing.Point(618, 830)
        Me.btnConfirm.Name = "btnConfirm"
        Me.btnConfirm.Size = New System.Drawing.Size(124, 35)
        Me.btnConfirm.TabIndex = 131
        Me.btnConfirm.Text = "選択確定"
        Me.btnConfirm.UseVisualStyleBackColor = True
        '
        'grpSearch
        '
        Me.grpSearch.BackColor = System.Drawing.Color.Transparent
        Me.grpSearch.Controls.Add(Me.btnSearch)
        Me.grpSearch.Controls.Add(Me.txtExasName)
        Me.grpSearch.Controls.Add(Me.txtExasCode)
        Me.grpSearch.Controls.Add(Me.Label1)
        Me.grpSearch.Controls.Add(Me.Label32)
        Me.grpSearch.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.grpSearch.Location = New System.Drawing.Point(22, 12)
        Me.grpSearch.Name = "grpSearch"
        Me.grpSearch.Size = New System.Drawing.Size(766, 119)
        Me.grpSearch.TabIndex = 130
        Me.grpSearch.TabStop = False
        Me.grpSearch.Text = "検索条件"
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(596, 67)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(124, 29)
        Me.btnSearch.TabIndex = 73
        Me.btnSearch.Text = "検索"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'txtExasName
        '
        Me.txtExasName.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtExasName.Location = New System.Drawing.Point(163, 72)
        Me.txtExasName.Name = "txtExasName"
        Me.txtExasName.Size = New System.Drawing.Size(377, 20)
        Me.txtExasName.TabIndex = 71
        '
        'txtExasCode
        '
        Me.txtExasCode.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtExasCode.Location = New System.Drawing.Point(163, 35)
        Me.txtExasCode.Name = "txtExasCode"
        Me.txtExasCode.Size = New System.Drawing.Size(127, 20)
        Me.txtExasCode.TabIndex = 70
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(31, 75)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 69
        Me.Label1.Text = "EXAS相手先名"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label32.Location = New System.Drawing.Point(31, 38)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(106, 13)
        Me.Label32.TabIndex = 63
        Me.Label32.Text = "EXAS相手先コード"
        '
        'vwResult
        '
        Me.vwResult.AccessibleDescription = "FpSpread1, Sheet1, Row 0, Column 0, "
        Me.vwResult.Location = New System.Drawing.Point(22, 148)
        Me.vwResult.Name = "vwResult"
        Me.vwResult.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwResult_Sheet1})
        Me.vwResult.Size = New System.Drawing.Size(992, 634)
        Me.vwResult.TabIndex = 133
        '
        'vwResult_Sheet1
        '
        Me.vwResult_Sheet1.Reset()
        Me.vwResult_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwResult_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwResult_Sheet1.ColumnCount = 9
        Me.vwResult_Sheet1.Cells.Get(0, 8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwResult_Sheet1.Cells.Get(0, 8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "相手先コード"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "相手先名"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "相手先名カナ"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "郵便番号"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "住所１＋２＋３"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "電話番号"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "FAX"
        Me.vwResult_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "使用停止"
        Me.vwResult_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.vwResult_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(0).Label = "選択"
        Me.vwResult_Sheet1.Columns.Get(0).Locked = False
        Me.vwResult_Sheet1.Columns.Get(0).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(0).Width = 39.0!
        Me.vwResult_Sheet1.Columns.Get(1).Label = "相手先コード"
        Me.vwResult_Sheet1.Columns.Get(1).Locked = True
        Me.vwResult_Sheet1.Columns.Get(1).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(1).Width = 84.0!
        Me.vwResult_Sheet1.Columns.Get(2).Label = "相手先名"
        Me.vwResult_Sheet1.Columns.Get(2).Locked = True
        Me.vwResult_Sheet1.Columns.Get(2).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(2).Width = 148.0!
        Me.vwResult_Sheet1.Columns.Get(3).Label = "相手先名カナ"
        Me.vwResult_Sheet1.Columns.Get(3).Locked = True
        Me.vwResult_Sheet1.Columns.Get(3).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(3).Width = 148.0!
        Me.vwResult_Sheet1.Columns.Get(4).Label = "郵便番号"
        Me.vwResult_Sheet1.Columns.Get(4).Locked = True
        Me.vwResult_Sheet1.Columns.Get(4).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(5).Label = "住所１＋２＋３"
        Me.vwResult_Sheet1.Columns.Get(5).Locked = True
        Me.vwResult_Sheet1.Columns.Get(5).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(5).Width = 229.0!
        Me.vwResult_Sheet1.Columns.Get(6).Label = "電話番号"
        Me.vwResult_Sheet1.Columns.Get(6).Locked = True
        Me.vwResult_Sheet1.Columns.Get(6).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(6).Width = 80.0!
        Me.vwResult_Sheet1.Columns.Get(7).Label = "FAX"
        Me.vwResult_Sheet1.Columns.Get(7).Locked = True
        Me.vwResult_Sheet1.Columns.Get(7).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(7).Width = 80.0!
        Me.vwResult_Sheet1.Columns.Get(8).Label = "使用停止"
        Me.vwResult_Sheet1.Columns.Get(8).Locked = True
        Me.vwResult_Sheet1.Columns.Get(8).Resizable = False
        Me.vwResult_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwResult_Sheet1.Columns.Get(8).Width = 68.0!
        Me.vwResult_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwResult_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTM0203
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTM.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1087, 892)
        Me.Controls.Add(Me.vwResult)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnConfirm)
        Me.Controls.Add(Me.grpSearch)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTM0203"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　EXAS相手先一覧"
        Me.grpSearch.ResumeLayout(False)
        Me.grpSearch.PerformLayout()
        CType(Me.vwResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwResult_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnConfirm As System.Windows.Forms.Button
    Friend WithEvents grpSearch As System.Windows.Forms.GroupBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents txtExasName As System.Windows.Forms.TextBox
    Friend WithEvents txtExasCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents vwResult As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwResult_Sheet1 As FarPoint.Win.Spread.SheetView
End Class
