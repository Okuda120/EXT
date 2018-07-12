<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTM0103
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
        Dim cultureInfo As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("ja-JP", False)
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTM0103))
        Me.BtnBack = New System.Windows.Forms.Button()
        Me.BtnReg = New System.Windows.Forms.Button()
        Me.ppVwList = New FarPoint.Win.Spread.FpSpread()
        Me.ppVwList_Sheet1 = New FarPoint.Win.Spread.SheetView()
        CType(Me.ppVwList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ppVwList_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnBack
        '
        Me.BtnBack.Location = New System.Drawing.Point(59, 319)
        Me.BtnBack.Name = "BtnBack"
        Me.BtnBack.Size = New System.Drawing.Size(114, 29)
        Me.BtnBack.TabIndex = 114
        Me.BtnBack.Text = "戻る"
        Me.BtnBack.UseVisualStyleBackColor = True
        '
        'BtnReg
        '
        Me.BtnReg.Location = New System.Drawing.Point(226, 319)
        Me.BtnReg.Name = "BtnReg"
        Me.BtnReg.Size = New System.Drawing.Size(114, 29)
        Me.BtnReg.TabIndex = 113
        Me.BtnReg.Text = "登録"
        Me.BtnReg.UseVisualStyleBackColor = True
        '
        'ppVwList
        '
        Me.ppVwList.AccessibleDescription = "ppVwList, Sheet1, Row 0, Column 0, 2000/4/1"
        Me.ppVwList.Location = New System.Drawing.Point(26, 27)
        Me.ppVwList.Name = "ppVwList"
        Me.ppVwList.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.ppVwList_Sheet1})
        Me.ppVwList.Size = New System.Drawing.Size(331, 255)
        Me.ppVwList.TabIndex = 115
        '
        'ppVwList_Sheet1
        '
        Me.ppVwList_Sheet1.Reset()
        Me.ppVwList_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.ppVwList_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.ppVwList_Sheet1.ColumnCount = 8
        Me.ppVwList_Sheet1.Cells.Get(0, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(0, 0).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.ppVwList_Sheet1.Cells.Get(0, 0).ParseFormatString = "yyyy/M/d"
        Me.ppVwList_Sheet1.Cells.Get(0, 0).Value = New Date(2000, 4, 1, 0, 0, 0, 0)
        Me.ppVwList_Sheet1.Cells.Get(0, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(0, 1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(0, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.ppVwList_Sheet1.Cells.Get(0, 1).ParseFormatString = "yyyy/M/d"
        Me.ppVwList_Sheet1.Cells.Get(0, 1).Value = New Date(2016, 3, 31, 0, 0, 0, 0)
        Me.ppVwList_Sheet1.Cells.Get(0, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(0, 2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(0, 2).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.ppVwList_Sheet1.Cells.Get(0, 2).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.ppVwList_Sheet1.Cells.Get(0, 2).ParseFormatString = "n"
        Me.ppVwList_Sheet1.Cells.Get(0, 2).Value = 8
        Me.ppVwList_Sheet1.Cells.Get(0, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(1, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(1, 0).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.ppVwList_Sheet1.Cells.Get(1, 0).ParseFormatString = "yyyy/M/d"
        Me.ppVwList_Sheet1.Cells.Get(1, 0).Value = New Date(2016, 4, 1, 0, 0, 0, 0)
        Me.ppVwList_Sheet1.Cells.Get(1, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(1, 1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(1, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(1, 2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Cells.Get(1, 2).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.ppVwList_Sheet1.Cells.Get(1, 2).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.ppVwList_Sheet1.Cells.Get(1, 2).ParseFormatString = "n"
        Me.ppVwList_Sheet1.Cells.Get(1, 2).Value = 10
        Me.ppVwList_Sheet1.Cells.Get(1, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "開始日"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "終了日"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "消費税率(%)"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "SEQ"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "更新区分"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "（修正前）開始日"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "（修正前）終了日"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "（修正前）消費税率（%）"
        Me.ppVwList_Sheet1.ColumnHeader.Rows.Get(0).Height = 32.0!
        Me.ppVwList_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(0).Label = "開始日"
        Me.ppVwList_Sheet1.Columns.Get(0).Width = 91.0!
        Me.ppVwList_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(1).Label = "終了日"
        Me.ppVwList_Sheet1.Columns.Get(1).Width = 91.0!
        Me.ppVwList_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(2).Label = "消費税率(%)"
        Me.ppVwList_Sheet1.Columns.Get(2).Width = 91.0!
        Me.ppVwList_Sheet1.Columns.Get(3).Label = "SEQ"
        Me.ppVwList_Sheet1.Columns.Get(3).Visible = False
        Me.ppVwList_Sheet1.Columns.Get(4).Label = "更新区分"
        Me.ppVwList_Sheet1.Columns.Get(4).Visible = False
        Me.ppVwList_Sheet1.Columns.Get(5).Label = "（修正前）開始日"
        Me.ppVwList_Sheet1.Columns.Get(5).Visible = False
        Me.ppVwList_Sheet1.Columns.Get(6).Label = "（修正前）終了日"
        Me.ppVwList_Sheet1.Columns.Get(6).Visible = False
        Me.ppVwList_Sheet1.Columns.Get(7).Label = "（修正前）消費税率（%）"
        Me.ppVwList_Sheet1.Columns.Get(7).Visible = False
        Me.ppVwList_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.ppVwList_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTM0103
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTM.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(405, 378)
        Me.Controls.Add(Me.ppVwList)
        Me.Controls.Add(Me.BtnBack)
        Me.Controls.Add(Me.BtnReg)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTM0103"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　消費税マスタメンテ"
        CType(Me.ppVwList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ppVwList_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnBack As System.Windows.Forms.Button
    Friend WithEvents BtnReg As System.Windows.Forms.Button
    Friend WithEvents ppVwList As FarPoint.Win.Spread.FpSpread
    Friend WithEvents ppVwList_Sheet1 As FarPoint.Win.Spread.SheetView
End Class
