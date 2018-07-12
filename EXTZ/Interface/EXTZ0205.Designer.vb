<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0205
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0205))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnKakutei = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.txtUchiNm = New System.Windows.Forms.TextBox()
        Me.txtPrjNm = New System.Windows.Forms.TextBox()
        Me.txtUchiCd = New System.Windows.Forms.TextBox()
        Me.txtPrjCd = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.fbResult = New FarPoint.Win.Spread.FpSpread()
        Me.fbResult_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.GroupBox1.SuspendLayout()
        CType(Me.fbResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbResult_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(6, 16)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(111, 19)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "プロジェクトコード"
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(182, 451)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 86
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
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
        'btnKakutei
        '
        Me.btnKakutei.Location = New System.Drawing.Point(421, 451)
        Me.btnKakutei.Name = "btnKakutei"
        Me.btnKakutei.Size = New System.Drawing.Size(96, 23)
        Me.btnKakutei.TabIndex = 89
        Me.btnKakutei.Text = "確定"
        Me.btnKakutei.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.btnSearch)
        Me.GroupBox1.Controls.Add(Me.txtUchiNm)
        Me.GroupBox1.Controls.Add(Me.txtPrjNm)
        Me.GroupBox1.Controls.Add(Me.txtUchiCd)
        Me.GroupBox1.Controls.Add(Me.txtPrjCd)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label65)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(665, 119)
        Me.GroupBox1.TabIndex = 90
        Me.GroupBox1.TabStop = False
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(532, 81)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(96, 23)
        Me.btnSearch.TabIndex = 87
        Me.btnSearch.Text = "検索"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'txtUchiNm
        '
        Me.txtUchiNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtUchiNm.Location = New System.Drawing.Point(326, 40)
        Me.txtUchiNm.MaxLength = 30
        Me.txtUchiNm.Name = "txtUchiNm"
        Me.txtUchiNm.Size = New System.Drawing.Size(278, 20)
        Me.txtUchiNm.TabIndex = 51
        '
        'txtPrjNm
        '
        Me.txtPrjNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtPrjNm.Location = New System.Drawing.Point(326, 13)
        Me.txtPrjNm.MaxLength = 30
        Me.txtPrjNm.Name = "txtPrjNm"
        Me.txtPrjNm.Size = New System.Drawing.Size(278, 20)
        Me.txtPrjNm.TabIndex = 50
        '
        'txtUchiCd
        '
        Me.txtUchiCd.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtUchiCd.Location = New System.Drawing.Point(123, 40)
        Me.txtUchiCd.MaxLength = 10
        Me.txtUchiCd.Name = "txtUchiCd"
        Me.txtUchiCd.Size = New System.Drawing.Size(78, 20)
        Me.txtUchiCd.TabIndex = 49
        '
        'txtPrjCd
        '
        Me.txtPrjCd.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtPrjCd.Location = New System.Drawing.Point(123, 13)
        Me.txtPrjCd.MaxLength = 7
        Me.txtPrjCd.Name = "txtPrjCd"
        Me.txtPrjCd.Size = New System.Drawing.Size(78, 20)
        Me.txtPrjCd.TabIndex = 48
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 19)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "内訳コード"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(223, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 19)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "プロジェクト名称"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(223, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 19)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "内訳名称"
        '
        'fbResult
        '
        Me.fbResult.AccessibleDescription = "FpSpread1, Sheet1, Row 0, Column 0, "
        Me.fbResult.Location = New System.Drawing.Point(12, 151)
        Me.fbResult.Name = "fbResult"
        Me.fbResult.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbResult_Sheet1})
        Me.fbResult.Size = New System.Drawing.Size(665, 261)
        Me.fbResult.TabIndex = 91
        '
        'fbResult_Sheet1
        '
        Me.fbResult_Sheet1.Reset()
        Me.fbResult_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbResult_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbResult_Sheet1.ColumnCount = 5
        Me.fbResult_Sheet1.RowCount = 0
        Me.fbResult_Sheet1.ActiveColumnIndex = -1
        Me.fbResult_Sheet1.ActiveRowIndex = -1
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "プロジェクト名"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "内訳コード"
        Me.fbResult_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "プロジェクト内訳名称"
        Me.fbResult_Sheet1.ColumnHeader.Rows.Get(0).Height = 39.0!
        Me.fbResult_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.fbResult_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbResult_Sheet1.Columns.Get(0).Label = "選択"
        Me.fbResult_Sheet1.Columns.Get(0).Width = 35.0!
        Me.fbResult_Sheet1.Columns.Get(1).Label = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.fbResult_Sheet1.Columns.Get(1).Width = 94.0!
        Me.fbResult_Sheet1.Columns.Get(2).Label = "プロジェクト名"
        Me.fbResult_Sheet1.Columns.Get(2).Width = 184.0!
        Me.fbResult_Sheet1.Columns.Get(3).Label = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "内訳コード"
        Me.fbResult_Sheet1.Columns.Get(3).Width = 87.0!
        Me.fbResult_Sheet1.Columns.Get(4).Label = "プロジェクト内訳名称"
        Me.fbResult_Sheet1.Columns.Get(4).Width = 220.0!
        Me.fbResult_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbResult_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTZ0205
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(709, 503)
        Me.Controls.Add(Me.fbResult)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnKakutei)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Button5)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0205"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　プロジェクト一覧"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.fbResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbResult_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnKakutei As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents txtUchiNm As System.Windows.Forms.TextBox
    Friend WithEvents txtPrjNm As System.Windows.Forms.TextBox
    Friend WithEvents txtUchiCd As System.Windows.Forms.TextBox
    Friend WithEvents txtPrjCd As System.Windows.Forms.TextBox
    Friend WithEvents fbResult As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbResult_Sheet1 As FarPoint.Win.Spread.SheetView

End Class
