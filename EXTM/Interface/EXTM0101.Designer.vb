<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTM0101
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
        Dim CheckBoxCellType1 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim TextCellType1 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim ComboBoxCellType1 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim CheckBoxCellType2 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim CheckBoxCellType3 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim CheckBoxCellType4 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTM0101))
        Me.BtnBack = New System.Windows.Forms.Button()
        Me.BtnReg = New System.Windows.Forms.Button()
        Me.ppVwList = New FarPoint.Win.Spread.FpSpread()
        Me.ppVwList_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkStsFlg = New System.Windows.Forms.CheckBox()
        Me.chkMstFlg = New System.Windows.Forms.CheckBox()
        Me.chkShoninFlg = New System.Windows.Forms.CheckBox()
        Me.txtMail = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmbBushoName = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnSearch = New System.Windows.Forms.Button()
        Me.txtUserKanjiName = New System.Windows.Forms.TextBox()
        Me.txtUserId = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        CType(Me.ppVwList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ppVwList_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnBack
        '
        Me.BtnBack.Location = New System.Drawing.Point(387, 702)
        Me.BtnBack.Name = "BtnBack"
        Me.BtnBack.Size = New System.Drawing.Size(114, 35)
        Me.BtnBack.TabIndex = 112
        Me.BtnBack.Text = "戻る"
        Me.BtnBack.UseVisualStyleBackColor = True
        '
        'BtnReg
        '
        Me.BtnReg.Location = New System.Drawing.Point(730, 702)
        Me.BtnReg.Name = "BtnReg"
        Me.BtnReg.Size = New System.Drawing.Size(114, 35)
        Me.BtnReg.TabIndex = 111
        Me.BtnReg.Text = "登録"
        Me.BtnReg.UseVisualStyleBackColor = True
        '
        'ppVwList
        '
        Me.ppVwList.AccessibleDescription = "ppVwList, Sheet1, Row 0, Column 0, "
        Me.ppVwList.Location = New System.Drawing.Point(42, 149)
        Me.ppVwList.Name = "ppVwList"
        Me.ppVwList.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.ppVwList_Sheet1})
        Me.ppVwList.Size = New System.Drawing.Size(1175, 533)
        Me.ppVwList.TabIndex = 113
        '
        'ppVwList_Sheet1
        '
        Me.ppVwList_Sheet1.Reset()
        Me.ppVwList_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.ppVwList_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.ppVwList_Sheet1.ColumnCount = 12
        Me.ppVwList_Sheet1.Cells.Get(0, 1).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.ppVwList_Sheet1.Cells.Get(0, 1).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.ppVwList_Sheet1.Cells.Get(0, 1).ParseFormatString = "n"
        Me.ppVwList_Sheet1.Cells.Get(0, 1).Value = 1234
        Me.ppVwList_Sheet1.Cells.Get(0, 2).Value = "abcde123"
        Me.ppVwList_Sheet1.Cells.Get(0, 3).Value = "予約太郎"
        Me.ppVwList_Sheet1.Cells.Get(0, 4).Value = "yoyaku@////"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "変更対象"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "ユーザID" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（社員番号）"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "パスワード"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "漢字氏名"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "メールアドレス"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "部署名"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "承認者"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "マスタ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "操作"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "無効"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "初回登録日"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "最終更新日"
        Me.ppVwList_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "更新区分"
        Me.ppVwList_Sheet1.ColumnHeader.Rows.Get(0).Height = 33.0!
        Me.ppVwList_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.ppVwList_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(0).Label = "変更対象"
        Me.ppVwList_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(0).Width = 39.0!
        TextCellType1.AllowEditorVerticalAlign = True
        Me.ppVwList_Sheet1.Columns.Get(1).CellType = TextCellType1
        Me.ppVwList_Sheet1.Columns.Get(1).Label = "ユーザID" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（社員番号）"
        Me.ppVwList_Sheet1.Columns.Get(1).Width = 134.0!
        Me.ppVwList_Sheet1.Columns.Get(2).Label = "パスワード"
        Me.ppVwList_Sheet1.Columns.Get(2).Width = 134.0!
        Me.ppVwList_Sheet1.Columns.Get(3).Label = "漢字氏名"
        Me.ppVwList_Sheet1.Columns.Get(3).Width = 134.0!
        Me.ppVwList_Sheet1.Columns.Get(4).Label = "メールアドレス"
        Me.ppVwList_Sheet1.Columns.Get(4).Width = 160.0!
        Me.ppVwList_Sheet1.Columns.Get(5).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        ComboBoxCellType1.AllowEditorVerticalAlign = True
        ComboBoxCellType1.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.ppVwList_Sheet1.Columns.Get(5).CellType = ComboBoxCellType1
        Me.ppVwList_Sheet1.Columns.Get(5).Label = "部署名"
        Me.ppVwList_Sheet1.Columns.Get(5).Width = 172.0!
        Me.ppVwList_Sheet1.Columns.Get(6).CellType = CheckBoxCellType2
        Me.ppVwList_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(6).Label = "承認者"
        Me.ppVwList_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(6).Width = 55.0!
        Me.ppVwList_Sheet1.Columns.Get(7).CellType = CheckBoxCellType3
        Me.ppVwList_Sheet1.Columns.Get(7).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(7).Label = "マスタ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "操作"
        Me.ppVwList_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(7).Width = 55.0!
        Me.ppVwList_Sheet1.Columns.Get(8).CellType = CheckBoxCellType4
        Me.ppVwList_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.ppVwList_Sheet1.Columns.Get(8).Label = "無効"
        Me.ppVwList_Sheet1.Columns.Get(8).Width = 55.0!
        Me.ppVwList_Sheet1.Columns.Get(9).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.ppVwList_Sheet1.Columns.Get(9).Label = "初回登録日"
        Me.ppVwList_Sheet1.Columns.Get(9).Locked = True
        Me.ppVwList_Sheet1.Columns.Get(9).Width = 91.0!
        Me.ppVwList_Sheet1.Columns.Get(10).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.ppVwList_Sheet1.Columns.Get(10).Label = "最終更新日"
        Me.ppVwList_Sheet1.Columns.Get(10).Locked = True
        Me.ppVwList_Sheet1.Columns.Get(10).Width = 91.0!
        Me.ppVwList_Sheet1.Columns.Get(11).Label = "更新区分"
        Me.ppVwList_Sheet1.Columns.Get(11).Visible = False
        Me.ppVwList_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.ppVwList_Sheet1.Rows.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(14).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(15).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(16).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(17).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(18).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(19).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(20).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(21).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(22).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(23).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(24).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(25).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(26).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(27).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(28).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(29).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(30).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(31).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(32).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(33).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(34).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(35).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(36).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(37).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(38).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(39).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(40).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(41).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(42).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(43).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(44).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(45).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(46).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(47).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(48).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(49).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(50).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(51).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(52).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(53).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(54).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(55).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(56).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(57).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(58).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(59).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(60).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(61).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(62).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(63).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(64).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(65).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(66).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(67).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(68).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(69).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(70).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(71).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(72).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(73).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(74).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(75).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(76).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(77).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(78).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(79).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(80).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(81).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(82).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(83).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(84).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(85).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(86).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(87).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(88).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(89).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(90).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(91).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(92).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(93).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(94).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(95).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(96).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(97).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(98).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(99).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(100).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(101).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(102).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(103).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(104).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(105).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(106).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(107).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(108).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(109).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(110).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(111).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(112).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(113).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(114).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(115).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(116).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(117).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(118).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(119).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(120).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(121).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(122).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(123).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(124).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(125).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(126).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(127).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(128).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(129).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(130).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(131).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(132).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(133).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(134).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(135).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(136).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(137).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(138).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(139).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(140).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(141).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(142).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(143).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(144).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(145).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(146).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(147).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(148).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(149).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(150).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(151).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(152).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(153).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(154).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(155).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(156).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(157).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(158).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(159).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(160).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(161).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(162).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(163).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(164).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(165).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(166).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(167).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(168).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(169).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(170).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(171).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(172).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(173).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(174).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(175).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(176).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(177).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(178).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(179).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(180).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(181).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(182).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(183).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(184).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(185).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(186).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(187).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(188).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(189).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(190).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(191).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(192).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(193).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(194).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(195).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(196).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(197).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(198).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(199).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(200).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(201).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(202).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(203).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(204).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(205).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(206).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(207).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(208).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(209).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(210).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(211).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(212).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(213).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(214).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(215).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(216).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(217).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(218).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(219).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(220).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(221).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(222).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(223).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(224).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(225).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(226).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(227).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(228).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(229).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(230).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(231).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(232).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(233).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(234).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(235).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(236).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(237).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(238).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(239).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(240).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(241).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(242).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(243).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(244).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(245).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(246).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(247).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(248).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(249).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(250).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(251).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(252).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(253).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(254).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(255).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(256).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(257).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(258).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(259).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(260).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(261).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(262).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(263).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(264).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(265).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(266).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(267).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(268).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(269).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(270).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(271).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(272).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(273).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(274).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(275).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(276).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(277).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(278).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(279).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(280).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(281).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(282).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(283).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(284).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(285).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(286).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(287).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(288).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(289).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(290).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(291).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(292).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(293).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(294).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(295).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(296).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(297).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(298).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(299).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(300).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(301).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(302).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(303).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(304).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(305).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(306).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(307).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(308).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(309).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(310).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(311).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(312).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(313).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(314).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(315).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(316).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(317).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(318).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(319).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(320).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(321).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(322).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(323).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(324).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(325).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(326).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(327).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(328).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(329).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(330).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(331).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(332).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(333).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(334).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(335).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(336).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(337).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(338).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(339).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(340).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(341).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(342).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(343).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(344).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(345).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(346).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(347).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(348).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(349).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(350).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(351).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(352).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(353).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(354).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(355).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(356).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(357).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(358).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(359).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(360).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(361).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(362).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(363).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(364).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(365).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(366).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(367).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(368).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(369).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(370).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(371).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(372).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(373).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(374).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(375).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(376).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(377).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(378).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(379).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(380).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(381).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(382).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(383).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(384).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(385).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(386).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(387).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(388).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(389).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(390).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(391).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(392).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(393).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(394).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(395).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(396).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(397).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(398).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(399).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(400).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(401).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(402).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(403).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(404).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(405).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(406).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(407).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(408).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(409).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(410).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(411).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(412).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(413).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(414).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(415).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(416).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(417).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(418).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(419).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(420).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(421).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(422).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(423).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(424).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(425).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(426).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(427).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(428).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(429).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(430).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(431).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(432).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(433).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(434).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(435).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(436).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(437).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(438).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(439).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(440).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(441).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(442).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(443).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(444).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(445).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(446).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(447).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(448).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(449).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(450).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(451).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(452).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(453).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(454).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(455).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(456).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(457).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(458).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(459).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(460).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(461).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(462).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(463).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(464).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(465).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(466).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(467).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(468).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(469).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(470).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(471).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(472).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(473).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(474).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(475).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(476).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(477).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(478).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(479).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(480).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(481).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(482).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(483).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(484).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(485).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(486).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(487).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(488).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(489).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(490).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(491).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(492).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(493).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(494).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(495).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(496).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(497).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(498).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.Rows.Get(499).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.ppVwList_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.chkStsFlg)
        Me.GroupBox1.Controls.Add(Me.chkMstFlg)
        Me.GroupBox1.Controls.Add(Me.chkShoninFlg)
        Me.GroupBox1.Controls.Add(Me.txtMail)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.cmbBushoName)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.BtnSearch)
        Me.GroupBox1.Controls.Add(Me.txtUserKanjiName)
        Me.GroupBox1.Controls.Add(Me.txtUserId)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label32)
        Me.GroupBox1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(42, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1175, 119)
        Me.GroupBox1.TabIndex = 131
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "検索条件"
        '
        'chkStsFlg
        '
        Me.chkStsFlg.AutoSize = True
        Me.chkStsFlg.Location = New System.Drawing.Point(736, 74)
        Me.chkStsFlg.Name = "chkStsFlg"
        Me.chkStsFlg.Size = New System.Drawing.Size(177, 17)
        Me.chkStsFlg.TabIndex = 80
        Me.chkStsFlg.Text = "無効なユーザーも含めて表示"
        Me.chkStsFlg.UseVisualStyleBackColor = True
        '
        'chkMstFlg
        '
        Me.chkMstFlg.AutoSize = True
        Me.chkMstFlg.Location = New System.Drawing.Point(526, 74)
        Me.chkMstFlg.Name = "chkMstFlg"
        Me.chkMstFlg.Size = New System.Drawing.Size(169, 17)
        Me.chkMstFlg.TabIndex = 79
        Me.chkMstFlg.Text = "マスタ操作可能者のみ表示"
        Me.chkMstFlg.UseVisualStyleBackColor = True
        '
        'chkShoninFlg
        '
        Me.chkShoninFlg.AutoSize = True
        Me.chkShoninFlg.Location = New System.Drawing.Point(371, 76)
        Me.chkShoninFlg.Name = "chkShoninFlg"
        Me.chkShoninFlg.Size = New System.Drawing.Size(114, 17)
        Me.chkShoninFlg.TabIndex = 78
        Me.chkShoninFlg.Text = "承認者のみ表示"
        Me.chkShoninFlg.UseVisualStyleBackColor = True
        '
        'txtMail
        '
        Me.txtMail.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtMail.Location = New System.Drawing.Point(772, 35)
        Me.txtMail.Name = "txtMail"
        Me.txtMail.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMail.Size = New System.Drawing.Size(253, 20)
        Me.txtMail.TabIndex = 77
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(685, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 13)
        Me.Label3.TabIndex = 76
        Me.Label3.Text = "メールアドレス"
        '
        'cmbBushoName
        '
        Me.cmbBushoName.FormattingEnabled = True
        Me.cmbBushoName.Location = New System.Drawing.Point(103, 72)
        Me.cmbBushoName.Name = "cmbBushoName"
        Me.cmbBushoName.Size = New System.Drawing.Size(237, 21)
        Me.cmbBushoName.TabIndex = 75
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(31, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 74
        Me.Label2.Text = "部署名"
        '
        'BtnSearch
        '
        Me.BtnSearch.Location = New System.Drawing.Point(1026, 84)
        Me.BtnSearch.Name = "BtnSearch"
        Me.BtnSearch.Size = New System.Drawing.Size(124, 29)
        Me.BtnSearch.TabIndex = 73
        Me.BtnSearch.Text = "検索"
        Me.BtnSearch.UseVisualStyleBackColor = True
        '
        'txtUserKanjiName
        '
        Me.txtUserKanjiName.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtUserKanjiName.Location = New System.Drawing.Point(455, 35)
        Me.txtUserKanjiName.Name = "txtUserKanjiName"
        Me.txtUserKanjiName.Size = New System.Drawing.Size(185, 20)
        Me.txtUserKanjiName.TabIndex = 71
        '
        'txtUserId
        '
        Me.txtUserId.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtUserId.Location = New System.Drawing.Point(103, 35)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(163, 20)
        Me.txtUserId.TabIndex = 70
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(368, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 69
        Me.Label1.Text = "漢字氏名"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label32.Location = New System.Drawing.Point(31, 38)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(51, 13)
        Me.Label32.TabIndex = 63
        Me.Label32.Text = "ユーザID"
        '
        'EXTM0101
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTM.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1249, 752)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ppVwList)
        Me.Controls.Add(Me.BtnBack)
        Me.Controls.Add(Me.BtnReg)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTM0101"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　ユーザマスタメンテ"
        CType(Me.ppVwList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ppVwList_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnBack As System.Windows.Forms.Button
    Friend WithEvents BtnReg As System.Windows.Forms.Button
    Friend WithEvents ppVwList As FarPoint.Win.Spread.FpSpread
    Friend WithEvents ppVwList_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbBushoName As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnSearch As System.Windows.Forms.Button
    Friend WithEvents txtUserKanjiName As System.Windows.Forms.TextBox
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents chkMstFlg As System.Windows.Forms.CheckBox
    Friend WithEvents chkShoninFlg As System.Windows.Forms.CheckBox
    Friend WithEvents txtMail As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkStsFlg As System.Windows.Forms.CheckBox
End Class
