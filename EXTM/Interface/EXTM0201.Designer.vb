<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTM0201
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
        Dim ButtonCellType1 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTM0201))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Btn_Kensaku = New System.Windows.Forms.Button()
        Me.Txt_AitesakiNm = New System.Windows.Forms.TextBox()
        Me.Txt_AitesakiCd = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Txt_RiyoCd = New System.Windows.Forms.TextBox()
        Me.Txt_Tel3 = New System.Windows.Forms.TextBox()
        Me.Txt_Tel2 = New System.Windows.Forms.TextBox()
        Me.Txt_Tel1 = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Txt_RiyoNm = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Txt_RiyoKana = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Rdo_Huka = New System.Windows.Forms.RadioButton()
        Me.Rdo_chui = New System.Windows.Forms.RadioButton()
        Me.Rdo_Tujyo = New System.Windows.Forms.RadioButton()
        Me.Btn_Touroku = New System.Windows.Forms.Button()
        Me.Btn_Kakutei = New System.Windows.Forms.Button()
        Me.Btn_Modoru = New System.Windows.Forms.Button()
        Me.FpSpread1 = New FarPoint.Win.Spread.FpSpread()
        Me.FpSpread1_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.Btn_Kensaku)
        Me.GroupBox1.Controls.Add(Me.Txt_AitesakiNm)
        Me.GroupBox1.Controls.Add(Me.Txt_AitesakiCd)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Txt_RiyoCd)
        Me.GroupBox1.Controls.Add(Me.Txt_Tel3)
        Me.GroupBox1.Controls.Add(Me.Txt_Tel2)
        Me.GroupBox1.Controls.Add(Me.Txt_Tel1)
        Me.GroupBox1.Controls.Add(Me.Label39)
        Me.GroupBox1.Controls.Add(Me.Label32)
        Me.GroupBox1.Controls.Add(Me.Txt_RiyoNm)
        Me.GroupBox1.Controls.Add(Me.Label24)
        Me.GroupBox1.Controls.Add(Me.Txt_RiyoKana)
        Me.GroupBox1.Controls.Add(Me.Label23)
        Me.GroupBox1.Controls.Add(Me.Label22)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(33, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1272, 203)
        Me.GroupBox1.TabIndex = 127
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "検索条件"
        '
        'Btn_Kensaku
        '
        Me.Btn_Kensaku.Location = New System.Drawing.Point(974, 156)
        Me.Btn_Kensaku.Name = "Btn_Kensaku"
        Me.Btn_Kensaku.Size = New System.Drawing.Size(124, 29)
        Me.Btn_Kensaku.TabIndex = 73
        Me.Btn_Kensaku.Text = "検索"
        Me.Btn_Kensaku.UseVisualStyleBackColor = True
        '
        'Txt_AitesakiNm
        '
        Me.Txt_AitesakiNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_AitesakiNm.Location = New System.Drawing.Point(163, 161)
        Me.Txt_AitesakiNm.Name = "Txt_AitesakiNm"
        Me.Txt_AitesakiNm.Size = New System.Drawing.Size(377, 20)
        Me.Txt_AitesakiNm.TabIndex = 71
        '
        'Txt_AitesakiCd
        '
        Me.Txt_AitesakiCd.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_AitesakiCd.Location = New System.Drawing.Point(163, 129)
        Me.Txt_AitesakiCd.Name = "Txt_AitesakiCd"
        Me.Txt_AitesakiCd.Size = New System.Drawing.Size(131, 20)
        Me.Txt_AitesakiCd.TabIndex = 70
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(31, 164)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 69
        Me.Label1.Text = "EXAS相手先名"
        '
        'Txt_RiyoCd
        '
        Me.Txt_RiyoCd.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_RiyoCd.Location = New System.Drawing.Point(163, 31)
        Me.Txt_RiyoCd.Name = "Txt_RiyoCd"
        Me.Txt_RiyoCd.Size = New System.Drawing.Size(131, 20)
        Me.Txt_RiyoCd.TabIndex = 68
        '
        'Txt_Tel3
        '
        Me.Txt_Tel3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_Tel3.Location = New System.Drawing.Point(931, 30)
        Me.Txt_Tel3.Name = "Txt_Tel3"
        Me.Txt_Tel3.Size = New System.Drawing.Size(66, 20)
        Me.Txt_Tel3.TabIndex = 67
        '
        'Txt_Tel2
        '
        Me.Txt_Tel2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_Tel2.Location = New System.Drawing.Point(846, 30)
        Me.Txt_Tel2.Name = "Txt_Tel2"
        Me.Txt_Tel2.Size = New System.Drawing.Size(66, 20)
        Me.Txt_Tel2.TabIndex = 66
        '
        'Txt_Tel1
        '
        Me.Txt_Tel1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_Tel1.Location = New System.Drawing.Point(761, 31)
        Me.Txt_Tel1.Name = "Txt_Tel1"
        Me.Txt_Tel1.Size = New System.Drawing.Size(66, 20)
        Me.Txt_Tel1.TabIndex = 65
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label39.Location = New System.Drawing.Point(662, 34)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(59, 13)
        Me.Label39.TabIndex = 64
        Me.Label39.Text = "電話番号"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label32.Location = New System.Drawing.Point(31, 132)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(106, 13)
        Me.Label32.TabIndex = 63
        Me.Label32.Text = "EXAS相手先コード"
        '
        'Txt_RiyoNm
        '
        Me.Txt_RiyoNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_RiyoNm.Location = New System.Drawing.Point(163, 94)
        Me.Txt_RiyoNm.Name = "Txt_RiyoNm"
        Me.Txt_RiyoNm.Size = New System.Drawing.Size(377, 20)
        Me.Txt_RiyoNm.TabIndex = 62
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label24.Location = New System.Drawing.Point(31, 94)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(102, 26)
        Me.Label24.TabIndex = 61
        Me.Label24.Text = "利用者名" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(会社名、団体名)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Txt_RiyoKana
        '
        Me.Txt_RiyoKana.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Txt_RiyoKana.Location = New System.Drawing.Point(163, 60)
        Me.Txt_RiyoKana.Name = "Txt_RiyoKana"
        Me.Txt_RiyoKana.Size = New System.Drawing.Size(377, 20)
        Me.Txt_RiyoKana.TabIndex = 60
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label23.Location = New System.Drawing.Point(31, 63)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(89, 26)
        Me.Label23.TabIndex = 59
        Me.Label23.Text = "利用者名(ｶﾅ)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(会社名、団体)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label22.Location = New System.Drawing.Point(31, 38)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(72, 13)
        Me.Label22.TabIndex = 58
        Me.Label22.Text = "利用者番号"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label19.Location = New System.Drawing.Point(663, 77)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(34, 12)
        Me.Label19.TabIndex = 35
        Me.Label19.Text = "レベル"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Rdo_Huka)
        Me.Panel1.Controls.Add(Me.Rdo_chui)
        Me.Panel1.Controls.Add(Me.Rdo_Tujyo)
        Me.Panel1.Location = New System.Drawing.Point(761, 63)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(278, 37)
        Me.Panel1.TabIndex = 34
        '
        'Rdo_Huka
        '
        Me.Rdo_Huka.AutoSize = True
        Me.Rdo_Huka.Location = New System.Drawing.Point(169, 14)
        Me.Rdo_Huka.Name = "Rdo_Huka"
        Me.Rdo_Huka.Size = New System.Drawing.Size(77, 17)
        Me.Rdo_Huka.TabIndex = 35
        Me.Rdo_Huka.Text = "利用不可"
        Me.Rdo_Huka.UseVisualStyleBackColor = True
        '
        'Rdo_chui
        '
        Me.Rdo_chui.AutoSize = True
        Me.Rdo_chui.Location = New System.Drawing.Point(84, 14)
        Me.Rdo_chui.Name = "Rdo_chui"
        Me.Rdo_chui.Size = New System.Drawing.Size(64, 17)
        Me.Rdo_chui.TabIndex = 34
        Me.Rdo_chui.Text = "要注意"
        Me.Rdo_chui.UseVisualStyleBackColor = True
        '
        'Rdo_Tujyo
        '
        Me.Rdo_Tujyo.AutoSize = True
        Me.Rdo_Tujyo.Checked = True
        Me.Rdo_Tujyo.Location = New System.Drawing.Point(15, 14)
        Me.Rdo_Tujyo.Name = "Rdo_Tujyo"
        Me.Rdo_Tujyo.Size = New System.Drawing.Size(51, 17)
        Me.Rdo_Tujyo.TabIndex = 33
        Me.Rdo_Tujyo.TabStop = True
        Me.Rdo_Tujyo.Text = "通常"
        Me.Rdo_Tujyo.UseVisualStyleBackColor = True
        '
        'Btn_Touroku
        '
        Me.Btn_Touroku.Location = New System.Drawing.Point(33, 251)
        Me.Btn_Touroku.Name = "Btn_Touroku"
        Me.Btn_Touroku.Size = New System.Drawing.Size(124, 29)
        Me.Btn_Touroku.TabIndex = 74
        Me.Btn_Touroku.Text = "新規登録"
        Me.Btn_Touroku.UseVisualStyleBackColor = True
        '
        'Btn_Kakutei
        '
        Me.Btn_Kakutei.Location = New System.Drawing.Point(991, 902)
        Me.Btn_Kakutei.Name = "Btn_Kakutei"
        Me.Btn_Kakutei.Size = New System.Drawing.Size(124, 35)
        Me.Btn_Kakutei.TabIndex = 128
        Me.Btn_Kakutei.Text = "選択確定"
        Me.Btn_Kakutei.UseVisualStyleBackColor = True
        '
        'Btn_Modoru
        '
        Me.Btn_Modoru.Location = New System.Drawing.Point(536, 902)
        Me.Btn_Modoru.Name = "Btn_Modoru"
        Me.Btn_Modoru.Size = New System.Drawing.Size(124, 35)
        Me.Btn_Modoru.TabIndex = 129
        Me.Btn_Modoru.Text = "戻る"
        Me.Btn_Modoru.UseVisualStyleBackColor = True
        '
        'FpSpread1
        '
        Me.FpSpread1.AccessibleDescription = "FpSpread1, Sheet1, Row 0, Column 0, "
        Me.FpSpread1.Location = New System.Drawing.Point(33, 297)
        Me.FpSpread1.Name = "FpSpread1"
        Me.FpSpread1.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread1_Sheet1})
        Me.FpSpread1.Size = New System.Drawing.Size(1570, 582)
        Me.FpSpread1.TabIndex = 130
        '
        'FpSpread1_Sheet1
        '
        Me.FpSpread1_Sheet1.Reset()
        Me.FpSpread1_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.FpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.FpSpread1_Sheet1.ColumnCount = 29
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "利用者番号"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "ﾚﾍﾞﾙ"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用者名"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "利用者名ｶﾅ"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "最終利用日"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "相手先ｺｰﾄﾞ"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "相手先名"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "電話番号"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "携帯"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "FAX"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "郵便番号"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "住所"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "詳細"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "レベル区分"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "責任者"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 16).Value = "TEL１"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 17).Value = "TEL２"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 18).Value = "TEL３"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 19).Value = "内線"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 20).Value = "FAX１"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 21).Value = "FAX２"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 22).Value = "FAX３"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 23).Value = "郵便番号１"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 24).Value = "郵便番号２"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 25).Value = "住所１"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 26).Value = "住所２"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 27).Value = "住所３"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 28).Value = "住所４"
        Me.FpSpread1_Sheet1.ColumnHeader.Rows.Get(0).Height = 21.0!
        Me.FpSpread1_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.FpSpread1_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.FpSpread1_Sheet1.Columns.Get(0).Label = "選択"
        Me.FpSpread1_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Columns.Get(0).Width = 34.0!
        Me.FpSpread1_Sheet1.Columns.Get(1).Label = "利用者番号"
        Me.FpSpread1_Sheet1.Columns.Get(1).Width = 81.0!
        Me.FpSpread1_Sheet1.Columns.Get(2).Label = "ﾚﾍﾞﾙ"
        Me.FpSpread1_Sheet1.Columns.Get(2).Width = 69.0!
        Me.FpSpread1_Sheet1.Columns.Get(3).Label = "利用者名"
        Me.FpSpread1_Sheet1.Columns.Get(3).Width = 165.0!
        Me.FpSpread1_Sheet1.Columns.Get(4).Label = "利用者名ｶﾅ"
        Me.FpSpread1_Sheet1.Columns.Get(4).Width = 130.0!
        Me.FpSpread1_Sheet1.Columns.Get(5).Label = "最終利用日"
        Me.FpSpread1_Sheet1.Columns.Get(5).Width = 93.0!
        Me.FpSpread1_Sheet1.Columns.Get(6).Label = "相手先ｺｰﾄﾞ"
        Me.FpSpread1_Sheet1.Columns.Get(6).Width = 82.0!
        Me.FpSpread1_Sheet1.Columns.Get(7).Label = "相手先名"
        Me.FpSpread1_Sheet1.Columns.Get(7).Width = 172.0!
        Me.FpSpread1_Sheet1.Columns.Get(8).Label = "電話番号"
        Me.FpSpread1_Sheet1.Columns.Get(8).Width = 86.0!
        Me.FpSpread1_Sheet1.Columns.Get(9).Label = "携帯"
        Me.FpSpread1_Sheet1.Columns.Get(9).Width = 86.0!
        Me.FpSpread1_Sheet1.Columns.Get(10).Label = "FAX"
        Me.FpSpread1_Sheet1.Columns.Get(10).Width = 86.0!
        Me.FpSpread1_Sheet1.Columns.Get(11).Label = "郵便番号"
        Me.FpSpread1_Sheet1.Columns.Get(11).Width = 77.0!
        Me.FpSpread1_Sheet1.Columns.Get(12).Label = "住所"
        Me.FpSpread1_Sheet1.Columns.Get(12).Width = 278.0!
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "確認・編集"
        Me.FpSpread1_Sheet1.Columns.Get(13).CellType = ButtonCellType1
        Me.FpSpread1_Sheet1.Columns.Get(13).Label = "詳細"
        Me.FpSpread1_Sheet1.Columns.Get(13).Width = 75.0!
        Me.FpSpread1_Sheet1.Columns.Get(14).Label = "レベル区分"
        Me.FpSpread1_Sheet1.Columns.Get(14).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(15).Label = "責任者"
        Me.FpSpread1_Sheet1.Columns.Get(15).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(16).Label = "TEL１"
        Me.FpSpread1_Sheet1.Columns.Get(16).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(17).Label = "TEL２"
        Me.FpSpread1_Sheet1.Columns.Get(17).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(18).Label = "TEL３"
        Me.FpSpread1_Sheet1.Columns.Get(18).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(19).Label = "内線"
        Me.FpSpread1_Sheet1.Columns.Get(19).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(20).Label = "FAX１"
        Me.FpSpread1_Sheet1.Columns.Get(20).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(21).Label = "FAX２"
        Me.FpSpread1_Sheet1.Columns.Get(21).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(22).Label = "FAX３"
        Me.FpSpread1_Sheet1.Columns.Get(22).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(23).Label = "郵便番号１"
        Me.FpSpread1_Sheet1.Columns.Get(23).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(24).Label = "郵便番号２"
        Me.FpSpread1_Sheet1.Columns.Get(24).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(25).Label = "住所１"
        Me.FpSpread1_Sheet1.Columns.Get(25).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(26).Label = "住所２"
        Me.FpSpread1_Sheet1.Columns.Get(26).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(27).Label = "住所３"
        Me.FpSpread1_Sheet1.Columns.Get(27).Visible = False
        Me.FpSpread1_Sheet1.Columns.Get(28).Label = "住所４"
        Me.FpSpread1_Sheet1.Columns.Get(28).Visible = False
        Me.FpSpread1_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.FpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(35, 886)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(494, 13)
        Me.Label2.TabIndex = 131
        Me.Label2.Text = "※表のデータは、コピーしたい部分を選択し、キーボードの「Ctrl」＋「C」を押すとコピーできます。"
        '
        'EXTM0201
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTM.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1612, 944)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.FpSpread1)
        Me.Controls.Add(Me.Btn_Modoru)
        Me.Controls.Add(Me.Btn_Kakutei)
        Me.Controls.Add(Me.Btn_Touroku)
        Me.Controls.Add(Me.GroupBox1)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTM0201"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　利用者一覧"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Rdo_Huka As System.Windows.Forms.RadioButton
    Friend WithEvents Rdo_chui As System.Windows.Forms.RadioButton
    Friend WithEvents Rdo_Tujyo As System.Windows.Forms.RadioButton
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Txt_AitesakiNm As System.Windows.Forms.TextBox
    Friend WithEvents Txt_AitesakiCd As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Txt_RiyoCd As System.Windows.Forms.TextBox
    Friend WithEvents Txt_Tel3 As System.Windows.Forms.TextBox
    Friend WithEvents Txt_Tel2 As System.Windows.Forms.TextBox
    Friend WithEvents Txt_Tel1 As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Txt_RiyoNm As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Txt_RiyoKana As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Btn_Kensaku As System.Windows.Forms.Button
    Friend WithEvents Btn_Touroku As System.Windows.Forms.Button
    Friend WithEvents Btn_Kakutei As System.Windows.Forms.Button
    Friend WithEvents Btn_Modoru As System.Windows.Forms.Button
    Friend WithEvents FpSpread1 As FarPoint.Win.Spread.FpSpread

    Friend WithEvents FpSpread1_Sheet1 As FarPoint.Win.Spread.SheetView

    Friend WithEvents vwList As FarPoint.Win.Spread.FpSpread
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
