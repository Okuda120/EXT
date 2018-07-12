<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0207
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
        Dim ButtonCellType1 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0207))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.btnComplate = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.lblSeikyuNo = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblKakuteiKin = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblShokei = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblTaxKin = New System.Windows.Forms.Label()
        Me.lblSeikyuNaiyo = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lblSeikyuKin = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.lblAiteNm = New System.Windows.Forms.Label()
        Me.lblAiteCd = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rdoNyukin2 = New System.Windows.Forms.RadioButton()
        Me.rdoNyukin1 = New System.Windows.Forms.RadioButton()
        Me.lblChoseiKin = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtNyukin = New System.Windows.Forms.TextBox()
        Me.btnSetNyukinTtl = New System.Windows.Forms.Button()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.btnExasList = New System.Windows.Forms.Button()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.lblNyukinTtl = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.lblNyukinYoteiDt = New System.Windows.Forms.Label()
        Me.lblSeikyuDt = New System.Windows.Forms.Label()
        Me.fbNyukin = New FarPoint.Win.Spread.FpSpread()
        Me.fbNyukin_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.dtpNyukinDt = New Common.DateTimePickerEx()
        Me.Panel1.SuspendLayout()
        CType(Me.fbNyukin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbNyukin_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.BackColor = System.Drawing.Color.Transparent
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(12, 18)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(103, 21)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "請求番号"
        '
        'btnComplate
        '
        Me.btnComplate.Location = New System.Drawing.Point(933, 410)
        Me.btnComplate.Name = "btnComplate"
        Me.btnComplate.Size = New System.Drawing.Size(140, 28)
        Me.btnComplate.TabIndex = 86
        Me.btnComplate.Text = "入力完了"
        Me.btnComplate.UseVisualStyleBackColor = True
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
        Me.btnBack.Location = New System.Drawing.Point(516, 410)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 28)
        Me.btnBack.TabIndex = 89
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'lblSeikyuNo
        '
        Me.lblSeikyuNo.BackColor = System.Drawing.Color.Silver
        Me.lblSeikyuNo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeikyuNo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSeikyuNo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSeikyuNo.Location = New System.Drawing.Point(121, 14)
        Me.lblSeikyuNo.Margin = New System.Windows.Forms.Padding(3)
        Me.lblSeikyuNo.Name = "lblSeikyuNo"
        Me.lblSeikyuNo.Size = New System.Drawing.Size(88, 20)
        Me.lblSeikyuNo.TabIndex = 90
        Me.lblSeikyuNo.Text = "000000"
        Me.lblSeikyuNo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 49)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 17)
        Me.Label1.TabIndex = 91
        Me.Label1.Text = "請求日"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 149)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 21)
        Me.Label4.TabIndex = 93
        Me.Label4.Text = "調整"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(12, 115)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(103, 21)
        Me.Label5.TabIndex = 94
        Me.Label5.Text = "確定額"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(12, 82)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(103, 21)
        Me.Label6.TabIndex = 95
        Me.Label6.Text = "入金予定日"
        '
        'lblKakuteiKin
        '
        Me.lblKakuteiKin.BackColor = System.Drawing.Color.Silver
        Me.lblKakuteiKin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKakuteiKin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblKakuteiKin.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblKakuteiKin.Location = New System.Drawing.Point(121, 116)
        Me.lblKakuteiKin.Margin = New System.Windows.Forms.Padding(3)
        Me.lblKakuteiKin.Name = "lblKakuteiKin"
        Me.lblKakuteiKin.Size = New System.Drawing.Size(110, 20)
        Me.lblKakuteiKin.TabIndex = 98
        Me.lblKakuteiKin.Text = "10,000"
        Me.lblKakuteiKin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Black
        Me.Label7.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.Location = New System.Drawing.Point(12, 181)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(210, 1)
        Me.Label7.TabIndex = 103
        Me.Label7.Text = "調整"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.Location = New System.Drawing.Point(12, 223)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(103, 17)
        Me.Label8.TabIndex = 106
        Me.Label8.Text = "消費税"
        '
        'lblShokei
        '
        Me.lblShokei.BackColor = System.Drawing.Color.Silver
        Me.lblShokei.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblShokei.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblShokei.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblShokei.Location = New System.Drawing.Point(121, 188)
        Me.lblShokei.Margin = New System.Windows.Forms.Padding(3)
        Me.lblShokei.Name = "lblShokei"
        Me.lblShokei.Size = New System.Drawing.Size(110, 20)
        Me.lblShokei.TabIndex = 105
        Me.lblShokei.Text = "000,000"
        Me.lblShokei.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.Location = New System.Drawing.Point(12, 192)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(103, 21)
        Me.Label10.TabIndex = 104
        Me.Label10.Text = "小計"
        '
        'lblTaxKin
        '
        Me.lblTaxKin.BackColor = System.Drawing.Color.Silver
        Me.lblTaxKin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTaxKin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblTaxKin.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTaxKin.Location = New System.Drawing.Point(121, 220)
        Me.lblTaxKin.Margin = New System.Windows.Forms.Padding(3)
        Me.lblTaxKin.Name = "lblTaxKin"
        Me.lblTaxKin.Size = New System.Drawing.Size(110, 20)
        Me.lblTaxKin.TabIndex = 107
        Me.lblTaxKin.Text = "0,000"
        Me.lblTaxKin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSeikyuNaiyo
        '
        Me.lblSeikyuNaiyo.BackColor = System.Drawing.Color.Silver
        Me.lblSeikyuNaiyo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeikyuNaiyo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSeikyuNaiyo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSeikyuNaiyo.Location = New System.Drawing.Point(121, 289)
        Me.lblSeikyuNaiyo.Margin = New System.Windows.Forms.Padding(3)
        Me.lblSeikyuNaiyo.Name = "lblSeikyuNaiyo"
        Me.lblSeikyuNaiyo.Size = New System.Drawing.Size(126, 20)
        Me.lblSeikyuNaiyo.TabIndex = 112
        Me.lblSeikyuNaiyo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label13.Location = New System.Drawing.Point(12, 292)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(103, 17)
        Me.Label13.TabIndex = 111
        Me.Label13.Text = "請求内容"
        '
        'lblSeikyuKin
        '
        Me.lblSeikyuKin.BackColor = System.Drawing.Color.Silver
        Me.lblSeikyuKin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeikyuKin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSeikyuKin.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSeikyuKin.Location = New System.Drawing.Point(121, 257)
        Me.lblSeikyuKin.Margin = New System.Windows.Forms.Padding(3)
        Me.lblSeikyuKin.Name = "lblSeikyuKin"
        Me.lblSeikyuKin.Size = New System.Drawing.Size(110, 20)
        Me.lblSeikyuKin.TabIndex = 110
        Me.lblSeikyuKin.Text = "000,000"
        Me.lblSeikyuKin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label15.Location = New System.Drawing.Point(12, 261)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(103, 21)
        Me.Label15.TabIndex = 109
        Me.Label15.Text = "請求金額"
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.Black
        Me.Label16.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label16.Location = New System.Drawing.Point(12, 250)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(210, 1)
        Me.Label16.TabIndex = 108
        Me.Label16.Text = "調整"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.Transparent
        Me.Label17.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label17.Location = New System.Drawing.Point(272, 18)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(103, 21)
        Me.Label17.TabIndex = 113
        Me.Label17.Text = "EXAS相手先情報"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label19.Location = New System.Drawing.Point(297, 82)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(103, 21)
        Me.Label19.TabIndex = 115
        Me.Label19.Text = "入金日"
        '
        'lblAiteNm
        '
        Me.lblAiteNm.BackColor = System.Drawing.Color.Silver
        Me.lblAiteNm.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAiteNm.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblAiteNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblAiteNm.Location = New System.Drawing.Point(476, 14)
        Me.lblAiteNm.Margin = New System.Windows.Forms.Padding(3)
        Me.lblAiteNm.Name = "lblAiteNm"
        Me.lblAiteNm.Size = New System.Drawing.Size(142, 20)
        Me.lblAiteNm.TabIndex = 116
        Me.lblAiteNm.Text = "相手先名"
        Me.lblAiteNm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAiteCd
        '
        Me.lblAiteCd.BackColor = System.Drawing.Color.Silver
        Me.lblAiteCd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAiteCd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblAiteCd.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblAiteCd.Location = New System.Drawing.Point(382, 14)
        Me.lblAiteCd.Margin = New System.Windows.Forms.Padding(3)
        Me.lblAiteCd.Name = "lblAiteCd"
        Me.lblAiteCd.Size = New System.Drawing.Size(88, 20)
        Me.lblAiteCd.TabIndex = 117
        Me.lblAiteCd.Text = "000000"
        Me.lblAiteCd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label22
        '
        Me.Label22.BackColor = System.Drawing.Color.Transparent
        Me.Label22.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label22.Location = New System.Drawing.Point(297, 115)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(103, 21)
        Me.Label22.TabIndex = 118
        Me.Label22.Text = "入金種別"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rdoNyukin2)
        Me.Panel1.Controls.Add(Me.rdoNyukin1)
        Me.Panel1.Location = New System.Drawing.Point(394, 108)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(176, 22)
        Me.Panel1.TabIndex = 119
        '
        'rdoNyukin2
        '
        Me.rdoNyukin2.AutoSize = True
        Me.rdoNyukin2.Location = New System.Drawing.Point(72, 4)
        Me.rdoNyukin2.Name = "rdoNyukin2"
        Me.rdoNyukin2.Size = New System.Drawing.Size(99, 17)
        Me.rdoNyukin2.TabIndex = 30
        Me.rdoNyukin2.Text = "現金(ALSOK)"
        Me.rdoNyukin2.UseVisualStyleBackColor = True
        '
        'rdoNyukin1
        '
        Me.rdoNyukin1.AutoSize = True
        Me.rdoNyukin1.Checked = True
        Me.rdoNyukin1.Location = New System.Drawing.Point(3, 4)
        Me.rdoNyukin1.Name = "rdoNyukin1"
        Me.rdoNyukin1.Size = New System.Drawing.Size(51, 17)
        Me.rdoNyukin1.TabIndex = 29
        Me.rdoNyukin1.TabStop = True
        Me.rdoNyukin1.Text = "振込"
        Me.rdoNyukin1.UseVisualStyleBackColor = True
        '
        'lblChoseiKin
        '
        Me.lblChoseiKin.BackColor = System.Drawing.Color.Silver
        Me.lblChoseiKin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblChoseiKin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblChoseiKin.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblChoseiKin.Location = New System.Drawing.Point(121, 145)
        Me.lblChoseiKin.Margin = New System.Windows.Forms.Padding(3)
        Me.lblChoseiKin.Name = "lblChoseiKin"
        Me.lblChoseiKin.Size = New System.Drawing.Size(110, 20)
        Me.lblChoseiKin.TabIndex = 121
        Me.lblChoseiKin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.Transparent
        Me.Label24.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label24.Location = New System.Drawing.Point(285, 49)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(103, 21)
        Me.Label24.TabIndex = 122
        Me.Label24.Text = "入金情報"
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label18.Location = New System.Drawing.Point(297, 144)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(103, 21)
        Me.Label18.TabIndex = 123
        Me.Label18.Text = "入金額"
        '
        'txtNyukin
        '
        Me.txtNyukin.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtNyukin.ImeMode = System.Windows.Forms.ImeMode.Alpha
        Me.txtNyukin.Location = New System.Drawing.Point(393, 141)
        Me.txtNyukin.MaxLength = 11
        Me.txtNyukin.Name = "txtNyukin"
        Me.txtNyukin.Size = New System.Drawing.Size(110, 20)
        Me.txtNyukin.TabIndex = 125
        '
        'btnSetNyukinTtl
        '
        Me.btnSetNyukinTtl.Location = New System.Drawing.Point(517, 139)
        Me.btnSetNyukinTtl.Name = "btnSetNyukinTtl"
        Me.btnSetNyukinTtl.Size = New System.Drawing.Size(96, 23)
        Me.btnSetNyukinTtl.TabIndex = 126
        Me.btnSetNyukinTtl.Text = "(A)をセット"
        Me.btnSetNyukinTtl.UseVisualStyleBackColor = True
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label25.Location = New System.Drawing.Point(655, 18)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(128, 21)
        Me.Label25.TabIndex = 127
        Me.Label25.Text = "EXAS入金データリンク"
        '
        'btnExasList
        '
        Me.btnExasList.Location = New System.Drawing.Point(1619, 13)
        Me.btnExasList.Name = "btnExasList"
        Me.btnExasList.Size = New System.Drawing.Size(141, 23)
        Me.btnExasList.TabIndex = 128
        Me.btnExasList.Text = "EXAS入金一覧"
        Me.btnExasList.UseVisualStyleBackColor = True
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label26.Location = New System.Drawing.Point(1412, 154)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(103, 21)
        Me.Label26.TabIndex = 129
        Me.Label26.Text = "入金合計額(A)"
        '
        'lblNyukinTtl
        '
        Me.lblNyukinTtl.BackColor = System.Drawing.Color.Silver
        Me.lblNyukinTtl.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNyukinTtl.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblNyukinTtl.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblNyukinTtl.Location = New System.Drawing.Point(1521, 150)
        Me.lblNyukinTtl.Margin = New System.Windows.Forms.Padding(3)
        Me.lblNyukinTtl.Name = "lblNyukinTtl"
        Me.lblNyukinTtl.Size = New System.Drawing.Size(156, 20)
        Me.lblNyukinTtl.TabIndex = 130
        Me.lblNyukinTtl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.Black
        Me.Label28.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label28.Location = New System.Drawing.Point(253, 14)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(1, 300)
        Me.Label28.TabIndex = 131
        Me.Label28.Text = "調整"
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.Color.Black
        Me.Label29.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label29.Location = New System.Drawing.Point(640, 14)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(1, 300)
        Me.Label29.TabIndex = 132
        Me.Label29.Text = "調整"
        '
        'lblNyukinYoteiDt
        '
        Me.lblNyukinYoteiDt.BackColor = System.Drawing.Color.Silver
        Me.lblNyukinYoteiDt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNyukinYoteiDt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblNyukinYoteiDt.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblNyukinYoteiDt.Location = New System.Drawing.Point(121, 78)
        Me.lblNyukinYoteiDt.Margin = New System.Windows.Forms.Padding(3)
        Me.lblNyukinYoteiDt.Name = "lblNyukinYoteiDt"
        Me.lblNyukinYoteiDt.Size = New System.Drawing.Size(88, 20)
        Me.lblNyukinYoteiDt.TabIndex = 134
        Me.lblNyukinYoteiDt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSeikyuDt
        '
        Me.lblSeikyuDt.BackColor = System.Drawing.Color.Silver
        Me.lblSeikyuDt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeikyuDt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSeikyuDt.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSeikyuDt.Location = New System.Drawing.Point(121, 45)
        Me.lblSeikyuDt.Margin = New System.Windows.Forms.Padding(3)
        Me.lblSeikyuDt.Name = "lblSeikyuDt"
        Me.lblSeikyuDt.Size = New System.Drawing.Size(88, 20)
        Me.lblSeikyuDt.TabIndex = 133
        Me.lblSeikyuDt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'fbNyukin
        '
        Me.fbNyukin.AccessibleDescription = "fbNyukin, Sheet1, Row 0, Column 0, "
        Me.fbNyukin.Location = New System.Drawing.Point(650, 52)
        Me.fbNyukin.Name = "fbNyukin"
        Me.fbNyukin.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbNyukin_Sheet1})
        Me.fbNyukin.Size = New System.Drawing.Size(1110, 77)
        Me.fbNyukin.TabIndex = 135
        '
        'fbNyukin_Sheet1
        '
        Me.fbNyukin_Sheet1.Reset()
        Me.fbNyukin_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbNyukin_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbNyukin_Sheet1.ColumnCount = 10
        Me.fbNyukin_Sheet1.RowCount = 0
        Me.fbNyukin_Sheet1.ActiveColumnIndex = -1
        Me.fbNyukin_Sheet1.ActiveRowIndex = -1
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "相手先ｺｰﾄﾞ"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "相手先名"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "入金予定日"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "入金日"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "入金額"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "請求日"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "入力日"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "EXAS請求NO"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "請求依頼番号（入力摘要）"
        Me.fbNyukin_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "　"
        Me.fbNyukin_Sheet1.ColumnHeader.Rows.Get(0).Height = 34.0!
        Me.fbNyukin_Sheet1.Columns.Get(0).Label = "相手先ｺｰﾄﾞ"
        Me.fbNyukin_Sheet1.Columns.Get(0).Width = 95.0!
        Me.fbNyukin_Sheet1.Columns.Get(1).Label = "相手先名"
        Me.fbNyukin_Sheet1.Columns.Get(1).Width = 235.0!
        Me.fbNyukin_Sheet1.Columns.Get(2).Label = "入金予定日"
        Me.fbNyukin_Sheet1.Columns.Get(2).Width = 95.0!
        Me.fbNyukin_Sheet1.Columns.Get(3).Label = "入金日"
        Me.fbNyukin_Sheet1.Columns.Get(3).Width = 95.0!
        Me.fbNyukin_Sheet1.Columns.Get(4).Label = "入金額"
        Me.fbNyukin_Sheet1.Columns.Get(4).Width = 95.0!
        Me.fbNyukin_Sheet1.Columns.Get(5).Label = "請求日"
        Me.fbNyukin_Sheet1.Columns.Get(5).Width = 95.0!
        Me.fbNyukin_Sheet1.Columns.Get(6).Label = "入力日"
        Me.fbNyukin_Sheet1.Columns.Get(6).Width = 95.0!
        Me.fbNyukin_Sheet1.Columns.Get(7).Label = "EXAS請求NO"
        Me.fbNyukin_Sheet1.Columns.Get(7).Width = 95.0!
        Me.fbNyukin_Sheet1.Columns.Get(8).Label = "請求依頼番号（入力摘要）"
        Me.fbNyukin_Sheet1.Columns.Get(8).Width = 93.0!
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "解除"
        Me.fbNyukin_Sheet1.Columns.Get(9).CellType = ButtonCellType1
        Me.fbNyukin_Sheet1.Columns.Get(9).Label = "　"
        Me.fbNyukin_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbNyukin_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'dtpNyukinDt
        '
        Me.dtpNyukinDt.ImeMode = System.Windows.Forms.ImeMode.Alpha
        Me.dtpNyukinDt.Location = New System.Drawing.Point(382, 77)
        Me.dtpNyukinDt.Name = "dtpNyukinDt"
        Me.dtpNyukinDt.Size = New System.Drawing.Size(148, 25)
        Me.dtpNyukinDt.TabIndex = 136
        '
        'EXTZ0207
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1786, 538)
        Me.Controls.Add(Me.dtpNyukinDt)
        Me.Controls.Add(Me.fbNyukin)
        Me.Controls.Add(Me.lblNyukinYoteiDt)
        Me.Controls.Add(Me.lblSeikyuDt)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.lblNyukinTtl)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.btnExasList)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.btnSetNyukinTtl)
        Me.Controls.Add(Me.txtNyukin)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.lblChoseiKin)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.lblAiteCd)
        Me.Controls.Add(Me.lblAiteNm)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.lblSeikyuNaiyo)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.lblSeikyuKin)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.lblTaxKin)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblShokei)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lblKakuteiKin)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblSeikyuNo)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnComplate)
        Me.Controls.Add(Me.Label65)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0207"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　入金登録"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.fbNyukin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbNyukin_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents btnComplate As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents lblSeikyuNo As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblKakuteiKin As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblShokei As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblTaxKin As System.Windows.Forms.Label
    Friend WithEvents lblSeikyuNaiyo As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lblSeikyuKin As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents lblAiteNm As System.Windows.Forms.Label
    Friend WithEvents lblAiteCd As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rdoNyukin2 As System.Windows.Forms.RadioButton
    Friend WithEvents rdoNyukin1 As System.Windows.Forms.RadioButton
    Friend WithEvents lblChoseiKin As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtNyukin As System.Windows.Forms.TextBox
    Friend WithEvents btnSetNyukinTtl As System.Windows.Forms.Button
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents btnExasList As System.Windows.Forms.Button
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents lblNyukinTtl As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents lblNyukinYoteiDt As System.Windows.Forms.Label
    Friend WithEvents lblSeikyuDt As System.Windows.Forms.Label
    Friend WithEvents fbNyukin As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbNyukin_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents dtpNyukinDt As Common.DateTimePickerEx

End Class
