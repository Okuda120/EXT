<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0103
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0103))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.txtRiyoNm = New System.Windows.Forms.TextBox()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.txtRiyoNm_Kana = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rdoStudio = New System.Windows.Forms.RadioButton()
        Me.rdoTheatre = New System.Windows.Forms.RadioButton()
        Me.btnRiyoshaSearch = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtShiyoTsuki_To = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtShiyoNen_To = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtShiyoTsuki_From = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtShiyoNen_From = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.vwDayUriageTheatre = New FarPoint.Win.Spread.FpSpread()
        Me.vwDayUriageTheatre_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwSeikyuChoseiTheatre = New FarPoint.Win.Spread.FpSpread()
        Me.vwSeikyuChoseiTheatre_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwDayUriageStudio = New FarPoint.Win.Spread.FpSpread()
        Me.vwDayUriageStudio_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwSeikyuChoseiStudio = New FarPoint.Win.Spread.FpSpread()
        Me.vwSeikyuChoseiStudio_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.vwDayUriageTheatre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwDayUriageTheatre_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwSeikyuChoseiTheatre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwSeikyuChoseiTheatre_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwDayUriageStudio, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwDayUriageStudio_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwSeikyuChoseiStudio, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwSeikyuChoseiStudio_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(21, 66)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(85, 21)
        Me.Label65.TabIndex = 2
        Me.Label65.Text = "使用月"
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label64.Location = New System.Drawing.Point(21, 131)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(82, 13)
        Me.Label64.TabIndex = 14
        Me.Label64.Text = "利用者名(ｶﾅ)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtRiyoNm
        '
        Me.txtRiyoNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtRiyoNm.Location = New System.Drawing.Point(170, 96)
        Me.txtRiyoNm.Name = "txtRiyoNm"
        Me.txtRiyoNm.Size = New System.Drawing.Size(306, 20)
        Me.txtRiyoNm.TabIndex = 13
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label54.Location = New System.Drawing.Point(21, 99)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(59, 13)
        Me.Label54.TabIndex = 12
        Me.Label54.Text = "利用者名"
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label53.Location = New System.Drawing.Point(317, 66)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(20, 13)
        Me.Label53.TabIndex = 7
        Me.Label53.Text = "～"
        '
        'txtRiyoNm_Kana
        '
        Me.txtRiyoNm_Kana.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtRiyoNm_Kana.Location = New System.Drawing.Point(170, 131)
        Me.txtRiyoNm_Kana.Name = "txtRiyoNm_Kana"
        Me.txtRiyoNm_Kana.Size = New System.Drawing.Size(306, 20)
        Me.txtRiyoNm_Kana.TabIndex = 15
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rdoStudio)
        Me.Panel1.Controls.Add(Me.rdoTheatre)
        Me.Panel1.Location = New System.Drawing.Point(153, 30)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(222, 22)
        Me.Panel1.TabIndex = 1
        '
        'rdoStudio
        '
        Me.rdoStudio.AutoSize = True
        Me.rdoStudio.Location = New System.Drawing.Point(99, 3)
        Me.rdoStudio.Name = "rdoStudio"
        Me.rdoStudio.Size = New System.Drawing.Size(65, 17)
        Me.rdoStudio.TabIndex = 1
        Me.rdoStudio.Text = "スタジオ"
        Me.rdoStudio.UseVisualStyleBackColor = True
        '
        'rdoTheatre
        '
        Me.rdoTheatre.AutoSize = True
        Me.rdoTheatre.Checked = True
        Me.rdoTheatre.Location = New System.Drawing.Point(3, 3)
        Me.rdoTheatre.Name = "rdoTheatre"
        Me.rdoTheatre.Size = New System.Drawing.Size(66, 17)
        Me.rdoTheatre.TabIndex = 0
        Me.rdoTheatre.TabStop = True
        Me.rdoTheatre.Text = "シアター"
        Me.rdoTheatre.UseVisualStyleBackColor = True
        '
        'btnRiyoshaSearch
        '
        Me.btnRiyoshaSearch.BackgroundImage = Global.EXTZ.My.Resources.Resources.マスタ検索
        Me.btnRiyoshaSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnRiyoshaSearch.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnRiyoshaSearch.Location = New System.Drawing.Point(482, 124)
        Me.btnRiyoshaSearch.Name = "btnRiyoshaSearch"
        Me.btnRiyoshaSearch.Size = New System.Drawing.Size(36, 27)
        Me.btnRiyoshaSearch.TabIndex = 16
        Me.btnRiyoshaSearch.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox5.Controls.Add(Me.Label6)
        Me.GroupBox5.Controls.Add(Me.txtShiyoTsuki_To)
        Me.GroupBox5.Controls.Add(Me.Label7)
        Me.GroupBox5.Controls.Add(Me.txtShiyoNen_To)
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Controls.Add(Me.txtShiyoTsuki_From)
        Me.GroupBox5.Controls.Add(Me.Label11)
        Me.GroupBox5.Controls.Add(Me.txtShiyoNen_From)
        Me.GroupBox5.Controls.Add(Me.btnSearch)
        Me.GroupBox5.Controls.Add(Me.btnRiyoshaSearch)
        Me.GroupBox5.Controls.Add(Me.Panel1)
        Me.GroupBox5.Controls.Add(Me.txtRiyoNm_Kana)
        Me.GroupBox5.Controls.Add(Me.Label53)
        Me.GroupBox5.Controls.Add(Me.Label54)
        Me.GroupBox5.Controls.Add(Me.txtRiyoNm)
        Me.GroupBox5.Controls.Add(Me.Label64)
        Me.GroupBox5.Controls.Add(Me.Label65)
        Me.GroupBox5.Controls.Add(Me.Label66)
        Me.GroupBox5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(799, 170)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(482, 67)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 12)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "月"
        '
        'txtShiyoTsuki_To
        '
        Me.txtShiyoTsuki_To.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtShiyoTsuki_To.Location = New System.Drawing.Point(440, 64)
        Me.txtShiyoTsuki_To.MaxLength = 2
        Me.txtShiyoTsuki_To.Name = "txtShiyoTsuki_To"
        Me.txtShiyoTsuki_To.Size = New System.Drawing.Size(36, 19)
        Me.txtShiyoTsuki_To.TabIndex = 10
        Me.txtShiyoTsuki_To.Text = "5"
        Me.txtShiyoTsuki_To.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.Location = New System.Drawing.Point(417, 67)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 12)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "年"
        '
        'txtShiyoNen_To
        '
        Me.txtShiyoNen_To.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtShiyoNen_To.Location = New System.Drawing.Point(352, 64)
        Me.txtShiyoNen_To.MaxLength = 4
        Me.txtShiyoNen_To.Name = "txtShiyoNen_To"
        Me.txtShiyoNen_To.Size = New System.Drawing.Size(59, 19)
        Me.txtShiyoNen_To.TabIndex = 8
        Me.txtShiyoNen_To.Text = "2015"
        Me.txtShiyoNen_To.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(284, 67)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 12)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "月"
        '
        'txtShiyoTsuki_From
        '
        Me.txtShiyoTsuki_From.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtShiyoTsuki_From.Location = New System.Drawing.Point(242, 64)
        Me.txtShiyoTsuki_From.MaxLength = 2
        Me.txtShiyoTsuki_From.Name = "txtShiyoTsuki_From"
        Me.txtShiyoTsuki_From.Size = New System.Drawing.Size(36, 19)
        Me.txtShiyoTsuki_From.TabIndex = 5
        Me.txtShiyoTsuki_From.Text = "5"
        Me.txtShiyoTsuki_From.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label11.Location = New System.Drawing.Point(219, 67)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(17, 12)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "年"
        '
        'txtShiyoNen_From
        '
        Me.txtShiyoNen_From.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtShiyoNen_From.Location = New System.Drawing.Point(154, 64)
        Me.txtShiyoNen_From.MaxLength = 4
        Me.txtShiyoNen_From.Name = "txtShiyoNen_From"
        Me.txtShiyoNen_From.Size = New System.Drawing.Size(59, 19)
        Me.txtShiyoNen_From.TabIndex = 3
        Me.txtShiyoNen_From.Text = "2015"
        Me.txtShiyoNen_From.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(577, 129)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(96, 23)
        Me.btnSearch.TabIndex = 17
        Me.btnSearch.Text = "検索"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'Label66
        '
        Me.Label66.AutoSize = True
        Me.Label66.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label66.Location = New System.Drawing.Point(21, 34)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(59, 13)
        Me.Label66.TabIndex = 0
        Me.Label66.Text = "検索対象"
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(785, 914)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(175, 36)
        Me.btnBack.TabIndex = 7
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 200)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "日別料金一覧"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 699)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(159, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "請求時の調整額(グロス調整)"
        '
        'vwDayUriageTheatre
        '
        Me.vwDayUriageTheatre.AccessibleDescription = "vwDayUriageTheatre, Sheet1, Row 0, Column 0, T00001"
        Me.vwDayUriageTheatre.Location = New System.Drawing.Point(12, 219)
        Me.vwDayUriageTheatre.Name = "vwDayUriageTheatre"
        Me.vwDayUriageTheatre.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwDayUriageTheatre_Sheet1})
        Me.vwDayUriageTheatre.Size = New System.Drawing.Size(1654, 458)
        Me.vwDayUriageTheatre.TabIndex = 3
        '
        'vwDayUriageTheatre_Sheet1
        '
        Me.vwDayUriageTheatre_Sheet1.Reset()
        Me.vwDayUriageTheatre_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwDayUriageTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwDayUriageTheatre_Sheet1.ColumnCount = 24
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 0).Value = "T00001"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 1).ParseFormatString = "yyyy/M/d"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 1).Value = New Date(2015, 4, 1, 0, 0, 0, 0)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 2).Value = "音楽祭"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 3).Value = "エンターテイメント株式会社"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 4).Value = "一般"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 5).Value = "着席"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 6).Value = "音楽"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 7).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 7).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 7).ParseFormatString = "n"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 7).Value = 1000000
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 8).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 8).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 8).ParseFormatString = "n"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 8).Value = 300000
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 9).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 9).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 9).ParseFormatString = "n"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 9).Value = -100000
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 10).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 10).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 10).ParseFormatString = "n"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 10).Value = 1200000
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 11).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 11).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 11).ParseFormatString = "n"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 11).Value = 50000
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 23).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 23).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 23).ParseFormatString = "n"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(0, 23).Value = 1250000
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 0).Value = "T00001"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 1).ParseFormatString = "yyyy/M/d"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 1).Value = New Date(2015, 4, 2, 0, 0, 0, 0)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 2).Value = "音楽祭"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 3).Value = "エンターテイメント株式会社"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 4).Value = "社内"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 5).Value = "スタンディング"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(1, 6).Value = "演劇"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 0).Value = "T00002"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 1).ParseFormatString = "yyyy/M/d"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 1).Value = New Date(2015, 4, 3, 0, 0, 0, 0)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 2).Value = "とっとちゃんシアター"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 3).Value = "株式会社黒船団"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 5).Value = "変則"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(2, 6).Value = "演芸"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(3, 0).Value = "T00003"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(3, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(3, 1).ParseFormatString = "yyyy/M/d"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(3, 1).Value = New Date(2015, 4, 6, 0, 0, 0, 0)
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(3, 5).Value = "催事"
        Me.vwDayUriageTheatre_Sheet1.Cells.Get(3, 6).Value = "ビジネス"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "予約番号"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "利用日"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "催事名"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用者名"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "利用" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "形状"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "催事分類"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "利用料(A)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "使用料(B)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "調整額(C)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "基本売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(D)=(A+B+C)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "001" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾄﾞﾘﾝｸ現金(E)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "002" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾜﾝﾄﾞﾘﾝｸ(F)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "003" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ｺｲﾝﾛｯｶｰ(G)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "004" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "雑収入(H)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "006" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(I)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 16).Value = "007" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(J)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 17).Value = "008" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(K)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 18).Value = "009" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(L)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 19).Value = "900" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICA現金(M)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 20).Value = "901" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICAｺｲﾝ(N)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 21).Value = "902-950" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "その他(O)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 22).Value = "現金合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(P)=(E～O)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Cells.Get(0, 23).Value = "総売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Q)=(D+P)"
        Me.vwDayUriageTheatre_Sheet1.ColumnHeader.Rows.Get(0).Height = 56.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(0).Label = "予約番号"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(0).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(0).Width = 68.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(1).Label = "利用日"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(1).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(1).Width = 77.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(2).Label = "催事名"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(2).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(2).Width = 149.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(3).Label = "利用者名"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(3).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(3).Width = 167.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(4).Label = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(4).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(4).Width = 61.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(5).Label = "利用" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "形状"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(5).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(5).Width = 82.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(6).Label = "催事分類"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(6).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(6).Width = 80.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(7).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(7).Label = "利用料(A)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(7).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(7).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(8).Label = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "使用料(B)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(8).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(8).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(9).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(9).Label = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "調整額(C)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(9).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(9).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(10).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(10).Label = "基本売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(D)=(A+B+C)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(10).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(10).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(11).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(11).Label = "001" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾄﾞﾘﾝｸ現金(E)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(11).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(11).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(12).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(12).Label = "002" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾜﾝﾄﾞﾘﾝｸ(F)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(12).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(12).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(13).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(13).Label = "003" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ｺｲﾝﾛｯｶｰ(G)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(13).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(13).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(14).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(14).Label = "004" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "雑収入(H)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(14).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(14).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(14).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(15).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(15).Label = "006" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(I)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(15).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(15).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(15).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(16).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(16).Label = "007" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(J)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(16).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(16).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(16).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(17).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(17).Label = "008" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(K)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(17).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(17).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(17).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(18).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(18).Label = "009" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(L)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(18).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(18).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(18).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(19).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(19).Label = "900" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICA現金(M)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(19).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(19).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(19).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(20).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(20).Label = "901" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICAｺｲﾝ(N)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(20).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(20).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(20).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(21).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(21).Label = "902-950" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "その他(O)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(21).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(21).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(21).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(22).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(22).Label = "現金合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(P)=(E～O)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(22).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(22).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(22).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(23).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(23).Label = "総売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Q)=(D+P)"
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(23).Locked = True
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(23).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageTheatre_Sheet1.Columns.Get(23).Width = 71.0!
        Me.vwDayUriageTheatre_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwDayUriageTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwSeikyuChoseiTheatre
        '
        Me.vwSeikyuChoseiTheatre.AccessibleDescription = "vwSeikyuChoseiTheatre, Sheet1, Row 0, Column 6, 音楽"
        Me.vwSeikyuChoseiTheatre.Location = New System.Drawing.Point(12, 717)
        Me.vwSeikyuChoseiTheatre.Name = "vwSeikyuChoseiTheatre"
        Me.vwSeikyuChoseiTheatre.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwSeikyuChoseiTheatre_Sheet1})
        Me.vwSeikyuChoseiTheatre.Size = New System.Drawing.Size(990, 177)
        Me.vwSeikyuChoseiTheatre.TabIndex = 6
        '
        'vwSeikyuChoseiTheatre_Sheet1
        '
        Me.vwSeikyuChoseiTheatre_Sheet1.Reset()
        Me.vwSeikyuChoseiTheatre_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwSeikyuChoseiTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnCount = 10
        Me.vwSeikyuChoseiTheatre_Sheet1.ActiveColumnIndex = 6
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 0).Value = "T00001"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 1).ParseFormatString = "yyyy/M/d"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 1).Value = New Date(2015, 4, 1, 0, 0, 0, 0)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 2).Value = "音楽祭"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 3).Value = "エンターテイメント株式会社"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 4).Value = "一般"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 5).Value = "着席"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(0, 6).Value = "音楽"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 0).Value = "T00001"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 1).ParseFormatString = "yyyy/M/d"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 1).Value = New Date(2015, 4, 2, 0, 0, 0, 0)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 2).Value = "音楽祭"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 3).Value = "エンターテイメント株式会社"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 4).Value = "社内"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 5).Value = "スタンディング"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(1, 6).Value = "演劇"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 0).Value = "T00002"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 1).ParseFormatString = "yyyy/M/d"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 1).Value = New Date(2015, 4, 3, 0, 0, 0, 0)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 2).Value = "とっとちゃんシアター"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 3).Value = "株式会社黒船団"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 5).Value = "変則"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(2, 6).Value = "演芸"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(3, 0).Value = "T00003"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(3, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(3, 1).ParseFormatString = "yyyy/M/d"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(3, 1).Value = New Date(2015, 4, 6, 0, 0, 0, 0)
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(3, 5).Value = "催事"
        Me.vwSeikyuChoseiTheatre_Sheet1.Cells.Get(3, 6).Value = "ビジネス"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "予約番号"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "利用日"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "催事名"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用者名"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "利用" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "形状"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "催事分類"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "請求依頼番号"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "請求内容"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "調整額"
        Me.vwSeikyuChoseiTheatre_Sheet1.ColumnHeader.Rows.Get(0).Height = 56.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(0).Label = "予約番号"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(0).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(0).Width = 68.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(1).Label = "利用日"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(1).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(1).Width = 77.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(2).Label = "催事名"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(2).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(2).Width = 149.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(3).Label = "利用者名"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(3).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(3).Width = 167.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(4).Label = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(4).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(4).Width = 61.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(5).Label = "利用" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "形状"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(5).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(5).Width = 82.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(6).Label = "催事分類"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(6).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(6).Width = 80.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(7).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(7).Label = "請求依頼番号"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(7).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(7).Width = 71.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(8).Label = "請求内容"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(8).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(8).Width = 104.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(9).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(9).Label = "調整額"
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(9).Locked = True
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiTheatre_Sheet1.Columns.Get(9).Width = 71.0!
        Me.vwSeikyuChoseiTheatre_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwSeikyuChoseiTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwDayUriageStudio
        '
        Me.vwDayUriageStudio.AccessibleDescription = "FpSpread3, Sheet1, Row 0, Column 0, S00001"
        Me.vwDayUriageStudio.Location = New System.Drawing.Point(12, 219)
        Me.vwDayUriageStudio.Name = "vwDayUriageStudio"
        Me.vwDayUriageStudio.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwDayUriageStudio_Sheet1})
        Me.vwDayUriageStudio.Size = New System.Drawing.Size(1654, 462)
        Me.vwDayUriageStudio.TabIndex = 2
        '
        'vwDayUriageStudio_Sheet1
        '
        Me.vwDayUriageStudio_Sheet1.Reset()
        Me.vwDayUriageStudio_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwDayUriageStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwDayUriageStudio_Sheet1.ColumnCount = 23
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 0).Value = "S00001"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 1).ParseFormatString = "yyyy/M/d"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 1).Value = New Date(2015, 4, 1, 0, 0, 0, 0)
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 2).Value = "爆弾岩"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 3).Value = "ミュージックオン株式会社"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 4).Value = "一般"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 5).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 5).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 5).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 5).Value = 201
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 6).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 6).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 6).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 6).Value = 50000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 7).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 7).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 7).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 7).Value = 20000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 8).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 8).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 8).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 8).Value = -10000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 9).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 9).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 9).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 9).Value = 60000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 10).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 10).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 10).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 10).Value = 50000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 11).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 11).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 11).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 11).Value = 12000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 12).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 12).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 12).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 12).Value = 72000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 22).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 22).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 22).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(0, 22).Value = 1250000
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 0).Value = "S00002"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 1).ParseFormatString = "yyyy/M/d"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 1).Value = New Date(2015, 4, 2, 0, 0, 0, 0)
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 2).Value = "ノベルノベル"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 3).Value = "池田　健人"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 4).Value = "社内"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 5).ParseFormatInfo = CType(cultureInfo.NumberFormat.Clone, System.Globalization.NumberFormatInfo)
        CType(Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 5).ParseFormatInfo, System.Globalization.NumberFormatInfo).NumberDecimalDigits = 0
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 5).ParseFormatString = "n"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(1, 5).Value = 202
        Me.vwDayUriageStudio_Sheet1.Cells.Get(2, 0).Value = "S00003"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(2, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwDayUriageStudio_Sheet1.Cells.Get(2, 1).ParseFormatString = "yyyy/M/d"
        Me.vwDayUriageStudio_Sheet1.Cells.Get(2, 1).Value = New Date(2015, 4, 3, 0, 0, 0, 0)
        Me.vwDayUriageStudio_Sheet1.Cells.Get(2, 5).Value = "H/L"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "予約番号"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "利用日"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "アーティスト名"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用者名"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "スタジオ"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "利用料(A)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "使用料(B)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "調整額(C)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "基本売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(D)=(A+B+C)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "001" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾄﾞﾘﾝｸ現金(E)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "002" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾜﾝﾄﾞﾘﾝｸ(F)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "003" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ｺｲﾝﾛｯｶｰ(G)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "004" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "雑収入(H)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "006" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(I)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "007" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(J)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 16).Value = "008" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(K)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 17).Value = "009" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(L)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 18).Value = "900" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICA現金(M)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 19).Value = "901" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICAｺｲﾝ(N)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 20).Value = "902-950" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "その他(O)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 21).Value = "現金合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(P)=(E～O)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Cells.Get(0, 22).Value = "総売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Q)=(D+P)"
        Me.vwDayUriageStudio_Sheet1.ColumnHeader.Rows.Get(0).Height = 56.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(0).Label = "予約番号"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(0).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(0).Width = 68.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(1).Label = "利用日"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(1).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(1).Width = 77.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(2).Label = "アーティスト名"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(2).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(2).Width = 149.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(3).Label = "利用者名"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(3).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(3).Width = 167.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(4).Label = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwDayUriageStudio_Sheet1.Columns.Get(4).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(4).Width = 61.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(5).Label = "スタジオ"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(5).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(5).Width = 82.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(6).Label = "利用料(A)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(6).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(6).Width = 68.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(7).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(7).Label = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "使用料(B)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(7).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(7).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(8).Label = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "調整額(C)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(8).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(8).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(9).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(9).Label = "基本売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(D)=(A+B+C)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(9).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(9).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(10).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(10).Label = "001" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾄﾞﾘﾝｸ現金(E)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(10).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(10).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(11).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(11).Label = "002" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ﾜﾝﾄﾞﾘﾝｸ(F)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(11).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(11).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(12).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(12).Label = "003" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ｺｲﾝﾛｯｶｰ(G)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(12).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(12).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(13).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(13).Label = "004" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "雑収入(H)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(13).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(13).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(14).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(14).Label = "006" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(I)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(14).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(14).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(14).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(15).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(15).Label = "007" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(J)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(15).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(15).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(15).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(16).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(16).Label = "008" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(K)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwDayUriageStudio_Sheet1.Columns.Get(16).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(16).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(16).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(17).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(17).Label = "009" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "空き(L)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwDayUriageStudio_Sheet1.Columns.Get(17).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(17).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(17).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(18).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(18).Label = "900" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICA現金(M)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(18).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(18).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(18).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(19).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(19).Label = "901" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SUICAｺｲﾝ(N)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(19).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(19).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(19).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(20).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(20).Label = "902-950" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "その他(O)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(20).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(20).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(20).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(21).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(21).Label = "現金合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(P)=(E～O)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(21).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(21).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(21).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.Columns.Get(22).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwDayUriageStudio_Sheet1.Columns.Get(22).Label = "総売上合計" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Q)=(D+P)"
        Me.vwDayUriageStudio_Sheet1.Columns.Get(22).Locked = True
        Me.vwDayUriageStudio_Sheet1.Columns.Get(22).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwDayUriageStudio_Sheet1.Columns.Get(22).Width = 71.0!
        Me.vwDayUriageStudio_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwDayUriageStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwSeikyuChoseiStudio
        '
        Me.vwSeikyuChoseiStudio.AccessibleDescription = "FpSpread4, Sheet1, Row 0, Column 0, S00001"
        Me.vwSeikyuChoseiStudio.Location = New System.Drawing.Point(12, 717)
        Me.vwSeikyuChoseiStudio.Name = "vwSeikyuChoseiStudio"
        Me.vwSeikyuChoseiStudio.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwSeikyuChoseiStudio_Sheet1})
        Me.vwSeikyuChoseiStudio.Size = New System.Drawing.Size(990, 177)
        Me.vwSeikyuChoseiStudio.TabIndex = 5
        '
        'vwSeikyuChoseiStudio_Sheet1
        '
        Me.vwSeikyuChoseiStudio_Sheet1.Reset()
        Me.vwSeikyuChoseiStudio_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwSeikyuChoseiStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnCount = 9
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(0, 0).Value = "S00001"
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(0, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(0, 1).ParseFormatString = "yyyy/M/d"
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(0, 1).Value = New Date(2015, 4, 1, 0, 0, 0, 0)
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(0, 3).Value = "エンターテイメント株式会社"
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(0, 4).Value = "一般"
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(1, 0).Value = "S00002"
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(1, 1).ParseFormatInfo = CType(cultureInfo.DateTimeFormat.Clone, System.Globalization.DateTimeFormatInfo)
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(1, 1).ParseFormatString = "yyyy/M/d"
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(1, 1).Value = New Date(2015, 4, 2, 0, 0, 0, 0)
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(1, 3).Value = "エンターテイメント株式会社"
        Me.vwSeikyuChoseiStudio_Sheet1.Cells.Get(1, 4).Value = "社内"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "予約番号"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "利用日"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "アーティスト名"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用者名"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "スタジオ"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "請求依頼番号"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "請求内容"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "調整額"
        Me.vwSeikyuChoseiStudio_Sheet1.ColumnHeader.Rows.Get(0).Height = 56.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(0).Label = "予約番号"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(0).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(0).Width = 68.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(1).Label = "利用日"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(1).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(1).Width = 77.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(2).Label = "アーティスト名"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(2).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(2).Width = 149.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(3).Label = "利用者名"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(3).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(3).Width = 167.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(4).Label = "貸出" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "種別"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(4).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(4).Width = 61.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(5).Label = "スタジオ"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(5).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(5).Width = 82.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(6).Label = "請求依頼番号"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(6).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(6).Width = 71.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(7).Label = "請求内容"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(7).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(7).Width = 104.0!
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(8).Label = "調整額"
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(8).Locked = True
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuChoseiStudio_Sheet1.Columns.Get(8).Width = 71.0!
        Me.vwSeikyuChoseiStudio_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwSeikyuChoseiStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(20, 903)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(494, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "※表のデータは、コピーしたい部分を選択し、キーボードの「Ctrl」＋「C」を押すとコピーできます。"
        '
        'EXTZ0103
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1684, 962)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.vwSeikyuChoseiStudio)
        Me.Controls.Add(Me.vwSeikyuChoseiTheatre)
        Me.Controls.Add(Me.vwDayUriageTheatre)
        Me.Controls.Add(Me.vwDayUriageStudio)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0103"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　日別売上一覧"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.vwDayUriageTheatre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwDayUriageTheatre_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwSeikyuChoseiTheatre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwSeikyuChoseiTheatre_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwDayUriageStudio, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwDayUriageStudio_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwSeikyuChoseiStudio, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwSeikyuChoseiStudio_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents txtRiyoNm As System.Windows.Forms.TextBox
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents txtRiyoNm_Kana As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rdoStudio As System.Windows.Forms.RadioButton
    Friend WithEvents rdoTheatre As System.Windows.Forms.RadioButton
    Friend WithEvents btnRiyoshaSearch As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents vwDayUriageTheatre As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwDayUriageTheatre_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwSeikyuChoseiTheatre As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwSeikyuChoseiTheatre_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwDayUriageStudio As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwDayUriageStudio_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwSeikyuChoseiStudio As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwSeikyuChoseiStudio_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtShiyoTsuki_To As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtShiyoNen_To As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtShiyoTsuki_From As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtShiyoNen_From As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
