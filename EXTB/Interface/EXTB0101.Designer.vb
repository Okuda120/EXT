﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTB0101
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
        Dim ButtonCellType2 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim ButtonCellType3 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim ButtonCellType4 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim ButtonCellType5 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim CheckBoxCellType1 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim CheckBoxCellType2 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTB0101))
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.lblYear = New System.Windows.Forms.Label()
        Me.txtMonth = New System.Windows.Forms.TextBox()
        Me.lblMonth = New System.Windows.Forms.Label()
        Me.lblKariYoyakuMikakunin = New System.Windows.Forms.Label()
        Me.lblKariyoyaku = New System.Windows.Forms.Label()
        Me.lblSeisikiYoyaku = New System.Windows.Forms.Label()
        Me.lblSeisikiYoyakuKanryo = New System.Windows.Forms.Label()
        Me.lblKyukanbi = New System.Windows.Forms.Label()
        Me.lblMaintenance = New System.Windows.Forms.Label()
        Me.btnRegistKariYoyaku = New System.Windows.Forms.Button()
        Me.btnRegistCancelWait = New System.Windows.Forms.Button()
        Me.btnChangeMaintebi = New System.Windows.Forms.Button()
        Me.btnChangeEigyobi = New System.Windows.Forms.Button()
        Me.btnToMenu = New System.Windows.Forms.Button()
        Me.btnSearchPrevMonth = New System.Windows.Forms.Button()
        Me.btnSearchNextMonth = New System.Windows.Forms.Button()
        Me.grpCancelWait = New System.Windows.Forms.GroupBox()
        Me.vwCancelWait = New FarPoint.Win.Spread.FpSpread()
        Me.vwCancelWait_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.grpCancelDone = New System.Windows.Forms.GroupBox()
        Me.vwCancelDone = New FarPoint.Win.Spread.FpSpread()
        Me.vwCancelDone_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.btnUserList = New System.Windows.Forms.Button()
        Me.vwCalandarFirst_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwCalandarFirst = New FarPoint.Win.Spread.FpSpread()
        Me.vwCalendarSecond = New FarPoint.Win.Spread.FpSpread()
        Me.vwCalendarSecond_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.btnChangeKyukanbi = New System.Windows.Forms.Button()
        Me.btnChangeHoliday = New System.Windows.Forms.Button()
        Me.btnChangeWeekday = New System.Windows.Forms.Button()
        Me.grpCancelWait.SuspendLayout()
        CType(Me.vwCancelWait, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwCancelWait_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCancelDone.SuspendLayout()
        CType(Me.vwCancelDone, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwCancelDone_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwCalandarFirst_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwCalandarFirst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwCalendarSecond, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwCalendarSecond_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(331, 30)
        Me.btnSearch.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(165, 36)
        Me.btnSearch.TabIndex = 4
        Me.btnSearch.Text = "表示"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'txtYear
        '
        Me.txtYear.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtYear.Location = New System.Drawing.Point(33, 31)
        Me.txtYear.Margin = New System.Windows.Forms.Padding(4)
        Me.txtYear.MaxLength = 4
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(124, 34)
        Me.txtYear.TabIndex = 0
        Me.txtYear.Text = "2015"
        Me.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.BackColor = System.Drawing.Color.Transparent
        Me.lblYear.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblYear.Location = New System.Drawing.Point(167, 38)
        Me.lblYear.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(35, 24)
        Me.lblYear.TabIndex = 1
        Me.lblYear.Text = "年"
        '
        'txtMonth
        '
        Me.txtMonth.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtMonth.Location = New System.Drawing.Point(213, 31)
        Me.txtMonth.Margin = New System.Windows.Forms.Padding(4)
        Me.txtMonth.MaxLength = 2
        Me.txtMonth.Name = "txtMonth"
        Me.txtMonth.Size = New System.Drawing.Size(61, 34)
        Me.txtMonth.TabIndex = 2
        Me.txtMonth.Text = "5"
        Me.txtMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblMonth
        '
        Me.lblMonth.AutoSize = True
        Me.lblMonth.BackColor = System.Drawing.Color.Transparent
        Me.lblMonth.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblMonth.Location = New System.Drawing.Point(284, 38)
        Me.lblMonth.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMonth.Name = "lblMonth"
        Me.lblMonth.Size = New System.Drawing.Size(35, 24)
        Me.lblMonth.TabIndex = 3
        Me.lblMonth.Text = "月"
        '
        'lblKariYoyakuMikakunin
        '
        Me.lblKariYoyakuMikakunin.BackColor = System.Drawing.Color.White
        Me.lblKariYoyakuMikakunin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKariYoyakuMikakunin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblKariYoyakuMikakunin.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblKariYoyakuMikakunin.Location = New System.Drawing.Point(784, 31)
        Me.lblKariYoyakuMikakunin.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblKariYoyakuMikakunin.Name = "lblKariYoyakuMikakunin"
        Me.lblKariYoyakuMikakunin.Size = New System.Drawing.Size(236, 38)
        Me.lblKariYoyakuMikakunin.TabIndex = 6
        Me.lblKariYoyakuMikakunin.Text = "仮予約"
        Me.lblKariYoyakuMikakunin.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblKariyoyaku
        '
        Me.lblKariyoyaku.BackColor = System.Drawing.Color.FromArgb(CType(CType(127, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblKariyoyaku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKariyoyaku.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblKariyoyaku.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblKariyoyaku.Location = New System.Drawing.Point(1028, 31)
        Me.lblKariyoyaku.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblKariyoyaku.Name = "lblKariyoyaku"
        Me.lblKariyoyaku.Size = New System.Drawing.Size(236, 38)
        Me.lblKariyoyaku.TabIndex = 7
        Me.lblKariyoyaku.Text = "決定"
        Me.lblKariyoyaku.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSeisikiYoyaku
        '
        Me.lblSeisikiYoyaku.BackColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.lblSeisikiYoyaku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeisikiYoyaku.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSeisikiYoyaku.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSeisikiYoyaku.Location = New System.Drawing.Point(1272, 31)
        Me.lblSeisikiYoyaku.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSeisikiYoyaku.Name = "lblSeisikiYoyaku"
        Me.lblSeisikiYoyaku.Size = New System.Drawing.Size(236, 38)
        Me.lblSeisikiYoyaku.TabIndex = 8
        Me.lblSeisikiYoyaku.Text = "申請受諾済"
        Me.lblSeisikiYoyaku.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSeisikiYoyakuKanryo
        '
        Me.lblSeisikiYoyakuKanryo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.lblSeisikiYoyakuKanryo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSeisikiYoyakuKanryo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSeisikiYoyakuKanryo.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSeisikiYoyakuKanryo.Location = New System.Drawing.Point(1516, 31)
        Me.lblSeisikiYoyakuKanryo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSeisikiYoyakuKanryo.Name = "lblSeisikiYoyakuKanryo"
        Me.lblSeisikiYoyakuKanryo.Size = New System.Drawing.Size(236, 38)
        Me.lblSeisikiYoyakuKanryo.TabIndex = 9
        Me.lblSeisikiYoyakuKanryo.Text = "精算完了"
        Me.lblSeisikiYoyakuKanryo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblKyukanbi
        '
        Me.lblKyukanbi.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.lblKyukanbi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKyukanbi.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblKyukanbi.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblKyukanbi.Location = New System.Drawing.Point(1760, 31)
        Me.lblKyukanbi.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblKyukanbi.Name = "lblKyukanbi"
        Me.lblKyukanbi.Size = New System.Drawing.Size(167, 38)
        Me.lblKyukanbi.TabIndex = 10
        Me.lblKyukanbi.Text = "休館日"
        Me.lblKyukanbi.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMaintenance
        '
        Me.lblMaintenance.BackColor = System.Drawing.Color.Gray
        Me.lblMaintenance.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMaintenance.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblMaintenance.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblMaintenance.Location = New System.Drawing.Point(1935, 31)
        Me.lblMaintenance.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMaintenance.Name = "lblMaintenance"
        Me.lblMaintenance.Size = New System.Drawing.Size(167, 38)
        Me.lblMaintenance.TabIndex = 11
        Me.lblMaintenance.Text = "メンテナンス"
        Me.lblMaintenance.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnRegistKariYoyaku
        '
        Me.btnRegistKariYoyaku.Location = New System.Drawing.Point(35, 1195)
        Me.btnRegistKariYoyaku.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRegistKariYoyaku.Name = "btnRegistKariYoyaku"
        Me.btnRegistKariYoyaku.Size = New System.Drawing.Size(228, 36)
        Me.btnRegistKariYoyaku.TabIndex = 16
        Me.btnRegistKariYoyaku.Text = "選択日を仮予約登録"
        Me.btnRegistKariYoyaku.UseVisualStyleBackColor = True
        '
        'btnRegistCancelWait
        '
        Me.btnRegistCancelWait.Location = New System.Drawing.Point(271, 1195)
        Me.btnRegistCancelWait.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRegistCancelWait.Name = "btnRegistCancelWait"
        Me.btnRegistCancelWait.Size = New System.Drawing.Size(228, 36)
        Me.btnRegistCancelWait.TabIndex = 17
        Me.btnRegistCancelWait.Text = "選択日をｷｬﾝｾﾙ待ち登録"
        Me.btnRegistCancelWait.UseVisualStyleBackColor = True
        '
        'btnChangeMaintebi
        '
        Me.btnChangeMaintebi.Location = New System.Drawing.Point(1432, 1195)
        Me.btnChangeMaintebi.Margin = New System.Windows.Forms.Padding(4)
        Me.btnChangeMaintebi.Name = "btnChangeMaintebi"
        Me.btnChangeMaintebi.Size = New System.Drawing.Size(228, 36)
        Me.btnChangeMaintebi.TabIndex = 21
        Me.btnChangeMaintebi.Text = "選択日をメンテ日とする"
        Me.btnChangeMaintebi.UseVisualStyleBackColor = True
        '
        'btnChangeEigyobi
        '
        Me.btnChangeEigyobi.Location = New System.Drawing.Point(1668, 1195)
        Me.btnChangeEigyobi.Margin = New System.Windows.Forms.Padding(4)
        Me.btnChangeEigyobi.Name = "btnChangeEigyobi"
        Me.btnChangeEigyobi.Size = New System.Drawing.Size(228, 36)
        Me.btnChangeEigyobi.TabIndex = 22
        Me.btnChangeEigyobi.Text = "選択日を営業日とする"
        Me.btnChangeEigyobi.UseVisualStyleBackColor = True
        '
        'btnToMenu
        '
        Me.btnToMenu.Location = New System.Drawing.Point(2213, 1195)
        Me.btnToMenu.Margin = New System.Windows.Forms.Padding(4)
        Me.btnToMenu.Name = "btnToMenu"
        Me.btnToMenu.Size = New System.Drawing.Size(228, 36)
        Me.btnToMenu.TabIndex = 25
        Me.btnToMenu.Text = "メニューへ"
        Me.btnToMenu.UseVisualStyleBackColor = True
        '
        'btnSearchPrevMonth
        '
        Me.btnSearchPrevMonth.Location = New System.Drawing.Point(33, 94)
        Me.btnSearchPrevMonth.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSearchPrevMonth.Name = "btnSearchPrevMonth"
        Me.btnSearchPrevMonth.Size = New System.Drawing.Size(228, 36)
        Me.btnSearchPrevMonth.TabIndex = 12
        Me.btnSearchPrevMonth.Text = "前月を表示"
        Me.btnSearchPrevMonth.UseVisualStyleBackColor = True
        '
        'btnSearchNextMonth
        '
        Me.btnSearchNextMonth.Location = New System.Drawing.Point(1380, 94)
        Me.btnSearchNextMonth.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSearchNextMonth.Name = "btnSearchNextMonth"
        Me.btnSearchNextMonth.Size = New System.Drawing.Size(228, 36)
        Me.btnSearchNextMonth.TabIndex = 13
        Me.btnSearchNextMonth.Text = "次月を表示"
        Me.btnSearchNextMonth.UseVisualStyleBackColor = True
        '
        'grpCancelWait
        '
        Me.grpCancelWait.BackColor = System.Drawing.Color.Transparent
        Me.grpCancelWait.Controls.Add(Me.vwCancelWait)
        Me.grpCancelWait.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.grpCancelWait.Location = New System.Drawing.Point(1916, 96)
        Me.grpCancelWait.Margin = New System.Windows.Forms.Padding(4)
        Me.grpCancelWait.Name = "grpCancelWait"
        Me.grpCancelWait.Padding = New System.Windows.Forms.Padding(4)
        Me.grpCancelWait.Size = New System.Drawing.Size(616, 562)
        Me.grpCancelWait.TabIndex = 23
        Me.grpCancelWait.TabStop = False
        Me.grpCancelWait.Text = "キャンセル待ち一覧"
        '
        'vwCancelWait
        '
        Me.vwCancelWait.AccessibleDescription = "vwCancelWait, Sheet1, Row 0, Column 0, 製品発表会"
        Me.vwCancelWait.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwCancelWait.Location = New System.Drawing.Point(24, 46)
        Me.vwCancelWait.Margin = New System.Windows.Forms.Padding(4)
        Me.vwCancelWait.Name = "vwCancelWait"
        Me.vwCancelWait.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwCancelWait_Sheet1})
        Me.vwCancelWait.Size = New System.Drawing.Size(581, 478)
        Me.vwCancelWait.TabIndex = 0
        Me.vwCancelWait.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        '
        'vwCancelWait_Sheet1
        '
        Me.vwCancelWait_Sheet1.Reset()
        Me.vwCancelWait_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwCancelWait_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwCancelWait_Sheet1.ColumnCount = 7
        Me.vwCancelWait_Sheet1.RowCount = 488
        Me.vwCancelWait_Sheet1.Cells.Get(0, 0).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(0, 0).Value = "製品発表会"
        Me.vwCancelWait_Sheet1.Cells.Get(0, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(0, 1).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(0, 1).Value = "株式会社システムシンク"
        Me.vwCancelWait_Sheet1.Cells.Get(0, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(0, 2).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(0, 2).Value = "5/2"
        Me.vwCancelWait_Sheet1.Cells.Get(0, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(0, 3).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(0, 3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(0, 3).Value = "○"
        Me.vwCancelWait_Sheet1.Cells.Get(0, 3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(0, 4).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(1, 0).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(1, 0).Value = "ＴＳライブ"
        Me.vwCancelWait_Sheet1.Cells.Get(1, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(1, 1).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(1, 1).Value = "サービス株式会社"
        Me.vwCancelWait_Sheet1.Cells.Get(1, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(1, 2).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(1, 2).Value = "5/2"
        Me.vwCancelWait_Sheet1.Cells.Get(1, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(1, 3).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(1, 3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(1, 3).Value = "○"
        Me.vwCancelWait_Sheet1.Cells.Get(1, 3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(1, 4).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(2, 0).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(2, 0).Value = "魔法のダンディズム"
        Me.vwCancelWait_Sheet1.Cells.Get(2, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(2, 1).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(2, 1).Value = "企画プランニング"
        Me.vwCancelWait_Sheet1.Cells.Get(2, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(2, 2).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(2, 2).Value = "5月週末"
        Me.vwCancelWait_Sheet1.Cells.Get(2, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(2, 3).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(2, 3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(2, 3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(2, 4).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(3, 0).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(3, 0).Value = "きらりんＱ"
        Me.vwCancelWait_Sheet1.Cells.Get(3, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(3, 1).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(3, 1).Value = "NPO就職支援ネットワーク"
        Me.vwCancelWait_Sheet1.Cells.Get(3, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(3, 2).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(3, 2).Value = "5/23,30の何れか"
        Me.vwCancelWait_Sheet1.Cells.Get(3, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(3, 3).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Cells.Get(3, 3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(3, 3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelWait_Sheet1.Cells.Get(3, 4).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "催事名"
        Me.vwCancelWait_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "利用者"
        Me.vwCancelWait_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "日程"
        Me.vwCancelWait_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "CAL"
        Me.vwCancelWait_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = " "
        Me.vwCancelWait_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwCancelWait_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "キャンセル待ちNO"
        Me.vwCancelWait_Sheet1.ColumnHeader.Rows.Get(0).Height = 21.0!
        Me.vwCancelWait_Sheet1.Columns.Get(0).Label = "催事名"
        Me.vwCancelWait_Sheet1.Columns.Get(0).Width = 101.0!
        Me.vwCancelWait_Sheet1.Columns.Get(1).Label = "利用者"
        Me.vwCancelWait_Sheet1.Columns.Get(1).Width = 140.0!
        Me.vwCancelWait_Sheet1.Columns.Get(2).Label = "日程"
        Me.vwCancelWait_Sheet1.Columns.Get(2).Width = 95.0!
        Me.vwCancelWait_Sheet1.Columns.Get(3).Label = "CAL"
        Me.vwCancelWait_Sheet1.Columns.Get(3).Width = 38.0!
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "詳細"
        Me.vwCancelWait_Sheet1.Columns.Get(4).CellType = ButtonCellType1
        Me.vwCancelWait_Sheet1.Columns.Get(4).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelWait_Sheet1.Columns.Get(4).Label = " "
        Me.vwCancelWait_Sheet1.Columns.Get(4).Width = 38.0!
        ButtonCellType2.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType2.Text = "日付未定キャンセル待ち登録"
        Me.vwCancelWait_Sheet1.Columns.Get(5).CellType = ButtonCellType2
        Me.vwCancelWait_Sheet1.Columns.Get(5).Label = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwCancelWait_Sheet1.Columns.Get(5).Width = 200.0!
        Me.vwCancelWait_Sheet1.Columns.Get(6).Label = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "キャンセル待ちNO"
        Me.vwCancelWait_Sheet1.Columns.Get(6).Visible = False
        Me.vwCancelWait_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwCancelWait_Sheet1.RowHeader.Columns.Get(0).Width = 23.0!
        Me.vwCancelWait_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'grpCancelDone
        '
        Me.grpCancelDone.BackColor = System.Drawing.Color.Transparent
        Me.grpCancelDone.Controls.Add(Me.vwCancelDone)
        Me.grpCancelDone.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.grpCancelDone.Location = New System.Drawing.Point(1916, 664)
        Me.grpCancelDone.Margin = New System.Windows.Forms.Padding(4)
        Me.grpCancelDone.Name = "grpCancelDone"
        Me.grpCancelDone.Padding = New System.Windows.Forms.Padding(4)
        Me.grpCancelDone.Size = New System.Drawing.Size(616, 361)
        Me.grpCancelDone.TabIndex = 24
        Me.grpCancelDone.TabStop = False
        Me.grpCancelDone.Text = "キャンセル済み一覧"
        '
        'vwCancelDone
        '
        Me.vwCancelDone.AccessibleDescription = "vwCancelDone, Sheet1, Row 0, Column 0, 新製品発表会"
        Me.vwCancelDone.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwCancelDone.Location = New System.Drawing.Point(24, 46)
        Me.vwCancelDone.Margin = New System.Windows.Forms.Padding(4)
        Me.vwCancelDone.Name = "vwCancelDone"
        Me.vwCancelDone.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwCancelDone_Sheet1})
        Me.vwCancelDone.Size = New System.Drawing.Size(581, 278)
        Me.vwCancelDone.TabIndex = 0
        Me.vwCancelDone.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        '
        'vwCancelDone_Sheet1
        '
        Me.vwCancelDone_Sheet1.Reset()
        Me.vwCancelDone_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwCancelDone_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwCancelDone_Sheet1.ColumnCount = 5
        Me.vwCancelDone_Sheet1.Cells.Get(0, 0).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelDone_Sheet1.Cells.Get(0, 0).Value = "新製品発表会"
        Me.vwCancelDone_Sheet1.Cells.Get(0, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelDone_Sheet1.Cells.Get(0, 1).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelDone_Sheet1.Cells.Get(0, 1).Value = "モーターワークス"
        Me.vwCancelDone_Sheet1.Cells.Get(0, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelDone_Sheet1.Cells.Get(0, 2).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelDone_Sheet1.Cells.Get(0, 2).Value = "5月週末"
        Me.vwCancelDone_Sheet1.Cells.Get(0, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        ButtonCellType3.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType3.Text = "詳細"
        Me.vwCancelDone_Sheet1.Cells.Get(0, 3).CellType = ButtonCellType3
        Me.vwCancelDone_Sheet1.Cells.Get(0, 3).Value = "詳細"
        Me.vwCancelDone_Sheet1.Cells.Get(1, 0).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelDone_Sheet1.Cells.Get(1, 0).Value = "入社式"
        Me.vwCancelDone_Sheet1.Cells.Get(1, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelDone_Sheet1.Cells.Get(1, 1).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelDone_Sheet1.Cells.Get(1, 1).Value = "株式会社六本木"
        Me.vwCancelDone_Sheet1.Cells.Get(1, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCancelDone_Sheet1.Cells.Get(1, 2).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelDone_Sheet1.Cells.Get(1, 2).Value = "5/11,5/12"
        Me.vwCancelDone_Sheet1.Cells.Get(1, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        ButtonCellType4.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType4.Text = "詳細"
        Me.vwCancelDone_Sheet1.Cells.Get(1, 3).CellType = ButtonCellType4
        Me.vwCancelDone_Sheet1.Cells.Get(1, 3).Value = "詳細"
        Me.vwCancelDone_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "催事名"
        Me.vwCancelDone_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "利用者"
        Me.vwCancelDone_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "日程"
        Me.vwCancelDone_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = " "
        Me.vwCancelDone_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "キャンセル待ちNO"
        Me.vwCancelDone_Sheet1.Columns.Get(0).Label = "催事名"
        Me.vwCancelDone_Sheet1.Columns.Get(0).Width = 185.0!
        Me.vwCancelDone_Sheet1.Columns.Get(1).Label = "利用者"
        Me.vwCancelDone_Sheet1.Columns.Get(1).Width = 105.0!
        Me.vwCancelDone_Sheet1.Columns.Get(2).Label = "日程"
        Me.vwCancelDone_Sheet1.Columns.Get(2).Width = 80.0!
        ButtonCellType5.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType5.Text = "詳細"
        Me.vwCancelDone_Sheet1.Columns.Get(3).CellType = ButtonCellType5
        Me.vwCancelDone_Sheet1.Columns.Get(3).Font = New System.Drawing.Font("MS UI Gothic", 9.0!)
        Me.vwCancelDone_Sheet1.Columns.Get(3).Label = " "
        Me.vwCancelDone_Sheet1.Columns.Get(3).Width = 42.0!
        Me.vwCancelDone_Sheet1.Columns.Get(4).Label = "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "キャンセル待ちNO"
        Me.vwCancelDone_Sheet1.Columns.Get(4).Visible = False
        Me.vwCancelDone_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwCancelDone_Sheet1.RowHeader.Columns.Get(0).Width = 19.0!
        Me.vwCancelDone_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'btnUserList
        '
        Me.btnUserList.BackColor = System.Drawing.Color.Yellow
        Me.btnUserList.Location = New System.Drawing.Point(504, 29)
        Me.btnUserList.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUserList.Name = "btnUserList"
        Me.btnUserList.Size = New System.Drawing.Size(232, 38)
        Me.btnUserList.TabIndex = 5
        Me.btnUserList.Text = "利用者要注意一覧を見る"
        Me.btnUserList.UseVisualStyleBackColor = False
        '
        'vwCalandarFirst_Sheet1
        '
        Me.vwCalandarFirst_Sheet1.Reset()
        Me.vwCalandarFirst_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwCalandarFirst_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwCalandarFirst_Sheet1.ColumnCount = 15
        Me.vwCalandarFirst_Sheet1.RowCount = 7
        Me.vwCalandarFirst_Sheet1.Cells.Get(0, 9).Value = True
        Me.vwCalandarFirst_Sheet1.Cells.Get(0, 10).Value = True
        Me.vwCalandarFirst_Sheet1.Cells.Get(1, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(1, 1).Value = "0900-1800" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社ミュージックプランニング" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "音楽祭"
        Me.vwCalandarFirst_Sheet1.Cells.Get(1, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(1, 5).Value = "1000-1500" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社アクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ライオンクイーン" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.vwCalandarFirst_Sheet1.Cells.Get(1, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(1, 8).Value = "0900-1800" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社エンターテイメント" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "六本木フェス"
        Me.vwCalandarFirst_Sheet1.Cells.Get(2, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(2, 5).Value = "1800-2100" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社フードデリシャス" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "世界のワイン試飲会"
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 1).Value = "1000-1500" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社システムシンク" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "製品発表会"
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 2).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 3).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 4).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 6).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 7).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 9).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 11).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 12).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 13).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(3, 14).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 1).Value = "0900-1300" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "サービス株式会社" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "TSライブ"
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 2).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 3).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 4).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 6).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 7).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 9).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 11).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 12).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 13).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(4, 14).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 2).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 3).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 4).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 6).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 7).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 9).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 11).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 12).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 13).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(5, 14).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(6, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalandarFirst_Sheet1.Cells.Get(6, 0).Value = "18:00～20:00" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "屋上庭園" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "テレビ朝日" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "社内イベント"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "1(金)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "2(土)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "3(日)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "4(月)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "5(火)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "6(水)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "7(木)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "8(金)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "9(土)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "10(日)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "11(月)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "12(火)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "13(水)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "14(木)"
        Me.vwCalandarFirst_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "15(金)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(0).Label = "1(金)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(0).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(1).Label = "2(土)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(1).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(2).Label = "3(日)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(2).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(3).Label = "4(月)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(3).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(4).Label = "5(火)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(4).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(5).Label = "6(水)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(5).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(6).Label = "7(木)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(6).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(7).Label = "8(金)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(7).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(8).Label = "9(土)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(8).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(9).Label = "10(日)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(9).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(10).Label = "11(月)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(10).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(11).Label = "12(火)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(11).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(12).Label = "13(水)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(12).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(13).Label = "14(木)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(13).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.Columns.Get(14).Label = "15(金)"
        Me.vwCalandarFirst_Sheet1.Columns.Get(14).Width = 85.0!
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(0, 0).Value = "選択"
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(1, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(1, 0).RowSpan = 2
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(1, 0).Value = "シ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ア" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "タ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "|"
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(1, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(2, 0).Value = " "
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(3, 0).RowSpan = 3
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(3, 0).Value = "キ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ャ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ン" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "セ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ル" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "待" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ち"
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(4, 0).Value = " "
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(5, 0).Value = " "
        Me.vwCalandarFirst_Sheet1.RowHeader.Cells.Get(6, 0).Value = "そ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "の" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "他"
        Me.vwCalandarFirst_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwCalandarFirst_Sheet1.Rows.Get(0).CellType = CheckBoxCellType1
        Me.vwCalandarFirst_Sheet1.Rows.Get(0).Height = 19.0!
        Me.vwCalandarFirst_Sheet1.Rows.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCalandarFirst_Sheet1.Rows.Get(0).Label = "選択"
        Me.vwCalandarFirst_Sheet1.Rows.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCalandarFirst_Sheet1.Rows.Get(1).Height = 105.0!
        Me.vwCalandarFirst_Sheet1.Rows.Get(1).Label = "シ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ア" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "タ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "|"
        Me.vwCalandarFirst_Sheet1.Rows.Get(2).Height = 50.0!
        Me.vwCalandarFirst_Sheet1.Rows.Get(2).Label = " "
        Me.vwCalandarFirst_Sheet1.Rows.Get(3).Height = 50.0!
        Me.vwCalandarFirst_Sheet1.Rows.Get(3).Label = "キ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ャ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ン" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "セ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ル" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "待" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ち"
        Me.vwCalandarFirst_Sheet1.Rows.Get(4).Height = 50.0!
        Me.vwCalandarFirst_Sheet1.Rows.Get(4).Label = " "
        Me.vwCalandarFirst_Sheet1.Rows.Get(5).Height = 50.0!
        Me.vwCalandarFirst_Sheet1.Rows.Get(5).Label = " "
        Me.vwCalandarFirst_Sheet1.Rows.Get(6).Height = 60.0!
        Me.vwCalandarFirst_Sheet1.Rows.Get(6).Label = "そ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "の" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "他"
        Me.vwCalandarFirst_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwCalandarFirst
        '
        Me.vwCalandarFirst.AccessibleDescription = "FpSpread1, Sheet1, Row 0, Column 0, "
        Me.vwCalandarFirst.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwCalandarFirst.Location = New System.Drawing.Point(33, 142)
        Me.vwCalandarFirst.Margin = New System.Windows.Forms.Padding(4)
        Me.vwCalandarFirst.Name = "vwCalandarFirst"
        Me.vwCalandarFirst.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwCalandarFirst_Sheet1})
        Me.vwCalandarFirst.Size = New System.Drawing.Size(1752, 510)
        Me.vwCalandarFirst.TabIndex = 14
        Me.vwCalandarFirst.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        '
        'vwCalendarSecond
        '
        Me.vwCalendarSecond.AccessibleDescription = "FpSpread2, Sheet1, Row 0, Column 0, "
        Me.vwCalendarSecond.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.vwCalendarSecond.Location = New System.Drawing.Point(33, 666)
        Me.vwCalendarSecond.Margin = New System.Windows.Forms.Padding(4)
        Me.vwCalendarSecond.Name = "vwCalendarSecond"
        Me.vwCalendarSecond.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwCalendarSecond_Sheet1})
        Me.vwCalendarSecond.Size = New System.Drawing.Size(1860, 510)
        Me.vwCalendarSecond.TabIndex = 15
        Me.vwCalendarSecond.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        '
        'vwCalendarSecond_Sheet1
        '
        Me.vwCalendarSecond_Sheet1.Reset()
        Me.vwCalendarSecond_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwCalendarSecond_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwCalendarSecond_Sheet1.ColumnCount = 16
        Me.vwCalendarSecond_Sheet1.RowCount = 7
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 0).Value = "0900-1800" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "演劇四季株式会社" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ごーちゃんと一緒"
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 7).BackColor = System.Drawing.Color.FromArgb(CType(CType(252, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 7).Value = "0900-1800" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "財団法人スポーツ振興協会" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "スポーツ振興イベント"
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(154, Byte), Integer), CType(CType(154, Byte), Integer), CType(CType(154, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 10).Value = "舞台メンテナンス"
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 12).BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 12).Value = "0900-1800" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社イベント企画" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "△△フェア"
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 13).BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 13).Value = "0900-1800" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社イベント企画" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "△△フェア"
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 14).BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(1, 14).Value = "0900-1800" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "株式会社イベント企画" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "△△フェア"
        Me.vwCalendarSecond_Sheet1.Cells.Get(2, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(2, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(154, Byte), Integer), CType(CType(154, Byte), Integer), CType(CType(154, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 2).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 3).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 4).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 6).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 7).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 9).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 11).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 12).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 13).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 14).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(3, 15).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 2).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 3).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 4).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 6).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 7).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 9).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 11).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 12).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 13).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 14).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(4, 15).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 1).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 2).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 3).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 4).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 5).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 6).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 7).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 8).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 9).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 10).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 11).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 12).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 13).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 14).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(5, 15).BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(6, 0).BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.vwCalendarSecond_Sheet1.Cells.Get(6, 0).Value = "18:00～20:00" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "テレビ朝日" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "社内イベント"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "16(土)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "17(日)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "18(月)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "19(火)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "20(水)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "21(木)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "22(金)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "23(土)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "24(日)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "25(月)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "26(火)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "27(水)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "28(木)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "29(金)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "30(土)"
        Me.vwCalendarSecond_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "31(日)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(0).Label = "16(土)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(0).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(1).Label = "17(日)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(1).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(2).Label = "18(月)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(2).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(3).Label = "19(火)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(3).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(4).Label = "20(水)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(4).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(5).Label = "21(木)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(5).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(6).Label = "22(金)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(6).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(7).Label = "23(土)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(7).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(8).Label = "24(日)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(8).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(9).Label = "25(月)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(9).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(10).Label = "26(火)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(10).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(11).Label = "27(水)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(11).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(12).Label = "28(木)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(12).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(13).Label = "29(金)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(13).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(14).Label = "30(土)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(14).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.Columns.Get(15).Label = "31(日)"
        Me.vwCalendarSecond_Sheet1.Columns.Get(15).Width = 85.0!
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(0, 0).Value = "選択"
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(1, 0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(1, 0).RowSpan = 2
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(1, 0).Value = "シ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ア" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "タ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "|"
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(1, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(3, 0).RowSpan = 3
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(3, 0).Value = "キ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ャ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ン" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "セ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ル" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "待" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ち"
        Me.vwCalendarSecond_Sheet1.RowHeader.Cells.Get(6, 0).Value = "そ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "の" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "他"
        Me.vwCalendarSecond_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwCalendarSecond_Sheet1.Rows.Get(0).CellType = CheckBoxCellType2
        Me.vwCalendarSecond_Sheet1.Rows.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwCalendarSecond_Sheet1.Rows.Get(0).Label = "選択"
        Me.vwCalendarSecond_Sheet1.Rows.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwCalendarSecond_Sheet1.Rows.Get(1).Height = 105.0!
        Me.vwCalendarSecond_Sheet1.Rows.Get(1).Label = "シ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ア" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "タ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "|"
        Me.vwCalendarSecond_Sheet1.Rows.Get(2).Height = 50.0!
        Me.vwCalendarSecond_Sheet1.Rows.Get(3).Height = 50.0!
        Me.vwCalendarSecond_Sheet1.Rows.Get(3).Label = "キ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ャ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ン" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "セ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ル" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "待" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ち"
        Me.vwCalendarSecond_Sheet1.Rows.Get(4).Height = 50.0!
        Me.vwCalendarSecond_Sheet1.Rows.Get(5).Height = 50.0!
        Me.vwCalendarSecond_Sheet1.Rows.Get(6).Height = 60.0!
        Me.vwCalendarSecond_Sheet1.Rows.Get(6).Label = "そ" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "の" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "他"
        Me.vwCalendarSecond_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'btnChangeKyukanbi
        '
        Me.btnChangeKyukanbi.Location = New System.Drawing.Point(1196, 1195)
        Me.btnChangeKyukanbi.Margin = New System.Windows.Forms.Padding(4)
        Me.btnChangeKyukanbi.Name = "btnChangeKyukanbi"
        Me.btnChangeKyukanbi.Size = New System.Drawing.Size(228, 36)
        Me.btnChangeKyukanbi.TabIndex = 20
        Me.btnChangeKyukanbi.Text = "選択日を休館日とする"
        Me.btnChangeKyukanbi.UseVisualStyleBackColor = True
        '
        'btnChangeHoliday
        '
        Me.btnChangeHoliday.Location = New System.Drawing.Point(647, 1195)
        Me.btnChangeHoliday.Margin = New System.Windows.Forms.Padding(4)
        Me.btnChangeHoliday.Name = "btnChangeHoliday"
        Me.btnChangeHoliday.Size = New System.Drawing.Size(228, 36)
        Me.btnChangeHoliday.TabIndex = 18
        Me.btnChangeHoliday.Text = "選択日を祝日（休日）とする"
        Me.btnChangeHoliday.UseVisualStyleBackColor = True
        '
        'btnChangeWeekday
        '
        Me.btnChangeWeekday.Location = New System.Drawing.Point(881, 1195)
        Me.btnChangeWeekday.Margin = New System.Windows.Forms.Padding(4)
        Me.btnChangeWeekday.Name = "btnChangeWeekday"
        Me.btnChangeWeekday.Size = New System.Drawing.Size(228, 36)
        Me.btnChangeWeekday.TabIndex = 19
        Me.btnChangeWeekday.Text = "選択日を平日とする"
        Me.btnChangeWeekday.UseVisualStyleBackColor = True
        '
        'EXTB0101
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTB.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(2548, 1284)
        Me.Controls.Add(Me.btnChangeWeekday)
        Me.Controls.Add(Me.btnChangeHoliday)
        Me.Controls.Add(Me.btnChangeKyukanbi)
        Me.Controls.Add(Me.vwCalendarSecond)
        Me.Controls.Add(Me.vwCalandarFirst)
        Me.Controls.Add(Me.btnUserList)
        Me.Controls.Add(Me.grpCancelDone)
        Me.Controls.Add(Me.grpCancelWait)
        Me.Controls.Add(Me.btnSearchNextMonth)
        Me.Controls.Add(Me.btnSearchPrevMonth)
        Me.Controls.Add(Me.btnToMenu)
        Me.Controls.Add(Me.btnChangeEigyobi)
        Me.Controls.Add(Me.btnChangeMaintebi)
        Me.Controls.Add(Me.btnRegistCancelWait)
        Me.Controls.Add(Me.btnRegistKariYoyaku)
        Me.Controls.Add(Me.lblMaintenance)
        Me.Controls.Add(Me.lblKyukanbi)
        Me.Controls.Add(Me.lblSeisikiYoyakuKanryo)
        Me.Controls.Add(Me.lblSeisikiYoyaku)
        Me.Controls.Add(Me.lblKariyoyaku)
        Me.Controls.Add(Me.lblKariYoyakuMikakunin)
        Me.Controls.Add(Me.lblMonth)
        Me.Controls.Add(Me.txtMonth)
        Me.Controls.Add(Me.lblYear)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.txtYear)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "EXTB0101"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　予約カレンダー（シアター）"
        Me.grpCancelWait.ResumeLayout(False)
        CType(Me.vwCancelWait, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwCancelWait_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCancelDone.ResumeLayout(False)
        CType(Me.vwCancelDone, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwCancelDone_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwCalandarFirst_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwCalandarFirst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwCalendarSecond, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwCalendarSecond_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents txtMonth As System.Windows.Forms.TextBox
    Friend WithEvents lblMonth As System.Windows.Forms.Label
    Friend WithEvents lblKariYoyakuMikakunin As System.Windows.Forms.Label
    Friend WithEvents lblKariyoyaku As System.Windows.Forms.Label
    Friend WithEvents lblSeisikiYoyaku As System.Windows.Forms.Label
    Friend WithEvents lblSeisikiYoyakuKanryo As System.Windows.Forms.Label
    Friend WithEvents lblKyukanbi As System.Windows.Forms.Label
    Friend WithEvents lblMaintenance As System.Windows.Forms.Label
    Friend WithEvents btnRegistKariYoyaku As System.Windows.Forms.Button
    Friend WithEvents btnRegistCancelWait As System.Windows.Forms.Button
    Friend WithEvents btnChangeMaintebi As System.Windows.Forms.Button
    Friend WithEvents btnChangeEigyobi As System.Windows.Forms.Button
    Friend WithEvents btnToMenu As System.Windows.Forms.Button
    Friend WithEvents btnSearchPrevMonth As System.Windows.Forms.Button
    Friend WithEvents btnSearchNextMonth As System.Windows.Forms.Button
    Friend WithEvents grpCancelWait As System.Windows.Forms.GroupBox
    Friend WithEvents grpCancelDone As System.Windows.Forms.GroupBox
    Friend WithEvents btnUserList As System.Windows.Forms.Button
    Friend WithEvents vwCalandarFirst_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwCalandarFirst As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwCalendarSecond As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwCalendarSecond_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwCancelWait As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwCancelWait_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwCancelDone As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwCancelDone_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents btnChangeKyukanbi As System.Windows.Forms.Button
    Friend WithEvents btnChangeHoliday As System.Windows.Forms.Button
    Friend WithEvents btnChangeWeekday As System.Windows.Forms.Button
End Class
