<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0214
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0214))
        Me.btnFileSyutu = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.lblTanto = New System.Windows.Forms.Label()
        Me.lblNaiyo = New System.Windows.Forms.Label()
        Me.cmbSeikyuNaiyo = New System.Windows.Forms.ComboBox()
        Me.lblExasAiteNm = New System.Windows.Forms.Label()
        Me.lblExasAite = New System.Windows.Forms.Label()
        Me.lblAite = New System.Windows.Forms.Label()
        Me.cmbSeikyusakiBusyo = New System.Windows.Forms.ComboBox()
        Me.lblBusyo = New System.Windows.Forms.Label()
        Me.txtTantoCD = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnFileSyutu
        '
        Me.btnFileSyutu.Location = New System.Drawing.Point(273, 281)
        Me.btnFileSyutu.Name = "btnFileSyutu"
        Me.btnFileSyutu.Size = New System.Drawing.Size(96, 23)
        Me.btnFileSyutu.TabIndex = 4
        Me.btnFileSyutu.Text = "ファイル出力"
        Me.btnFileSyutu.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(158, 281)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 5
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'lblTanto
        '
        Me.lblTanto.BackColor = System.Drawing.Color.Transparent
        Me.lblTanto.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTanto.Location = New System.Drawing.Point(31, 111)
        Me.lblTanto.Name = "lblTanto"
        Me.lblTanto.Size = New System.Drawing.Size(154, 18)
        Me.lblTanto.TabIndex = 92
        Me.lblTanto.Text = "G請求書担当者コード　＊"
        '
        'lblNaiyo
        '
        Me.lblNaiyo.BackColor = System.Drawing.Color.Transparent
        Me.lblNaiyo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblNaiyo.Location = New System.Drawing.Point(31, 158)
        Me.lblNaiyo.Name = "lblNaiyo"
        Me.lblNaiyo.Size = New System.Drawing.Size(154, 18)
        Me.lblNaiyo.TabIndex = 94
        Me.lblNaiyo.Text = "G請求内容                ＊"
        '
        'cmbSeikyuNaiyo
        '
        Me.cmbSeikyuNaiyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSeikyuNaiyo.FormattingEnabled = True
        Me.cmbSeikyuNaiyo.Location = New System.Drawing.Point(218, 158)
        Me.cmbSeikyuNaiyo.Name = "cmbSeikyuNaiyo"
        Me.cmbSeikyuNaiyo.Size = New System.Drawing.Size(188, 20)
        Me.cmbSeikyuNaiyo.TabIndex = 2
        '
        'lblExasAiteNm
        '
        Me.lblExasAiteNm.BackColor = System.Drawing.Color.Silver
        Me.lblExasAiteNm.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblExasAiteNm.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblExasAiteNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblExasAiteNm.Location = New System.Drawing.Point(222, 8)
        Me.lblExasAiteNm.Name = "lblExasAiteNm"
        Me.lblExasAiteNm.Size = New System.Drawing.Size(283, 20)
        Me.lblExasAiteNm.TabIndex = 142
        Me.lblExasAiteNm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblExasAite
        '
        Me.lblExasAite.BackColor = System.Drawing.Color.Silver
        Me.lblExasAite.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblExasAite.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblExasAite.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblExasAite.Location = New System.Drawing.Point(128, 8)
        Me.lblExasAite.Name = "lblExasAite"
        Me.lblExasAite.Size = New System.Drawing.Size(88, 20)
        Me.lblExasAite.TabIndex = 141
        Me.lblExasAite.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAite
        '
        Me.lblAite.AutoSize = True
        Me.lblAite.BackColor = System.Drawing.Color.Transparent
        Me.lblAite.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblAite.Location = New System.Drawing.Point(6, 10)
        Me.lblAite.Name = "lblAite"
        Me.lblAite.Size = New System.Drawing.Size(77, 13)
        Me.lblAite.TabIndex = 140
        Me.lblAite.Text = "EXAS相手先"
        '
        'cmbSeikyusakiBusyo
        '
        Me.cmbSeikyusakiBusyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSeikyusakiBusyo.FormattingEnabled = True
        Me.cmbSeikyusakiBusyo.Location = New System.Drawing.Point(218, 64)
        Me.cmbSeikyusakiBusyo.Name = "cmbSeikyusakiBusyo"
        Me.cmbSeikyusakiBusyo.Size = New System.Drawing.Size(188, 20)
        Me.cmbSeikyusakiBusyo.TabIndex = 0
        '
        'lblBusyo
        '
        Me.lblBusyo.BackColor = System.Drawing.Color.Transparent
        Me.lblBusyo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblBusyo.Location = New System.Drawing.Point(31, 64)
        Me.lblBusyo.Name = "lblBusyo"
        Me.lblBusyo.Size = New System.Drawing.Size(154, 18)
        Me.lblBusyo.TabIndex = 143
        Me.lblBusyo.Text = "G請求先部署             ＊"
        '
        'txtTantoCD
        '
        Me.txtTantoCD.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtTantoCD.ImeMode = System.Windows.Forms.ImeMode.Alpha
        Me.txtTantoCD.Location = New System.Drawing.Point(218, 111)
        Me.txtTantoCD.MaxLength = 6
        Me.txtTantoCD.Name = "txtTantoCD"
        Me.txtTantoCD.Size = New System.Drawing.Size(106, 20)
        Me.txtTantoCD.TabIndex = 1
        '
        'EXTZ0214
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(515, 341)
        Me.Controls.Add(Me.txtTantoCD)
        Me.Controls.Add(Me.cmbSeikyusakiBusyo)
        Me.Controls.Add(Me.lblBusyo)
        Me.Controls.Add(Me.lblExasAiteNm)
        Me.Controls.Add(Me.lblExasAite)
        Me.Controls.Add(Me.lblAite)
        Me.Controls.Add(Me.cmbSeikyuNaiyo)
        Me.Controls.Add(Me.lblNaiyo)
        Me.Controls.Add(Me.lblTanto)
        Me.Controls.Add(Me.btnFileSyutu)
        Me.Controls.Add(Me.btnBack)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0214"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　グループ請求情報入力"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnFileSyutu As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents lblTanto As System.Windows.Forms.Label
    Friend WithEvents lblNaiyo As System.Windows.Forms.Label
    Friend WithEvents cmbSeikyuNaiyo As System.Windows.Forms.ComboBox
    Friend WithEvents lblExasAiteNm As System.Windows.Forms.Label
    Friend WithEvents lblExasAite As System.Windows.Forms.Label
    Friend WithEvents lblAite As System.Windows.Forms.Label
    Friend WithEvents cmbSeikyusakiBusyo As System.Windows.Forms.ComboBox
    Friend WithEvents lblBusyo As System.Windows.Forms.Label
    Friend WithEvents txtTantoCD As System.Windows.Forms.TextBox
End Class
