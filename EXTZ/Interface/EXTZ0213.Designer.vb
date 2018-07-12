<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0213
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0213))
        Me.btnComplate = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.lblTitle1 = New System.Windows.Forms.Label()
        Me.lblTitle2 = New System.Windows.Forms.Label()
        Me.lblTitle3 = New System.Windows.Forms.Label()
        Me.lblTitle4 = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.txtCom = New System.Windows.Forms.TextBox()
        Me.dtpIraiCheck = New Common.DateTimePickerEx()
        Me.SuspendLayout()
        '
        'btnComplate
        '
        Me.btnComplate.Location = New System.Drawing.Point(273, 281)
        Me.btnComplate.Name = "btnComplate"
        Me.btnComplate.Size = New System.Drawing.Size(96, 23)
        Me.btnComplate.TabIndex = 91
        Me.btnComplate.Text = "入力完了"
        Me.btnComplate.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(158, 281)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 90
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'lblTitle1
        '
        Me.lblTitle1.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTitle1.Location = New System.Drawing.Point(31, 29)
        Me.lblTitle1.Name = "lblTitle1"
        Me.lblTitle1.Size = New System.Drawing.Size(83, 18)
        Me.lblTitle1.TabIndex = 92
        Me.lblTitle1.Text = "承認者"
        '
        'lblTitle2
        '
        Me.lblTitle2.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTitle2.Location = New System.Drawing.Point(31, 67)
        Me.lblTitle2.Name = "lblTitle2"
        Me.lblTitle2.Size = New System.Drawing.Size(83, 18)
        Me.lblTitle2.TabIndex = 93
        Me.lblTitle2.Text = "確認日"
        '
        'lblTitle3
        '
        Me.lblTitle3.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTitle3.Location = New System.Drawing.Point(31, 108)
        Me.lblTitle3.Name = "lblTitle3"
        Me.lblTitle3.Size = New System.Drawing.Size(115, 18)
        Me.lblTitle3.TabIndex = 94
        Me.lblTitle3.Text = "承認／差し戻し　＊"
        '
        'lblTitle4
        '
        Me.lblTitle4.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTitle4.Location = New System.Drawing.Point(31, 147)
        Me.lblTitle4.Name = "lblTitle4"
        Me.lblTitle4.Size = New System.Drawing.Size(115, 18)
        Me.lblTitle4.TabIndex = 95
        Me.lblTitle4.Text = "コメント"
        '
        'lblName
        '
        Me.lblName.BackColor = System.Drawing.Color.Silver
        Me.lblName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblName.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblName.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblName.Location = New System.Drawing.Point(151, 25)
        Me.lblName.Margin = New System.Windows.Forms.Padding(3)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(157, 21)
        Me.lblName.TabIndex = 138
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbStatus
        '
        Me.cmbStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Location = New System.Drawing.Point(151, 108)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(109, 20)
        Me.cmbStatus.TabIndex = 139
        '
        'txtCom
        '
        Me.txtCom.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtCom.Location = New System.Drawing.Point(151, 147)
        Me.txtCom.MaxLength = 100
        Me.txtCom.Multiline = True
        Me.txtCom.Name = "txtCom"
        Me.txtCom.Size = New System.Drawing.Size(334, 103)
        Me.txtCom.TabIndex = 142
        '
        'dtpIraiCheck
        '
        Me.dtpIraiCheck.Location = New System.Drawing.Point(151, 62)
        Me.dtpIraiCheck.Name = "dtpIraiCheck"
        Me.dtpIraiCheck.Size = New System.Drawing.Size(129, 23)
        Me.dtpIraiCheck.TabIndex = 143
        '
        'EXTZ0213
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(515, 341)
        Me.Controls.Add(Me.dtpIraiCheck)
        Me.Controls.Add(Me.txtCom)
        Me.Controls.Add(Me.cmbStatus)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblTitle4)
        Me.Controls.Add(Me.lblTitle3)
        Me.Controls.Add(Me.lblTitle2)
        Me.Controls.Add(Me.lblTitle1)
        Me.Controls.Add(Me.btnComplate)
        Me.Controls.Add(Me.btnBack)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0213"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　承認記録追加"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnComplate As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents lblTitle1 As System.Windows.Forms.Label
    Friend WithEvents lblTitle2 As System.Windows.Forms.Label
    Friend WithEvents lblTitle3 As System.Windows.Forms.Label
    Friend WithEvents lblTitle4 As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents cmbStatus As System.Windows.Forms.ComboBox
    Friend WithEvents txtCom As System.Windows.Forms.TextBox
    Friend WithEvents dtpIraiCheck As Common.DateTimePickerEx
End Class
