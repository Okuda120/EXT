<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0211
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0211))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnRegister = New System.Windows.Forms.Button()
        Me.rdoCancel = New System.Windows.Forms.RadioButton()
        Me.rdoDelete = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.BackColor = System.Drawing.Color.Transparent
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(14, 18)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(318, 31)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "予約取り消しの種類(理由)を選択してください。"
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(180, 178)
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
        'btnRegister
        '
        Me.btnRegister.Location = New System.Drawing.Point(296, 178)
        Me.btnRegister.Name = "btnRegister"
        Me.btnRegister.Size = New System.Drawing.Size(96, 23)
        Me.btnRegister.TabIndex = 89
        Me.btnRegister.Text = "確定"
        Me.btnRegister.UseVisualStyleBackColor = True
        '
        'rdoCancel
        '
        Me.rdoCancel.AutoSize = True
        Me.rdoCancel.BackColor = System.Drawing.Color.Transparent
        Me.rdoCancel.Checked = True
        Me.rdoCancel.Location = New System.Drawing.Point(17, 50)
        Me.rdoCancel.Name = "rdoCancel"
        Me.rdoCancel.Size = New System.Drawing.Size(77, 17)
        Me.rdoCancel.TabIndex = 90
        Me.rdoCancel.TabStop = True
        Me.rdoCancel.Text = "キャンセル"
        Me.rdoCancel.UseVisualStyleBackColor = False
        '
        'rdoDelete
        '
        Me.rdoDelete.AutoSize = True
        Me.rdoDelete.BackColor = System.Drawing.Color.Transparent
        Me.rdoDelete.Location = New System.Drawing.Point(17, 108)
        Me.rdoDelete.Name = "rdoDelete"
        Me.rdoDelete.Size = New System.Drawing.Size(51, 17)
        Me.rdoDelete.TabIndex = 91
        Me.rdoDelete.Text = "取消"
        Me.rdoDelete.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(14, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(360, 35)
        Me.Label3.TabIndex = 92
        Me.Label3.Text = "利用者からのキャンセル依頼の場合はこちらを選択してください。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "カレンダーからは消えますが、引き続き管理を行うことができます。"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 128)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(360, 35)
        Me.Label1.TabIndex = 93
        Me.Label1.Text = "誤って登録してしまった場合はこちらを選択してください。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "登録データは完全に削除されます。"
        '
        'EXTZ0211
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(404, 213)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.rdoDelete)
        Me.Controls.Add(Me.rdoCancel)
        Me.Controls.Add(Me.btnRegister)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label65)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0211"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　予約取消"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnRegister As System.Windows.Forms.Button
    Friend WithEvents rdoCancel As System.Windows.Forms.RadioButton
    Friend WithEvents rdoDelete As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
