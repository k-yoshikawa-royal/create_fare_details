<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Label0 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Cf_TextBox1 = New System.Windows.Forms.TextBox()
        Me.Cf_Label1 = New System.Windows.Forms.Label()
        Me.Cf_TextBox101 = New System.Windows.Forms.TextBox()
        Me.Cf_Label100 = New System.Windows.Forms.Label()
        Me.Cf_TextBox100 = New System.Windows.Forms.TextBox()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(13, 6)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(308, 348)
        Me.TabControl1.TabIndex = 18
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage1.Controls.Add(Me.Button4)
        Me.TabPage1.Controls.Add(Me.Button9)
        Me.TabPage1.Controls.Add(Me.Label0)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.Button3)
        Me.TabPage1.Controls.Add(Me.TextBox1)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(300, 322)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "メイン"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(16, 150)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(269, 23)
        Me.Button4.TabIndex = 72
        Me.Button4.Text = "データ集計／計算処理"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(219, 293)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(75, 23)
        Me.Button9.TabIndex = 71
        Me.Button9.Text = "閉じる"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Label0
        '
        Me.Label0.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label0.Location = New System.Drawing.Point(2, 19)
        Me.Label0.Name = "Label0"
        Me.Label0.Size = New System.Drawing.Size(263, 19)
        Me.Label0.TabIndex = 70
        Me.Label0.Text = "グローバル運賃明細　取込"
        Me.Label0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(15, 146)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(269, 19)
        Me.Label2.TabIndex = 69
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(15, 120)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(269, 23)
        Me.Button1.TabIndex = 68
        Me.Button1.Text = "Excelファイルからデータ取り込み"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button3.Location = New System.Drawing.Point(213, 91)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(71, 23)
        Me.Button3.TabIndex = 67
        Me.Button3.Text = "ファイル選択"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(15, 91)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(192, 19)
        Me.TextBox1.TabIndex = 66
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(17, 69)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(263, 19)
        Me.Label1.TabIndex = 65
        Me.Label1.Text = "運賃明細　取り込み"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.LightYellow
        Me.TabPage2.Controls.Add(Me.Cf_TextBox101)
        Me.TabPage2.Controls.Add(Me.Cf_Label100)
        Me.TabPage2.Controls.Add(Me.Cf_TextBox100)
        Me.TabPage2.Controls.Add(Me.Button2)
        Me.TabPage2.Controls.Add(Me.Cf_TextBox1)
        Me.TabPage2.Controls.Add(Me.Cf_Label1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(300, 322)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "設定"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(220, 293)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(74, 23)
        Me.Button2.TabIndex = 11
        Me.Button2.Text = "設定保存"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Cf_TextBox1
        '
        Me.Cf_TextBox1.Location = New System.Drawing.Point(113, 16)
        Me.Cf_TextBox1.Name = "Cf_TextBox1"
        Me.Cf_TextBox1.Size = New System.Drawing.Size(170, 19)
        Me.Cf_TextBox1.TabIndex = 3
        '
        'Cf_Label1
        '
        Me.Cf_Label1.Location = New System.Drawing.Point(20, 19)
        Me.Cf_Label1.Name = "Cf_Label1"
        Me.Cf_Label1.Size = New System.Drawing.Size(87, 19)
        Me.Cf_Label1.TabIndex = 4
        Me.Cf_Label1.Text = "保存フォルダ"
        '
        'Cf_TextBox101
        '
        Me.Cf_TextBox101.Location = New System.Drawing.Point(6, 296)
        Me.Cf_TextBox101.Name = "Cf_TextBox101"
        Me.Cf_TextBox101.Size = New System.Drawing.Size(136, 19)
        Me.Cf_TextBox101.TabIndex = 14
        '
        'Cf_Label100
        '
        Me.Cf_Label100.Location = New System.Drawing.Point(6, 250)
        Me.Cf_Label100.Name = "Cf_Label100"
        Me.Cf_Label100.Size = New System.Drawing.Size(136, 19)
        Me.Cf_Label100.TabIndex = 13
        Me.Cf_Label100.Text = "テンポラリデータ"
        Me.Cf_Label100.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Cf_TextBox100
        '
        Me.Cf_TextBox100.Location = New System.Drawing.Point(6, 271)
        Me.Cf_TextBox100.Name = "Cf_TextBox100"
        Me.Cf_TextBox100.Size = New System.Drawing.Size(136, 19)
        Me.Cf_TextBox100.TabIndex = 12
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(333, 361)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "Form1"
        Me.Text = "グローバル運賃明細取り込み"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Button2 As Button
    Friend WithEvents Cf_TextBox1 As TextBox
    Friend WithEvents Cf_Label1 As Label
    Friend WithEvents Label0 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button9 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Cf_TextBox101 As TextBox
    Friend WithEvents Cf_Label100 As Label
    Friend WithEvents Cf_TextBox100 As TextBox
End Class
