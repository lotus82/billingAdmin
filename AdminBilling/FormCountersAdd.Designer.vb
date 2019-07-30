<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCountersAdd
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormCountersAdd))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox_Model_Name = New System.Windows.Forms.TextBox()
        Me.NumericUpDown_Digit_Capacity = New System.Windows.Forms.NumericUpDown()
        Me.NumericUpDown_Coeff = New System.Windows.Forms.NumericUpDown()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ComboBox_RES_ID = New System.Windows.Forms.ComboBox()
        Me.ComboBox_SERVICE_ID = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.NumericUpDown_Period_Ver = New System.Windows.Forms.NumericUpDown()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.NumericUpDown_Precision = New System.Windows.Forms.NumericUpDown()
        CType(Me.NumericUpDown_Digit_Capacity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Coeff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Period_Ver, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Precision, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Название"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(249, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(26, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Тип"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 119)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Разрядность"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(229, 119)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Коэффициент"
        '
        'TextBox_Model_Name
        '
        Me.TextBox_Model_Name.Location = New System.Drawing.Point(13, 77)
        Me.TextBox_Model_Name.Name = "TextBox_Model_Name"
        Me.TextBox_Model_Name.Size = New System.Drawing.Size(234, 20)
        Me.TextBox_Model_Name.TabIndex = 4
        '
        'NumericUpDown_Digit_Capacity
        '
        Me.NumericUpDown_Digit_Capacity.Location = New System.Drawing.Point(17, 146)
        Me.NumericUpDown_Digit_Capacity.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.NumericUpDown_Digit_Capacity.Name = "NumericUpDown_Digit_Capacity"
        Me.NumericUpDown_Digit_Capacity.Size = New System.Drawing.Size(79, 20)
        Me.NumericUpDown_Digit_Capacity.TabIndex = 6
        '
        'NumericUpDown_Coeff
        '
        Me.NumericUpDown_Coeff.Increment = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NumericUpDown_Coeff.Location = New System.Drawing.Point(231, 146)
        Me.NumericUpDown_Coeff.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.NumericUpDown_Coeff.Name = "NumericUpDown_Coeff"
        Me.NumericUpDown_Coeff.Size = New System.Drawing.Size(75, 20)
        Me.NumericUpDown_Coeff.TabIndex = 7
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(313, 176)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(125, 23)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Добавить прибор"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ComboBox_RES_ID
        '
        Me.ComboBox_RES_ID.FormattingEnabled = True
        Me.ComboBox_RES_ID.Location = New System.Drawing.Point(252, 77)
        Me.ComboBox_RES_ID.Name = "ComboBox_RES_ID"
        Me.ComboBox_RES_ID.Size = New System.Drawing.Size(112, 21)
        Me.ComboBox_RES_ID.TabIndex = 9
        '
        'ComboBox_SERVICE_ID
        '
        Me.ComboBox_SERVICE_ID.FormattingEnabled = True
        Me.ComboBox_SERVICE_ID.Location = New System.Drawing.Point(372, 77)
        Me.ComboBox_SERVICE_ID.Name = "ComboBox_SERVICE_ID"
        Me.ComboBox_SERVICE_ID.Size = New System.Drawing.Size(106, 21)
        Me.ComboBox_SERVICE_ID.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(369, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Услуга"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(485, 42)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(119, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Период поверки (мес)"
        '
        'NumericUpDown_Period_Ver
        '
        Me.NumericUpDown_Period_Ver.Location = New System.Drawing.Point(488, 77)
        Me.NumericUpDown_Period_Ver.Maximum = New Decimal(New Integer() {120, 0, 0, 0})
        Me.NumericUpDown_Period_Ver.Name = "NumericUpDown_Period_Ver"
        Me.NumericUpDown_Period_Ver.Size = New System.Drawing.Size(116, 20)
        Me.NumericUpDown_Period_Ver.TabIndex = 14
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(120, 119)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(54, 13)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Точность"
        '
        'NumericUpDown_Precision
        '
        Me.NumericUpDown_Precision.Location = New System.Drawing.Point(123, 146)
        Me.NumericUpDown_Precision.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NumericUpDown_Precision.Name = "NumericUpDown_Precision"
        Me.NumericUpDown_Precision.Size = New System.Drawing.Size(71, 20)
        Me.NumericUpDown_Precision.TabIndex = 16
        '
        'FormCountersAdd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(740, 225)
        Me.Controls.Add(Me.NumericUpDown_Precision)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.NumericUpDown_Period_Ver)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ComboBox_SERVICE_ID)
        Me.Controls.Add(Me.ComboBox_RES_ID)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.NumericUpDown_Coeff)
        Me.Controls.Add(Me.NumericUpDown_Digit_Capacity)
        Me.Controls.Add(Me.TextBox_Model_Name)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormCountersAdd"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Добавление прибора учета"
        CType(Me.NumericUpDown_Digit_Capacity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Coeff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Period_Ver, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Precision, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox_Model_Name As TextBox
    Friend WithEvents NumericUpDown_Digit_Capacity As NumericUpDown
    Friend WithEvents NumericUpDown_Coeff As NumericUpDown
    Friend WithEvents Button1 As Button
    Friend WithEvents ComboBox_RES_ID As ComboBox
    Friend WithEvents ComboBox_SERVICE_ID As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents NumericUpDown_Period_Ver As NumericUpDown
    Friend WithEvents Label7 As Label
    Friend WithEvents NumericUpDown_Precision As NumericUpDown
End Class
