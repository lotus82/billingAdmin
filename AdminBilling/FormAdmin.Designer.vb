<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormAdmin
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormAdmin))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ЖилфондToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УлицыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДомаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТехучасткиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УслугиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоставщикиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИсполнителиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УслугиToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТарифыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЛюдиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТипыДокументовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТипыОтношенийToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТипРегистрацииToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОператорыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДанныеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.РолиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриборыУчетыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриборыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДокументыToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.КвитанцииToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ШаблоныToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПараметрыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.БазаДанныхToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.АрхивацияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВосстановлениеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ЖилфондToolStripMenuItem, Me.УслугиToolStripMenuItem, Me.ЛюдиToolStripMenuItem, Me.ОператорыToolStripMenuItem, Me.ПриборыУчетыToolStripMenuItem, Me.КвитанцииToolStripMenuItem, Me.БазаДанныхToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1004, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ЖилфондToolStripMenuItem
        '
        Me.ЖилфондToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.УлицыToolStripMenuItem, Me.ДомаToolStripMenuItem, Me.ТехучасткиToolStripMenuItem})
        Me.ЖилфондToolStripMenuItem.Name = "ЖилфондToolStripMenuItem"
        Me.ЖилфондToolStripMenuItem.Size = New System.Drawing.Size(73, 20)
        Me.ЖилфондToolStripMenuItem.Text = "Жилфонд"
        '
        'УлицыToolStripMenuItem
        '
        Me.УлицыToolStripMenuItem.Name = "УлицыToolStripMenuItem"
        Me.УлицыToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.УлицыToolStripMenuItem.Text = "Улицы"
        '
        'ДомаToolStripMenuItem
        '
        Me.ДомаToolStripMenuItem.Name = "ДомаToolStripMenuItem"
        Me.ДомаToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.ДомаToolStripMenuItem.Text = "Дома"
        '
        'ТехучасткиToolStripMenuItem
        '
        Me.ТехучасткиToolStripMenuItem.Name = "ТехучасткиToolStripMenuItem"
        Me.ТехучасткиToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.ТехучасткиToolStripMenuItem.Text = "Техучастки"
        '
        'УслугиToolStripMenuItem
        '
        Me.УслугиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПоставщикиToolStripMenuItem, Me.ИсполнителиToolStripMenuItem, Me.УслугиToolStripMenuItem1, Me.ТарифыToolStripMenuItem})
        Me.УслугиToolStripMenuItem.Name = "УслугиToolStripMenuItem"
        Me.УслугиToolStripMenuItem.Size = New System.Drawing.Size(57, 20)
        Me.УслугиToolStripMenuItem.Text = "Услуги"
        '
        'ПоставщикиToolStripMenuItem
        '
        Me.ПоставщикиToolStripMenuItem.Name = "ПоставщикиToolStripMenuItem"
        Me.ПоставщикиToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.ПоставщикиToolStripMenuItem.Text = "Получатели"
        '
        'ИсполнителиToolStripMenuItem
        '
        Me.ИсполнителиToolStripMenuItem.Name = "ИсполнителиToolStripMenuItem"
        Me.ИсполнителиToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.ИсполнителиToolStripMenuItem.Text = "Исполнители"
        '
        'УслугиToolStripMenuItem1
        '
        Me.УслугиToolStripMenuItem1.Name = "УслугиToolStripMenuItem1"
        Me.УслугиToolStripMenuItem1.Size = New System.Drawing.Size(149, 22)
        Me.УслугиToolStripMenuItem1.Text = "Услуги"
        '
        'ТарифыToolStripMenuItem
        '
        Me.ТарифыToolStripMenuItem.Name = "ТарифыToolStripMenuItem"
        Me.ТарифыToolStripMenuItem.Size = New System.Drawing.Size(149, 22)
        Me.ТарифыToolStripMenuItem.Text = "Тарифы"
        '
        'ЛюдиToolStripMenuItem
        '
        Me.ЛюдиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ТипыДокументовToolStripMenuItem, Me.ТипыОтношенийToolStripMenuItem, Me.ТипРегистрацииToolStripMenuItem})
        Me.ЛюдиToolStripMenuItem.Name = "ЛюдиToolStripMenuItem"
        Me.ЛюдиToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.ЛюдиToolStripMenuItem.Text = "Люди"
        '
        'ТипыДокументовToolStripMenuItem
        '
        Me.ТипыДокументовToolStripMenuItem.Name = "ТипыДокументовToolStripMenuItem"
        Me.ТипыДокументовToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.ТипыДокументовToolStripMenuItem.Text = "Типы документов"
        '
        'ТипыОтношенийToolStripMenuItem
        '
        Me.ТипыОтношенийToolStripMenuItem.Name = "ТипыОтношенийToolStripMenuItem"
        Me.ТипыОтношенийToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.ТипыОтношенийToolStripMenuItem.Text = "Типы отношений"
        '
        'ТипРегистрацииToolStripMenuItem
        '
        Me.ТипРегистрацииToolStripMenuItem.Name = "ТипРегистрацииToolStripMenuItem"
        Me.ТипРегистрацииToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.ТипРегистрацииToolStripMenuItem.Text = "Типы регистрации"
        '
        'ОператорыToolStripMenuItem
        '
        Me.ОператорыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ДанныеToolStripMenuItem, Me.РолиToolStripMenuItem})
        Me.ОператорыToolStripMenuItem.Name = "ОператорыToolStripMenuItem"
        Me.ОператорыToolStripMenuItem.Size = New System.Drawing.Size(82, 20)
        Me.ОператорыToolStripMenuItem.Text = "Операторы"
        '
        'ДанныеToolStripMenuItem
        '
        Me.ДанныеToolStripMenuItem.Name = "ДанныеToolStripMenuItem"
        Me.ДанныеToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ДанныеToolStripMenuItem.Text = "Данные"
        '
        'РолиToolStripMenuItem
        '
        Me.РолиToolStripMenuItem.Name = "РолиToolStripMenuItem"
        Me.РолиToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.РолиToolStripMenuItem.Text = "Роли"
        '
        'ПриборыУчетыToolStripMenuItem
        '
        Me.ПриборыУчетыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПриборыToolStripMenuItem, Me.ДокументыToolStripMenuItem1})
        Me.ПриборыУчетыToolStripMenuItem.Name = "ПриборыУчетыToolStripMenuItem"
        Me.ПриборыУчетыToolStripMenuItem.Size = New System.Drawing.Size(105, 20)
        Me.ПриборыУчетыToolStripMenuItem.Text = "Приборы учета"
        '
        'ПриборыToolStripMenuItem
        '
        Me.ПриборыToolStripMenuItem.Name = "ПриборыToolStripMenuItem"
        Me.ПриборыToolStripMenuItem.Size = New System.Drawing.Size(137, 22)
        Me.ПриборыToolStripMenuItem.Text = "Приборы"
        '
        'ДокументыToolStripMenuItem1
        '
        Me.ДокументыToolStripMenuItem1.Name = "ДокументыToolStripMenuItem1"
        Me.ДокументыToolStripMenuItem1.Size = New System.Drawing.Size(137, 22)
        Me.ДокументыToolStripMenuItem1.Text = "Документы"
        '
        'КвитанцииToolStripMenuItem
        '
        Me.КвитанцииToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ШаблоныToolStripMenuItem, Me.ПараметрыToolStripMenuItem})
        Me.КвитанцииToolStripMenuItem.Name = "КвитанцииToolStripMenuItem"
        Me.КвитанцииToolStripMenuItem.Size = New System.Drawing.Size(78, 20)
        Me.КвитанцииToolStripMenuItem.Text = "Квитанции"
        '
        'ШаблоныToolStripMenuItem
        '
        Me.ШаблоныToolStripMenuItem.Name = "ШаблоныToolStripMenuItem"
        Me.ШаблоныToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.ШаблоныToolStripMenuItem.Text = "Шаблоны"
        '
        'ПараметрыToolStripMenuItem
        '
        Me.ПараметрыToolStripMenuItem.Name = "ПараметрыToolStripMenuItem"
        Me.ПараметрыToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.ПараметрыToolStripMenuItem.Text = "Параметры"
        '
        'БазаДанныхToolStripMenuItem
        '
        Me.БазаДанныхToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.АрхивацияToolStripMenuItem, Me.ВосстановлениеToolStripMenuItem})
        Me.БазаДанныхToolStripMenuItem.Name = "БазаДанныхToolStripMenuItem"
        Me.БазаДанныхToolStripMenuItem.Size = New System.Drawing.Size(86, 20)
        Me.БазаДанныхToolStripMenuItem.Text = "База данных"
        '
        'АрхивацияToolStripMenuItem
        '
        Me.АрхивацияToolStripMenuItem.Name = "АрхивацияToolStripMenuItem"
        Me.АрхивацияToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.АрхивацияToolStripMenuItem.Text = "Архивация"
        '
        'ВосстановлениеToolStripMenuItem
        '
        Me.ВосстановлениеToolStripMenuItem.Name = "ВосстановлениеToolStripMenuItem"
        Me.ВосстановлениеToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.ВосстановлениеToolStripMenuItem.Text = "Восстановление"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label2.Location = New System.Drawing.Point(230, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Label2"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.Location = New System.Drawing.Point(11, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Label3"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label4.Location = New System.Drawing.Point(77, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 17)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Label4"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label5.Location = New System.Drawing.Point(11, 113)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 17)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Label5"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label6.Location = New System.Drawing.Point(77, 113)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Label6"
        '
        'FormAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(1004, 507)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "FormAdmin"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Администрирование"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ЖилфондToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УслугиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЛюдиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОператорыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриборыУчетыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УлицыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДомаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПоставщикиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ТарифыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ТипРегистрацииToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДанныеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents РолиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДокументыToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents КвитанцииToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ШаблоныToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПараметрыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents ТехучасткиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ТипыДокументовToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ТипыОтношенийToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИсполнителиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УслугиToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ПриборыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents БазаДанныхToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents АрхивацияToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВосстановлениеToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
