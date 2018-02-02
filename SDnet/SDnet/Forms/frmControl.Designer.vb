<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmControl
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmControl))
        Me.lstDisplay = New System.Windows.Forms.ListBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkModoTransparente = New System.Windows.Forms.CheckBox()
        Me.chkInicializarCanales = New System.Windows.Forms.CheckBox()
        Me.chkConBases = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbModoDespacho = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkGrabarLog = New System.Windows.Forms.CheckBox()
        Me.chkStopList = New System.Windows.Forms.CheckBox()
        Me.botCancelar = New System.Windows.Forms.Button()
        Me.botAceptar = New System.Windows.Forms.Button()
        Me.tmrViajesPendientes = New System.Windows.Forms.Timer(Me.components)
        Me.tmrViajesAgendados = New System.Windows.Forms.Timer(Me.components)
        Me.tmrHoraAutoflot = New System.Windows.Forms.Timer(Me.components)
        Me.tmrDesactivar = New System.Windows.Forms.Timer(Me.components)
        Me.tmrAlarmasCanceladas = New System.Windows.Forms.Timer(Me.components)
        Me.tmrViajesPendientes2 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lstDisplay
        '
        Me.lstDisplay.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lstDisplay.FormattingEnabled = True
        Me.lstDisplay.Location = New System.Drawing.Point(1, 12)
        Me.lstDisplay.Name = "lstDisplay"
        Me.lstDisplay.Size = New System.Drawing.Size(464, 407)
        Me.lstDisplay.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkModoTransparente)
        Me.GroupBox1.Controls.Add(Me.chkInicializarCanales)
        Me.GroupBox1.Controls.Add(Me.chkConBases)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cmbModoDespacho)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(471, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(228, 155)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Opciones de Despacho"
        '
        'chkModoTransparente
        '
        Me.chkModoTransparente.AutoSize = True
        Me.chkModoTransparente.Location = New System.Drawing.Point(17, 121)
        Me.chkModoTransparente.Name = "chkModoTransparente"
        Me.chkModoTransparente.Size = New System.Drawing.Size(119, 17)
        Me.chkModoTransparente.TabIndex = 4
        Me.chkModoTransparente.Text = "Modo Transparente"
        Me.chkModoTransparente.UseVisualStyleBackColor = True
        '
        'chkInicializarCanales
        '
        Me.chkInicializarCanales.AutoSize = True
        Me.chkInicializarCanales.Location = New System.Drawing.Point(17, 72)
        Me.chkInicializarCanales.Name = "chkInicializarCanales"
        Me.chkInicializarCanales.Size = New System.Drawing.Size(110, 17)
        Me.chkInicializarCanales.TabIndex = 3
        Me.chkInicializarCanales.Text = "Inicializar Canales"
        Me.chkInicializarCanales.UseVisualStyleBackColor = True
        '
        'chkConBases
        '
        Me.chkConBases.AutoSize = True
        Me.chkConBases.Location = New System.Drawing.Point(17, 96)
        Me.chkConBases.Name = "chkConBases"
        Me.chkConBases.Size = New System.Drawing.Size(77, 17)
        Me.chkConBases.TabIndex = 2
        Me.chkConBases.Text = "Con Bases"
        Me.chkConBases.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Modo:"
        '
        'cmbModoDespacho
        '
        Me.cmbModoDespacho.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbModoDespacho.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmbModoDespacho.FormattingEnabled = True
        Me.cmbModoDespacho.Items.AddRange(New Object() {"Sin Adyacencias", "Adyacencias 1", "Adyacencias 2"})
        Me.cmbModoDespacho.Location = New System.Drawing.Point(56, 32)
        Me.cmbModoDespacho.Name = "cmbModoDespacho"
        Me.cmbModoDespacho.Size = New System.Drawing.Size(146, 21)
        Me.cmbModoDespacho.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkGrabarLog)
        Me.GroupBox2.Controls.Add(Me.chkStopList)
        Me.GroupBox2.Location = New System.Drawing.Point(471, 175)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(228, 91)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Opciones de Operación"
        '
        'chkGrabarLog
        '
        Me.chkGrabarLog.AutoSize = True
        Me.chkGrabarLog.Location = New System.Drawing.Point(12, 31)
        Me.chkGrabarLog.Name = "chkGrabarLog"
        Me.chkGrabarLog.Size = New System.Drawing.Size(144, 17)
        Me.chkGrabarLog.TabIndex = 5
        Me.chkGrabarLog.Text = "Grabar Log de Actividad "
        Me.chkGrabarLog.UseVisualStyleBackColor = True
        '
        'chkStopList
        '
        Me.chkStopList.AutoSize = True
        Me.chkStopList.Location = New System.Drawing.Point(12, 58)
        Me.chkStopList.Name = "chkStopList"
        Me.chkStopList.Size = New System.Drawing.Size(128, 17)
        Me.chkStopList.TabIndex = 4
        Me.chkStopList.Text = "Detener Visualización"
        Me.chkStopList.UseVisualStyleBackColor = True
        '
        'botCancelar
        '
        Me.botCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.botCancelar.Location = New System.Drawing.Point(596, 272)
        Me.botCancelar.Name = "botCancelar"
        Me.botCancelar.Size = New System.Drawing.Size(100, 25)
        Me.botCancelar.TabIndex = 8
        Me.botCancelar.Text = "Cancelar"
        Me.botCancelar.UseVisualStyleBackColor = True
        '
        'botAceptar
        '
        Me.botAceptar.Location = New System.Drawing.Point(490, 272)
        Me.botAceptar.Name = "botAceptar"
        Me.botAceptar.Size = New System.Drawing.Size(100, 25)
        Me.botAceptar.TabIndex = 7
        Me.botAceptar.Text = "Cerrar Servidor"
        Me.botAceptar.UseVisualStyleBackColor = True
        '
        'tmrViajesPendientes
        '
        '
        'tmrViajesAgendados
        '
        '
        'tmrHoraAutoflot
        '
        '
        'tmrDesactivar
        '
        '
        'tmrAlarmasCanceladas
        '
        '
        'frmControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(708, 441)
        Me.Controls.Add(Me.botCancelar)
        Me.Controls.Add(Me.botAceptar)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lstDisplay)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmControl"
        Me.Text = "Servidor de Despachos"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstDisplay As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkModoTransparente As System.Windows.Forms.CheckBox
    Friend WithEvents chkInicializarCanales As System.Windows.Forms.CheckBox
    Friend WithEvents chkConBases As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbModoDespacho As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkGrabarLog As System.Windows.Forms.CheckBox
    Friend WithEvents chkStopList As System.Windows.Forms.CheckBox
    Friend WithEvents botCancelar As System.Windows.Forms.Button
    Friend WithEvents botAceptar As System.Windows.Forms.Button
    Friend WithEvents tmrViajesPendientes As System.Windows.Forms.Timer
    Friend WithEvents tmrViajesAgendados As System.Windows.Forms.Timer
    Friend WithEvents tmrHoraAutoflot As System.Windows.Forms.Timer
    Friend WithEvents tmrDesactivar As System.Windows.Forms.Timer
    Friend WithEvents tmrAlarmasCanceladas As System.Windows.Forms.Timer
    Friend WithEvents tmrViajesPendientes2 As System.Windows.Forms.Timer

End Class
