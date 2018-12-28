<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblPFC = New System.Windows.Forms.Label()
        Me.cboPFC = New System.Windows.Forms.ComboBox()
        Me.cmdReport = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboProductRating = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboReleasingPhase = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboSelectCustomer = New System.Windows.Forms.ComboBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnQMS = New System.Windows.Forms.Button()
        Me.pqcsSet = New System.Windows.Forms.Button()
        Me.pSet = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblPFC)
        Me.GroupBox1.Controls.Add(Me.cboPFC)
        Me.GroupBox1.Controls.Add(Me.cmdReport)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.cboProductRating)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cboReleasingPhase)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cboSelectCustomer)
        Me.GroupBox1.Controls.Add(Me.btnExit)
        Me.GroupBox1.Location = New System.Drawing.Point(267, 108)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(357, 267)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lblPFC
        '
        Me.lblPFC.AutoSize = True
        Me.lblPFC.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPFC.Location = New System.Drawing.Point(15, 171)
        Me.lblPFC.Name = "lblPFC"
        Me.lblPFC.Size = New System.Drawing.Size(63, 15)
        Me.lblPFC.TabIndex = 13
        Me.lblPFC.Text = "Select PFC"
        '
        'cboPFC
        '
        Me.cboPFC.FormattingEnabled = True
        Me.cboPFC.Location = New System.Drawing.Point(187, 169)
        Me.cboPFC.Name = "cboPFC"
        Me.cboPFC.Size = New System.Drawing.Size(154, 21)
        Me.cboPFC.TabIndex = 12
        '
        'cmdReport
        '
        Me.cmdReport.Location = New System.Drawing.Point(187, 196)
        Me.cmdReport.Name = "cmdReport"
        Me.cmdReport.Size = New System.Drawing.Size(154, 41)
        Me.cmdReport.TabIndex = 11
        Me.cmdReport.Text = "Create Report"
        Me.cmdReport.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(15, 127)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(123, 15)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Select Product Rating"
        '
        'cboProductRating
        '
        Me.cboProductRating.FormattingEnabled = True
        Me.cboProductRating.Location = New System.Drawing.Point(187, 125)
        Me.cboProductRating.Name = "cboProductRating"
        Me.cboProductRating.Size = New System.Drawing.Size(154, 21)
        Me.cboProductRating.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(129, 15)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Select Releasing Phase"
        '
        'cboReleasingPhase
        '
        Me.cboReleasingPhase.FormattingEnabled = True
        Me.cboReleasingPhase.Location = New System.Drawing.Point(187, 80)
        Me.cboReleasingPhase.Name = "cboReleasingPhase"
        Me.cboReleasingPhase.Size = New System.Drawing.Size(154, 21)
        Me.cboReleasingPhase.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 15)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Select the Customer"
        '
        'cboSelectCustomer
        '
        Me.cboSelectCustomer.FormattingEnabled = True
        Me.cboSelectCustomer.Location = New System.Drawing.Point(140, 37)
        Me.cboSelectCustomer.Name = "cboSelectCustomer"
        Me.cboSelectCustomer.Size = New System.Drawing.Size(201, 21)
        Me.cboSelectCustomer.TabIndex = 5
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.Red
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExit.Location = New System.Drawing.Point(18, 196)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(154, 41)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Exit Application"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnQMS
        '
        Me.btnQMS.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnQMS.Location = New System.Drawing.Point(693, 194)
        Me.btnQMS.Name = "btnQMS"
        Me.btnQMS.Size = New System.Drawing.Size(74, 78)
        Me.btnQMS.TabIndex = 4
        Me.btnQMS.Text = "QMS"
        Me.btnQMS.UseVisualStyleBackColor = True
        '
        'pqcsSet
        '
        Me.pqcsSet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.pqcsSet.Location = New System.Drawing.Point(657, 55)
        Me.pqcsSet.Name = "pqcsSet"
        Me.pqcsSet.Size = New System.Drawing.Size(74, 78)
        Me.pqcsSet.TabIndex = 2
        Me.pqcsSet.Text = "QMP"
        Me.pqcsSet.UseVisualStyleBackColor = True
        '
        'pSet
        '
        Me.pSet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.pSet.Location = New System.Drawing.Point(675, 15)
        Me.pSet.Name = "pSet"
        Me.pSet.Size = New System.Drawing.Size(151, 78)
        Me.pSet.TabIndex = 1
        Me.pSet.Text = "Add/Edit Process"
        Me.pSet.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(373, 40)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(317, 53)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Generate Report"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(882, 503)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.pSet)
        Me.Controls.Add(Me.pqcsSet)
        Me.Controls.Add(Me.btnQMS)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents pqcsSet As System.Windows.Forms.Button
    Friend WithEvents pSet As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnQMS As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboSelectCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboReleasingPhase As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboProductRating As System.Windows.Forms.ComboBox
    Friend WithEvents cmdReport As System.Windows.Forms.Button
    Friend WithEvents lblPFC As System.Windows.Forms.Label
    Friend WithEvents cboPFC As System.Windows.Forms.ComboBox

End Class
