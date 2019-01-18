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
        Me.GroupBox1.Location = New System.Drawing.Point(267, 130)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(357, 242)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lblPFC
        '
        Me.lblPFC.AutoSize = True
        Me.lblPFC.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPFC.Location = New System.Drawing.Point(15, 66)
        Me.lblPFC.Name = "lblPFC"
        Me.lblPFC.Size = New System.Drawing.Size(63, 15)
        Me.lblPFC.TabIndex = 13
        Me.lblPFC.Text = "Select PFC"
        '
        'cboPFC
        '
        Me.cboPFC.FormattingEnabled = True
        Me.cboPFC.Location = New System.Drawing.Point(187, 64)
        Me.cboPFC.Name = "cboPFC"
        Me.cboPFC.Size = New System.Drawing.Size(154, 21)
        Me.cboPFC.TabIndex = 12
        '
        'cmdReport
        '
        Me.cmdReport.Location = New System.Drawing.Point(187, 179)
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
        Me.Label3.Location = New System.Drawing.Point(15, 139)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(123, 15)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Select Product Rating"
        '
        'cboProductRating
        '
        Me.cboProductRating.FormattingEnabled = True
        Me.cboProductRating.Location = New System.Drawing.Point(187, 137)
        Me.cboProductRating.Name = "cboProductRating"
        Me.cboProductRating.Size = New System.Drawing.Size(154, 21)
        Me.cboProductRating.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 103)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(129, 15)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Select Releasing Phase"
        '
        'cboReleasingPhase
        '
        Me.cboReleasingPhase.FormattingEnabled = True
        Me.cboReleasingPhase.Location = New System.Drawing.Point(187, 101)
        Me.cboReleasingPhase.Name = "cboReleasingPhase"
        Me.cboReleasingPhase.Size = New System.Drawing.Size(154, 21)
        Me.cboReleasingPhase.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 15)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Select the Customer"
        '
        'cboSelectCustomer
        '
        Me.cboSelectCustomer.FormattingEnabled = True
        Me.cboSelectCustomer.Location = New System.Drawing.Point(187, 28)
        Me.cboSelectCustomer.Name = "cboSelectCustomer"
        Me.cboSelectCustomer.Size = New System.Drawing.Size(154, 21)
        Me.cboSelectCustomer.TabIndex = 5
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.Red
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExit.Location = New System.Drawing.Point(18, 179)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(154, 41)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Exit Application"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(882, 503)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
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
