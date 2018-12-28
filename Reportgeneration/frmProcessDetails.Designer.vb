<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProcessDetails
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
        Me.lbQMP = New System.Windows.Forms.ListBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lbQMS = New System.Windows.Forms.ListBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lbFrequency = New System.Windows.Forms.ListBox()
        Me.lbMeasuringTool = New System.Windows.Forms.ListBox()
        Me.lbRecordingSheet = New System.Windows.Forms.ListBox()
        Me.lbCheckedBy = New System.Windows.Forms.ListBox()
        Me.ListBox5 = New System.Windows.Forms.ListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbQMP
        '
        Me.lbQMP.FormattingEnabled = True
        Me.lbQMP.Location = New System.Drawing.Point(15, 29)
        Me.lbQMP.Name = "lbQMP"
        Me.lbQMP.Size = New System.Drawing.Size(266, 108)
        Me.lbQMP.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lbQMP)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(299, 163)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Quality Manufacturing Parameter"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lbQMS)
        Me.GroupBox2.Location = New System.Drawing.Point(331, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(299, 163)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Quality Manufacturing Specification"
        '
        'lbQMS
        '
        Me.lbQMS.FormattingEnabled = True
        Me.lbQMS.Location = New System.Drawing.Point(15, 29)
        Me.lbQMS.Name = "lbQMS"
        Me.lbQMS.Size = New System.Drawing.Size(266, 108)
        Me.lbQMS.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ListBox5)
        Me.GroupBox3.Controls.Add(Me.lbRecordingSheet)
        Me.GroupBox3.Controls.Add(Me.lbCheckedBy)
        Me.GroupBox3.Controls.Add(Me.lbMeasuringTool)
        Me.GroupBox3.Controls.Add(Me.lbFrequency)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 199)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(618, 163)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Control Quality Specification(Frequency, Measuring Tool, Checked by, Recording Sh" & _
            "eet)"
        '
        'lbFrequency
        '
        Me.lbFrequency.FormattingEnabled = True
        Me.lbFrequency.Location = New System.Drawing.Point(15, 29)
        Me.lbFrequency.Name = "lbFrequency"
        Me.lbFrequency.Size = New System.Drawing.Size(94, 108)
        Me.lbFrequency.TabIndex = 0
        '
        'lbMeasuringTool
        '
        Me.lbMeasuringTool.FormattingEnabled = True
        Me.lbMeasuringTool.Location = New System.Drawing.Point(126, 29)
        Me.lbMeasuringTool.Name = "lbMeasuringTool"
        Me.lbMeasuringTool.Size = New System.Drawing.Size(102, 108)
        Me.lbMeasuringTool.TabIndex = 1
        '
        'lbRecordingSheet
        '
        Me.lbRecordingSheet.FormattingEnabled = True
        Me.lbRecordingSheet.Location = New System.Drawing.Point(370, 29)
        Me.lbRecordingSheet.Name = "lbRecordingSheet"
        Me.lbRecordingSheet.Size = New System.Drawing.Size(106, 108)
        Me.lbRecordingSheet.TabIndex = 3
        '
        'lbCheckedBy
        '
        Me.lbCheckedBy.FormattingEnabled = True
        Me.lbCheckedBy.Location = New System.Drawing.Point(247, 29)
        Me.lbCheckedBy.Name = "lbCheckedBy"
        Me.lbCheckedBy.Size = New System.Drawing.Size(101, 108)
        Me.lbCheckedBy.TabIndex = 2
        '
        'ListBox5
        '
        Me.ListBox5.FormattingEnabled = True
        Me.ListBox5.Location = New System.Drawing.Point(496, 29)
        Me.ListBox5.Name = "ListBox5"
        Me.ListBox5.Size = New System.Drawing.Size(104, 108)
        Me.ListBox5.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(652, 22)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 35)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmProcessDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(881, 415)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmProcessDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Process Details"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lbQMP As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lbQMS As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ListBox5 As System.Windows.Forms.ListBox
    Friend WithEvents lbRecordingSheet As System.Windows.Forms.ListBox
    Friend WithEvents lbCheckedBy As System.Windows.Forms.ListBox
    Friend WithEvents lbMeasuringTool As System.Windows.Forms.ListBox
    Friend WithEvents lbFrequency As System.Windows.Forms.ListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
