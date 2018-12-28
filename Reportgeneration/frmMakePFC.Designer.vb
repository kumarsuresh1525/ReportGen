<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMakePFC
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.cmdPFC = New System.Windows.Forms.Button()
        Me.cmdAddBOM = New System.Windows.Forms.Button()
        Me.cmdAddProcess = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.dgvSelectedProcess = New System.Windows.Forms.DataGridView()
        Me.sno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pname = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.qmp = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.qms = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.freqAtInit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.freq = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.measuringTool = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.checkedBy = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.recordingSheet = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.machineJigTool = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.DGVinfo = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.gbProcess = New System.Windows.Forms.GroupBox()
        Me.cboSubProcessName = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboProcessName = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtProcessNo = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dgvSelectedProcess, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DGVinfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbProcess.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnBack)
        Me.GroupBox1.Controls.Add(Me.lblMessage)
        Me.GroupBox1.Controls.Add(Me.cmdPFC)
        Me.GroupBox1.Controls.Add(Me.cmdAddBOM)
        Me.GroupBox1.Controls.Add(Me.cmdAddProcess)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.gbProcess)
        Me.GroupBox1.Location = New System.Drawing.Point(20, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1338, 699)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'btnBack
        '
        Me.btnBack.BackColor = System.Drawing.Color.Red
        Me.btnBack.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBack.Font = New System.Drawing.Font("Calibri", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.Location = New System.Drawing.Point(1202, 644)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(119, 42)
        Me.btnBack.TabIndex = 29
        Me.btnBack.Text = "Back"
        Me.btnBack.UseVisualStyleBackColor = False
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Font = New System.Drawing.Font("Calibri", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.Location = New System.Drawing.Point(602, 87)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(62, 23)
        Me.lblMessage.TabIndex = 2
        Me.lblMessage.Text = "Label4"
        '
        'cmdPFC
        '
        Me.cmdPFC.Location = New System.Drawing.Point(1202, 440)
        Me.cmdPFC.Name = "cmdPFC"
        Me.cmdPFC.Size = New System.Drawing.Size(119, 39)
        Me.cmdPFC.TabIndex = 28
        Me.cmdPFC.Text = "Save PFC"
        Me.cmdPFC.UseVisualStyleBackColor = True
        '
        'cmdAddBOM
        '
        Me.cmdAddBOM.Location = New System.Drawing.Point(1202, 377)
        Me.cmdAddBOM.Name = "cmdAddBOM"
        Me.cmdAddBOM.Size = New System.Drawing.Size(119, 39)
        Me.cmdAddBOM.TabIndex = 27
        Me.cmdAddBOM.Text = "Add BOM"
        Me.cmdAddBOM.UseVisualStyleBackColor = True
        '
        'cmdAddProcess
        '
        Me.cmdAddProcess.Location = New System.Drawing.Point(1096, 36)
        Me.cmdAddProcess.Name = "cmdAddProcess"
        Me.cmdAddProcess.Size = New System.Drawing.Size(119, 39)
        Me.cmdAddProcess.TabIndex = 26
        Me.cmdAddProcess.Text = "Add Process"
        Me.cmdAddProcess.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.TreeView1)
        Me.GroupBox4.Controls.Add(Me.dgvSelectedProcess)
        Me.GroupBox4.Location = New System.Drawing.Point(16, 377)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(1046, 309)
        Me.GroupBox4.TabIndex = 25
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Selected Process"
        '
        'TreeView1
        '
        Me.TreeView1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeView1.Location = New System.Drawing.Point(20, 19)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(1008, 273)
        Me.TreeView1.TabIndex = 2
        '
        'dgvSelectedProcess
        '
        Me.dgvSelectedProcess.AllowUserToAddRows = False
        Me.dgvSelectedProcess.AllowUserToDeleteRows = False
        Me.dgvSelectedProcess.AllowUserToResizeColumns = False
        Me.dgvSelectedProcess.AllowUserToResizeRows = False
        Me.dgvSelectedProcess.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSelectedProcess.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvSelectedProcess.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSelectedProcess.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.sno, Me.pno, Me.pname, Me.qmp, Me.qms, Me.freqAtInit, Me.freq, Me.measuringTool, Me.checkedBy, Me.recordingSheet, Me.machineJigTool})
        Me.dgvSelectedProcess.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvSelectedProcess.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvSelectedProcess.Location = New System.Drawing.Point(20, 19)
        Me.dgvSelectedProcess.Name = "dgvSelectedProcess"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSelectedProcess.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvSelectedProcess.RowHeadersVisible = False
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.Color.White
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.Color.Black
        Me.dgvSelectedProcess.RowsDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvSelectedProcess.Size = New System.Drawing.Size(1008, 273)
        Me.dgvSelectedProcess.TabIndex = 24
        '
        'sno
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.sno.DefaultCellStyle = DataGridViewCellStyle2
        Me.sno.HeaderText = "S.No"
        Me.sno.Name = "sno"
        Me.sno.ReadOnly = True
        Me.sno.Width = 50
        '
        'pno
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.White
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.pno.DefaultCellStyle = DataGridViewCellStyle3
        Me.pno.HeaderText = "Process No"
        Me.pno.Name = "pno"
        Me.pno.ReadOnly = True
        Me.pno.Width = 150
        '
        'pname
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.pname.DefaultCellStyle = DataGridViewCellStyle4
        Me.pname.HeaderText = "Process Name"
        Me.pname.Name = "pname"
        Me.pname.ReadOnly = True
        Me.pname.Width = 400
        '
        'qmp
        '
        Me.qmp.HeaderText = "QMP"
        Me.qmp.Name = "qmp"
        Me.qmp.ReadOnly = True
        '
        'qms
        '
        Me.qms.HeaderText = "QMS"
        Me.qms.Name = "qms"
        Me.qms.ReadOnly = True
        '
        'freqAtInit
        '
        Me.freqAtInit.HeaderText = "Freq. At Init"
        Me.freqAtInit.Name = "freqAtInit"
        Me.freqAtInit.ReadOnly = True
        '
        'freq
        '
        Me.freq.HeaderText = "Frequency"
        Me.freq.Name = "freq"
        Me.freq.ReadOnly = True
        '
        'measuringTool
        '
        Me.measuringTool.HeaderText = "Measuring Tool"
        Me.measuringTool.Name = "measuringTool"
        Me.measuringTool.ReadOnly = True
        '
        'checkedBy
        '
        Me.checkedBy.HeaderText = "Checked BY"
        Me.checkedBy.Name = "checkedBy"
        Me.checkedBy.ReadOnly = True
        '
        'recordingSheet
        '
        Me.recordingSheet.HeaderText = "Recording Sheet"
        Me.recordingSheet.Name = "recordingSheet"
        Me.recordingSheet.ReadOnly = True
        '
        'machineJigTool
        '
        Me.machineJigTool.HeaderText = "Machine Jig Tool"
        Me.machineJigTool.Name = "machineJigTool"
        Me.machineJigTool.ReadOnly = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DGVinfo)
        Me.GroupBox3.Location = New System.Drawing.Point(16, 106)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1305, 265)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Select Model"
        '
        'DGVinfo
        '
        Me.DGVinfo.AllowUserToAddRows = False
        Me.DGVinfo.AllowUserToDeleteRows = False
        Me.DGVinfo.AllowUserToResizeColumns = False
        Me.DGVinfo.AllowUserToResizeRows = False
        Me.DGVinfo.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVinfo.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.DGVinfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVinfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column12, Me.Column11, Me.Column6, Me.Column7, Me.Column8, Me.Column9, Me.Column10})
        Me.DGVinfo.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle15.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGVinfo.DefaultCellStyle = DataGridViewCellStyle15
        Me.DGVinfo.Location = New System.Drawing.Point(6, 18)
        Me.DGVinfo.Name = "DGVinfo"
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVinfo.RowHeadersDefaultCellStyle = DataGridViewCellStyle16
        Me.DGVinfo.RowHeadersVisible = False
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle17.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle17.SelectionBackColor = System.Drawing.Color.White
        DataGridViewCellStyle17.SelectionForeColor = System.Drawing.Color.Black
        Me.DGVinfo.RowsDefaultCellStyle = DataGridViewCellStyle17
        Me.DGVinfo.Size = New System.Drawing.Size(1282, 228)
        Me.DGVinfo.TabIndex = 23
        '
        'Column1
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle9.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle9.NullValue = False
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Column1.DefaultCellStyle = DataGridViewCellStyle9
        Me.Column1.HeaderText = "Select"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Column1.Width = 50
        '
        'Column2
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.Color.White
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Column2.DefaultCellStyle = DataGridViewCellStyle10
        Me.Column2.HeaderText = "Part Number"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 200
        '
        'Column3
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle11
        Me.Column3.HeaderText = "Part Name"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 300
        '
        'Column4
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle12.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle12
        Me.Column4.HeaderText = "Qty"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.Width = 50
        '
        'Column5
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle13.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle13.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Column5.DefaultCellStyle = DataGridViewCellStyle13
        Me.Column5.HeaderText = "XYZ"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        Me.Column5.Width = 50
        '
        'Column12
        '
        Me.Column12.HeaderText = "Product Number"
        Me.Column12.Name = "Column12"
        Me.Column12.ReadOnly = True
        Me.Column12.Width = 200
        '
        'Column11
        '
        Me.Column11.HeaderText = "Product Name"
        Me.Column11.Name = "Column11"
        Me.Column11.ReadOnly = True
        Me.Column11.Width = 200
        '
        'Column6
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle14.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle14.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Column6.DefaultCellStyle = DataGridViewCellStyle14
        Me.Column6.HeaderText = "Customer Name"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        Me.Column6.Width = 200
        '
        'Column7
        '
        Me.Column7.HeaderText = "Customer Number"
        Me.Column7.Name = "Column7"
        Me.Column7.ReadOnly = True
        Me.Column7.Width = 200
        '
        'Column8
        '
        Me.Column8.HeaderText = "Customer"
        Me.Column8.Name = "Column8"
        Me.Column8.ReadOnly = True
        '
        'Column9
        '
        Me.Column9.HeaderText = "Model"
        Me.Column9.Name = "Column9"
        Me.Column9.ReadOnly = True
        '
        'Column10
        '
        Me.Column10.HeaderText = "Part Assembly"
        Me.Column10.Name = "Column10"
        Me.Column10.ReadOnly = True
        Me.Column10.Width = 200
        '
        'gbProcess
        '
        Me.gbProcess.Controls.Add(Me.cboSubProcessName)
        Me.gbProcess.Controls.Add(Me.Label3)
        Me.gbProcess.Controls.Add(Me.cboProcessName)
        Me.gbProcess.Controls.Add(Me.Label2)
        Me.gbProcess.Controls.Add(Me.txtProcessNo)
        Me.gbProcess.Controls.Add(Me.Label1)
        Me.gbProcess.Location = New System.Drawing.Point(16, 19)
        Me.gbProcess.Name = "gbProcess"
        Me.gbProcess.Size = New System.Drawing.Size(1059, 65)
        Me.gbProcess.TabIndex = 1
        Me.gbProcess.TabStop = False
        '
        'cboSubProcessName
        '
        Me.cboSubProcessName.FormattingEnabled = True
        Me.cboSubProcessName.Location = New System.Drawing.Point(700, 27)
        Me.cboSubProcessName.Name = "cboSubProcessName"
        Me.cboSubProcessName.Size = New System.Drawing.Size(338, 21)
        Me.cboSubProcessName.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(587, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 15)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Sub Process Name"
        '
        'cboProcessName
        '
        Me.cboProcessName.FormattingEnabled = True
        Me.cboProcessName.Location = New System.Drawing.Point(279, 27)
        Me.cboProcessName.Name = "cboProcessName"
        Me.cboProcessName.Size = New System.Drawing.Size(250, 21)
        Me.cboProcessName.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(189, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 15)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Process Name"
        '
        'txtProcessNo
        '
        Me.txtProcessNo.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProcessNo.Location = New System.Drawing.Point(103, 27)
        Me.txtProcessNo.Name = "txtProcessNo"
        Me.txtProcessNo.ReadOnly = True
        Me.txtProcessNo.Size = New System.Drawing.Size(41, 23)
        Me.txtProcessNo.TabIndex = 1
        Me.txtProcessNo.Text = "0"
        Me.txtProcessNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(27, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Process No."
        '
        'Timer1
        '
        '
        'frmMakePFC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1354, 741)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmMakePFC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.dgvSelectedProcess, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DGVinfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbProcess.ResumeLayout(False)
        Me.gbProcess.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents gbProcess As System.Windows.Forms.GroupBox
    Friend WithEvents txtProcessNo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboSubProcessName As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboProcessName As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents DGVinfo As System.Windows.Forms.DataGridView
    Public WithEvents dgvSelectedProcess As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdAddProcess As System.Windows.Forms.Button
    Friend WithEvents cmdAddBOM As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cmdPFC As System.Windows.Forms.Button
    Friend WithEvents sno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents pno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents pname As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents qmp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents qms As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents freqAtInit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents freq As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents measuringTool As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents checkedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents recordingSheet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents machineJigTool As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents btnBack As System.Windows.Forms.Button
End Class
