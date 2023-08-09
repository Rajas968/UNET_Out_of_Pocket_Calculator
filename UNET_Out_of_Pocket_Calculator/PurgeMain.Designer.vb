<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Purgemain
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
        Me.chkSkipParsing = New System.Windows.Forms.CheckBox()
        Me.chkExcludeRio = New System.Windows.Forms.CheckBox()
        Me.txtRestFindBatchSize = New System.Windows.Forms.TextBox()
        Me.chkMaintainRecordInfo = New System.Windows.Forms.CheckBox()
        Me.chkInclBoneyard = New System.Windows.Forms.CheckBox()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.txtEmployeeID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblTestServer = New System.Windows.Forms.Label()
        Me.PurgeStartDT = New System.Windows.Forms.DateTimePicker()
        Me.PurgeEndDT = New System.Windows.Forms.DateTimePicker()
        Me.SuspendLayout()
        '
        'chkSkipParsing
        '
        Me.chkSkipParsing.AutoSize = True
        Me.chkSkipParsing.Location = New System.Drawing.Point(550, 172)
        Me.chkSkipParsing.Name = "chkSkipParsing"
        Me.chkSkipParsing.Size = New System.Drawing.Size(105, 20)
        Me.chkSkipParsing.TabIndex = 19
        Me.chkSkipParsing.Text = "Skip Parsing"
        Me.chkSkipParsing.UseVisualStyleBackColor = True
        '
        'chkExcludeRio
        '
        Me.chkExcludeRio.AutoSize = True
        Me.chkExcludeRio.Location = New System.Drawing.Point(550, 135)
        Me.chkExcludeRio.Name = "chkExcludeRio"
        Me.chkExcludeRio.Size = New System.Drawing.Size(101, 20)
        Me.chkExcludeRio.TabIndex = 18
        Me.chkExcludeRio.Text = "Exclude Rio"
        Me.chkExcludeRio.UseVisualStyleBackColor = True
        '
        'txtRestFindBatchSize
        '
        Me.txtRestFindBatchSize.Location = New System.Drawing.Point(550, 107)
        Me.txtRestFindBatchSize.Name = "txtRestFindBatchSize"
        Me.txtRestFindBatchSize.Size = New System.Drawing.Size(33, 22)
        Me.txtRestFindBatchSize.TabIndex = 17
        Me.txtRestFindBatchSize.Text = "50"
        '
        'chkMaintainRecordInfo
        '
        Me.chkMaintainRecordInfo.AutoSize = True
        Me.chkMaintainRecordInfo.Location = New System.Drawing.Point(550, 337)
        Me.chkMaintainRecordInfo.Name = "chkMaintainRecordInfo"
        Me.chkMaintainRecordInfo.Size = New System.Drawing.Size(214, 20)
        Me.chkMaintainRecordInfo.TabIndex = 16
        Me.chkMaintainRecordInfo.Text = "Maintain Additional Record Info"
        Me.chkMaintainRecordInfo.UseVisualStyleBackColor = True
        '
        'chkInclBoneyard
        '
        Me.chkInclBoneyard.AutoSize = True
        Me.chkInclBoneyard.Location = New System.Drawing.Point(550, 290)
        Me.chkInclBoneyard.Name = "chkInclBoneyard"
        Me.chkInclBoneyard.Size = New System.Drawing.Size(244, 20)
        Me.chkInclBoneyard.TabIndex = 15
        Me.chkInclBoneyard.Text = "Include results purged > 8 years ago"
        Me.chkInclBoneyard.UseVisualStyleBackColor = True
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(550, 240)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(100, 22)
        Me.txtFirstName.TabIndex = 14
        '
        'txtEmployeeID
        '
        Me.txtEmployeeID.Location = New System.Drawing.Point(271, 152)
        Me.txtEmployeeID.Name = "txtEmployeeID"
        Me.txtEmployeeID.Size = New System.Drawing.Size(224, 22)
        Me.txtEmployeeID.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(146, 152)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 16)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Enter Employee ID"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(301, 215)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(146, 33)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Get Claim Data"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lblTestServer
        '
        Me.lblTestServer.AutoSize = True
        Me.lblTestServer.Location = New System.Drawing.Point(262, 94)
        Me.lblTestServer.Name = "lblTestServer"
        Me.lblTestServer.Size = New System.Drawing.Size(143, 16)
        Me.lblTestServer.TabIndex = 13
        Me.lblTestServer.Text = """UAT"" TEST SERVER"
        '
        'PurgeStartDT
        '
        Me.PurgeStartDT.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PurgeStartDT.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.PurgeStartDT.Location = New System.Drawing.Point(77, 305)
        Me.PurgeStartDT.Name = "PurgeStartDT"
        Me.PurgeStartDT.Size = New System.Drawing.Size(153, 28)
        Me.PurgeStartDT.TabIndex = 20
        '
        'PurgeEndDT
        '
        Me.PurgeEndDT.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PurgeEndDT.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.PurgeEndDT.Location = New System.Drawing.Point(265, 305)
        Me.PurgeEndDT.Name = "PurgeEndDT"
        Me.PurgeEndDT.Size = New System.Drawing.Size(157, 28)
        Me.PurgeEndDT.TabIndex = 21
        '
        'Purgemain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(858, 450)
        Me.Controls.Add(Me.PurgeEndDT)
        Me.Controls.Add(Me.PurgeStartDT)
        Me.Controls.Add(Me.chkSkipParsing)
        Me.Controls.Add(Me.chkExcludeRio)
        Me.Controls.Add(Me.txtRestFindBatchSize)
        Me.Controls.Add(Me.chkMaintainRecordInfo)
        Me.Controls.Add(Me.chkInclBoneyard)
        Me.Controls.Add(Me.txtFirstName)
        Me.Controls.Add(Me.lblTestServer)
        Me.Controls.Add(Me.txtEmployeeID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Purgemain"
        Me.Text = "PurgeMain"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkSkipParsing As CheckBox
    Friend WithEvents chkExcludeRio As CheckBox
    Friend WithEvents txtRestFindBatchSize As TextBox
    Friend WithEvents chkMaintainRecordInfo As CheckBox
    Friend WithEvents chkInclBoneyard As CheckBox
    Friend WithEvents txtFirstName As TextBox
    Friend WithEvents txtEmployeeID As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents lblTestServer As Label
    Friend WithEvents PurgeStartDT As DateTimePicker
    Friend WithEvents PurgeEndDT As DateTimePicker
End Class
