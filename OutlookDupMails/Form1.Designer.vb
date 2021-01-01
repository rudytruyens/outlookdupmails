<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblstatus = New System.Windows.Forms.Label()
        Me.chkSameSize = New System.Windows.Forms.CheckBox()
        Me.chkSameSentDate = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblsummary = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.lblsubfolders = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(15, 368)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 30)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Start"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lblstatus
        '
        Me.lblstatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblstatus.Location = New System.Drawing.Point(12, 401)
        Me.lblstatus.Name = "lblstatus"
        Me.lblstatus.Size = New System.Drawing.Size(511, 28)
        Me.lblstatus.TabIndex = 1
        Me.lblstatus.Text = "Click start to search for double emails"
        '
        'chkSameSize
        '
        Me.chkSameSize.AutoSize = True
        Me.chkSameSize.Location = New System.Drawing.Point(12, 57)
        Me.chkSameSize.Name = "chkSameSize"
        Me.chkSameSize.Size = New System.Drawing.Size(178, 21)
        Me.chkSameSize.TabIndex = 2
        Me.chkSameSize.Text = "the same message size"
        Me.chkSameSize.UseVisualStyleBackColor = True
        '
        'chkSameSentDate
        '
        Me.chkSameSentDate.AutoSize = True
        Me.chkSameSentDate.Checked = True
        Me.chkSameSentDate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSameSentDate.Location = New System.Drawing.Point(12, 84)
        Me.chkSameSentDate.Name = "chkSameSentDate"
        Me.chkSameSentDate.Size = New System.Drawing.Size(151, 21)
        Me.chkSameSentDate.TabIndex = 3
        Me.chkSameSentDate.Text = "the same sent date"
        Me.chkSameSentDate.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(752, 34)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Click start to search for duplicate emails (with same mail body content)"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(165, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "and having also:"
        '
        'lblsummary
        '
        Me.lblsummary.Location = New System.Drawing.Point(12, 120)
        Me.lblsummary.Name = "lblsummary"
        Me.lblsummary.Size = New System.Drawing.Size(587, 44)
        Me.lblsummary.TabIndex = 6
        Me.lblsummary.UseCompatibleTextRendering = True
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Location = New System.Drawing.Point(15, 225)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.ListBox1.Size = New System.Drawing.Size(584, 132)
        Me.ListBox1.TabIndex = 7
        '
        'lblsubfolders
        '
        Me.lblsubfolders.Location = New System.Drawing.Point(15, 178)
        Me.lblsubfolders.Name = "lblsubfolders"
        Me.lblsubfolders.Size = New System.Drawing.Size(587, 44)
        Me.lblsubfolders.TabIndex = 8
        Me.lblsubfolders.UseCompatibleTextRendering = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(630, 438)
        Me.Controls.Add(Me.lblsubfolders)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.lblsummary)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkSameSentDate)
        Me.Controls.Add(Me.chkSameSize)
        Me.Controls.Add(Me.lblstatus)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents lblstatus As Label
    Friend WithEvents chkSameSize As CheckBox
    Friend WithEvents chkSameSentDate As CheckBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents lblsummary As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents lblsubfolders As Label
End Class
