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
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.txtSource = New System.Windows.Forms.TextBox()
        Me.txtVidType = New System.Windows.Forms.TextBox()
        Me.txtVidDest = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtOtherDest = New System.Windows.Forms.TextBox()
        Me.txtReport = New System.Windows.Forms.TextBox()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.chkNestYear = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'BackgroundWorker1
        '
        '
        'txtSource
        '
        Me.txtSource.Location = New System.Drawing.Point(16, 31)
        Me.txtSource.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtSource.Name = "txtSource"
        Me.txtSource.Size = New System.Drawing.Size(756, 22)
        Me.txtSource.TabIndex = 0
        '
        'txtVidType
        '
        Me.txtVidType.Location = New System.Drawing.Point(16, 79)
        Me.txtVidType.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtVidType.Name = "txtVidType"
        Me.txtVidType.Size = New System.Drawing.Size(756, 22)
        Me.txtVidType.TabIndex = 1
        Me.txtVidType.Text = "mpg,mp4,avi,thm,modd,moff,flv,m4v"
        '
        'txtVidDest
        '
        Me.txtVidDest.Location = New System.Drawing.Point(16, 127)
        Me.txtVidDest.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtVidDest.Name = "txtVidDest"
        Me.txtVidDest.Size = New System.Drawing.Size(756, 22)
        Me.txtVidDest.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Source directory"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 59)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(308, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Video file types (suffix, comma separated, omit dot)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 107)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(156, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Destination for video files"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 155)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(198, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Destination for pix and other files"
        '
        'txtOtherDest
        '
        Me.txtOtherDest.Location = New System.Drawing.Point(16, 175)
        Me.txtOtherDest.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtOtherDest.Name = "txtOtherDest"
        Me.txtOtherDest.Size = New System.Drawing.Size(756, 22)
        Me.txtOtherDest.TabIndex = 7
        '
        'txtReport
        '
        Me.txtReport.Location = New System.Drawing.Point(16, 235)
        Me.txtReport.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtReport.Multiline = True
        Me.txtReport.Name = "txtReport"
        Me.txtReport.Size = New System.Drawing.Size(756, 339)
        Me.txtReport.TabIndex = 8
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(827, 239)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(100, 28)
        Me.btnStart.TabIndex = 9
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(827, 274)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(100, 28)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'chkNestYear
        '
        Me.chkNestYear.AutoSize = True
        Me.chkNestYear.Location = New System.Drawing.Point(20, 207)
        Me.chkNestYear.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkNestYear.Name = "chkNestYear"
        Me.chkNestYear.Size = New System.Drawing.Size(487, 20)
        Me.chkNestYear.TabIndex = 11
        Me.chkNestYear.Text = "Nest destination under a YEAR folder in target. (This folder must already exist)"
        Me.chkNestYear.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1067, 612)
        Me.Controls.Add(Me.chkNestYear)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.txtReport)
        Me.Controls.Add(Me.txtOtherDest)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtVidDest)
        Me.Controls.Add(Me.txtVidType)
        Me.Controls.Add(Me.txtSource)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "Form1"
        Me.Text = "FileSplit 2024"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents txtSource As TextBox
    Friend WithEvents txtVidType As TextBox
    Friend WithEvents txtVidDest As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents txtOtherDest As TextBox
    Friend WithEvents txtReport As TextBox
    Friend WithEvents btnStart As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents chkNestYear As CheckBox
End Class
