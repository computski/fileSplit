<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BackgroundWorker1
        '
        '
        'txtSource
        '
        Me.txtSource.Location = New System.Drawing.Point(12, 25)
        Me.txtSource.Name = "txtSource"
        Me.txtSource.Size = New System.Drawing.Size(568, 20)
        Me.txtSource.TabIndex = 0
        '
        'txtVidType
        '
        Me.txtVidType.Location = New System.Drawing.Point(12, 64)
        Me.txtVidType.Name = "txtVidType"
        Me.txtVidType.Size = New System.Drawing.Size(568, 20)
        Me.txtVidType.TabIndex = 1
        Me.txtVidType.Text = "mpg,mp4,avi,thm,modd,moff,flv,m4v"
        '
        'txtVidDest
        '
        Me.txtVidDest.Location = New System.Drawing.Point(12, 103)
        Me.txtVidDest.Name = "txtVidDest"
        Me.txtVidDest.Size = New System.Drawing.Size(568, 20)
        Me.txtVidDest.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Source directory"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(244, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Video file types (suffix, comma separated, omit dot)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 87)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(125, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Destination for video files"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 126)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(160, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Destination for pix and other files"
        '
        'txtOtherDest
        '
        Me.txtOtherDest.Location = New System.Drawing.Point(12, 142)
        Me.txtOtherDest.Name = "txtOtherDest"
        Me.txtOtherDest.Size = New System.Drawing.Size(568, 20)
        Me.txtOtherDest.TabIndex = 7
        '
        'txtReport
        '
        Me.txtReport.Location = New System.Drawing.Point(12, 191)
        Me.txtReport.Multiline = True
        Me.txtReport.Name = "txtReport"
        Me.txtReport.Size = New System.Drawing.Size(568, 247)
        Me.txtReport.TabIndex = 8
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(621, 191)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 9
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(621, 220)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(621, 250)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "exif"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Button1)
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
        Me.Name = "Form1"
        Me.Text = "FileSplit 2022"
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
    Friend WithEvents Button1 As Button
End Class
