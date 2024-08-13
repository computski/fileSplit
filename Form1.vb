
Imports System.ComponentModel
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions

'given a source directory, or a zip file (google)
'will post file to videos or photos, into a yyyy-mm folder
'if directory does not exist, will create it
'if target file exists (based on name) it will not post it but flag as a failure
'otherwise it posts a copy of the file

'the original is not modified
'source is not scanned for directories

'Google is an issue for dateStamps.  generally a file will have lastModified=exif date taken, but when you download from
'google it sets both create and modifed = now.  there is no single library that can pull the meta data for both video and photos
'so the quick fix is to work off the filename, assuming it contains a data reference.
'Better fix: we use the COM object to build a dictionary of filenames/datestamps and then pass this to the background worker

'2024-08-12 improvements.
'READ ZIP files. google photos downloads to a zip file
'MODIFY file attributes.  created and access dates are today (the download date).  Modified date is the source create date.  We want to copy this to the created date.
'ADD OPTION to use the filename datestamp (if present) in preference to info from the file itself.  e.g. 20160210_202715  alternatively we could take the earliest date
'found in the attribs and use that as the create date.






'more BGW and tread info
'https://stackoverflow.com/questions/23531371/how-to-access-com-object-within-a-backgroundworker
'https://docs.microsoft.com/en-us/dotnet/api/system.stathreadattribute?redirectedfrom=MSDN&view=net-6.0
'https://docs.microsoft.com/en-us/previous-versions/ms809971(v=msdn.10)

'solved it.  Create the com object in btnStart and then pass in as the e arguement when starting the BGW



Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'https://www.dotnetperls.com/backgroundworker-vbnet
        BackgroundWorker1.WorkerReportsProgress = True
        BackgroundWorker1.WorkerSupportsCancellation = True
        txtSource.Text = GetSetting("flspt22", "general", "sourceDir", "C:\Users\v817353\Documents\FOLDER_A\")
        txtVidType.Text = GetSetting("flspt22", "general", "vidType", ".mpg,.mp4,.avi,.thm")
        txtVidDest.Text = GetSetting("flspt22", "general", "vidDir", "C:\Users\v817353\Documents\FOLDER_VID\")
        txtOtherDest.Text = GetSetting("flspt22", "general", "otherDir", "C:\Users\v817353\Documents\FOLDER_PIX\")
        chkNestYear.Checked = GetSetting("flspt22", "general", "nestYear", False)


    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        '*** run the background process once, do not allow retriggers
        txtReport.Clear()

        '2022-02-20 correct use of COM
        Dim objShell As Object
        objShell = New Shell32.Shell
        If (BackgroundWorker1.IsBusy = False) Then BackgroundWorker1.RunWorkerAsync(objShell)

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If BackgroundWorker1.WorkerSupportsCancellation Then

            'Cancel the asynchronous operation.
            BackgroundWorker1.CancelAsync()
        End If
    End Sub

    Private Sub txtSource_TextChanged(sender As Object, e As EventArgs) Handles txtSource.TextChanged
        SaveSetting("flspt22", "general", "sourceDir", txtSource.Text)
    End Sub

    Private Sub txtVidType_TextChanged(sender As Object, e As EventArgs) Handles txtVidType.TextChanged
        SaveSetting("flspt22", "general", "vidType", txtVidType.Text)
    End Sub

    Private Sub txtVidDest_TextChanged(sender As Object, e As EventArgs) Handles txtVidDest.TextChanged
        SaveSetting("flspt22", "general", "vidDir", txtVidDest.Text)
    End Sub

    Private Sub txtOtherDest_TextChanged(sender As Object, e As EventArgs) Handles txtOtherDest.TextChanged
        SaveSetting("flspt22", "general", "otherDir", txtOtherDest.Text)
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        '*** The COM object e is passed through to processFile
        Dim worker As BackgroundWorker = sender

        '*** 2024-08-12 new approach, we iterate the directory here, including iterating zip files
        '*** we then call out to a process that handles a single file

        '*** nope, not that easy. with regular files, they exist in the source dir and you can copy them to target
        '*** zip files are in an archive. we don't really want to extract a file to the source dir and then copy it.  We'd rather 
        '*** create a file in the target dir and then copy the zip content to it

        Dim sourceDI As New DirectoryInfo(txtSource.Text)

        For Each fi As FileInfo In sourceDI.GetFiles
            Dim fs As IO.FileStream = New FileStream(fi.FullName, FileMode.Open)

            'design.  loop through source dir (at the one level).  pass files out for processing (as filestream)
            'if we find a zip file, loop through its contents passing each file out for processing (as stream, with fileinfo separate)

            'the file handling routine needs to check the properties to find the
            'create date, modified date, access date
            'and if the modified date is < create date, set the create date to this value
            'Do not rename the filenames.  Filename is not used to determine the destination.


            '*** is this a zip file?
            If fi.Extension.ToUpper = ".ZIP" Then
                '*** iterate the zip
                Dim zip As New System.IO.Compression.ZipArchive(fs)
                For Each z As System.IO.Compression.ZipArchiveEntry In zip.Entries
                    processFile(z, sender, e)
                Next

            Else
                '*** regular file
                processFile(fs, sender, e)

            End If

            fs.Close()
        Next


        worker.ReportProgress(1, vbCrLf & "+++ Complete +++")
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        txtReport.AppendText(e.UserState.ToString & vbCrLf)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If (e.Cancelled = True) Then

            txtReport.AppendText("Canceled!")

        ElseIf Not (e.Error Is Nothing) Then

            txtReport.AppendText("Fatal Error: " + e.Error.Message)

        Else
            'txtReport.Text = "Done!"
        End If
    End Sub

    ''' <summary>
    ''' Process a filestream of a single source file
    ''' </summary>
    ''' <param name="fs">Filestream of the source file</param>
    ''' <param name="sender">worker object sender</param>
    ''' <param name="e">worker object events</param>
    Sub processFile(fs As FileStream, sender As Object, e As DoWorkEventArgs)
        '*** errors thrown here are captured in the BackgroundWorker1_RunWorkerCompleted routine
        If (sender.CancellationPending = True) Then e.Cancel = True : Return
        Dim vidType() As String = Split(txtVidType.Text.ToUpper, ",")
        Dim dtTimeStamp As DateTime = Nothing
        Dim fi As FileInfo = New FileInfo(fs.Name)   'full path

        'lets assume that createdDate and LastWriteTime are the dates to work with.  Ignore the filename
        dtTimeStamp = fi.CreationTime
        If dtTimeStamp > fi.LastWriteTime Then dtTimeStamp = fi.LastWriteTime
        fixBrokenDate(dtTimeStamp, fi.Name)

        'sender.ReportProgress(0, fi.FullName & " " & dtTimeStamp.ToShortDateString)

        '*** calculate which yyyy-mm directory to place this file into
        Dim ymDir As String = Format(dtTimeStamp, "yyyy-MM")
        '*** don't try and nest to junk years
        If chkNestYear.Checked And dtTimeStamp.Year < 1980 Then Exit Sub

        If vidType.Contains(Mid(fi.Extension, 2).ToUpper) Then
            '*** video file types
            'sender.ReportProgress(0, "V: " & fi.Name & " " & fi.CreationTime & " " & fi.LastWriteTime)
            Dim diTarget As DirectoryInfo = New DirectoryInfo(txtVidDest.Text)
            If Not diTarget.Exists Then Throw New ArgumentException("video target directory does not exist")

            '*** if we are nesting, look for nested level
            If chkNestYear.Checked Then
                diTarget = New DirectoryInfo(txtVidDest.Text & "\" & dtTimeStamp.Year & "\")
                If Not diTarget.Exists Then Throw New ArgumentException("NESTED video target directory does not exist " & dtTimeStamp.Year)
            End If

            '*** if the path does not exist, create it
            Dim wDir As New DirectoryInfo(diTarget.FullName & "\" & ymDir & " vid\")

            If Not wDir.Exists Then
                wDir.Create()
                sender.ReportProgress(0, "create: " & wDir.Name)
            End If

            '*** copy file to target, do not overwrite
            Try
                fi.CopyTo(Path.Combine(wDir.FullName, fi.Name), False)
                '*** update the target with dtTimeStamp
                Dim fTarget As FileInfo = New FileInfo(Path.Combine(wDir.FullName, fi.Name))
                fTarget.LastWriteTime = dtTimeStamp
                fTarget.CreationTime = dtTimeStamp   'added 2024-08-12

                sender.ReportProgress(0, "V: " & fi.Name)
            Catch
                sender.ReportProgress(0, "skip: " & fi.Name)
            End Try

        Else
            '*** other (picture) file types
            Dim diTarget As DirectoryInfo = New DirectoryInfo(txtOtherDest.Text)
            If Not diTarget.Exists Then Throw New ArgumentException("picture target directory does not exist")

            '*** if we are nesting, look for nested level
            If chkNestYear.Checked Then
                diTarget = New DirectoryInfo(txtOtherDest.Text & "\" & dtTimeStamp.Year & "\")
                If Not diTarget.Exists Then Throw New ArgumentException("NESTED picture target directory does not exist " & dtTimeStamp.Year)
            End If


            '*** if the path does not exist, create it
            Dim wDir As New DirectoryInfo(diTarget.FullName & "\" & ymDir & " pix\")
            If Not wDir.Exists Then
                wDir.Create()
                sender.ReportProgress(0, "create: " & wDir.Name)
            End If

            '*** copy file to target, do not overwrite
            Try
                fi.CopyTo(Path.Combine(wDir.FullName, fi.Name), False)
                '*** update the target with dtTimeStamp
                Dim fTarget As FileInfo = New FileInfo(Path.Combine(wDir.FullName, fi.Name))
                fTarget.LastWriteTime = dtTimeStamp
                fTarget.CreationTime = dtTimeStamp  'added 2024-08-12
                sender.ReportProgress(0, "P: " & fi.Name)
            Catch
                sender.ReportProgress(0, "skip: " & fi.Name)
            End Try


        End If


    End Sub
    ''' <summary>
    ''' Process a single zip entry from a source zip archive
    ''' Only the LastWriteTime attribute on the file is available
    ''' </summary>
    ''' <param name="z">ZipArchiveEntry of a single source file in the source zip</param>
    ''' <param name="sender">worker object sender</param>
    ''' <param name="e">worker object events</param>
    Sub processFile(z As ZipArchiveEntry, sender As Object, e As DoWorkEventArgs)
        '*** errors thrown here are captured in the BackgroundWorker1_RunWorkerCompleted routine
        If (sender.CancellationPending = True) Then e.Cancel = True : Return
        Dim sFilename As String = z.Name

        Dim vidType() As String = Split(txtVidType.Text.ToUpper, ",")
        Dim dtTimeStamp As DateTime = Nothing
        Dim fi As FileInfo = New FileInfo(z.FullName)  'full path

        '*** with zip files, the only date that is preserved is the lastwritetime
        dtTimeStamp = z.LastWriteTime.DateTime
        fixBrokenDate(dtTimeStamp, z.Name)

        'sender.ReportProgress(0, "zipfile: " & sFilename & " " & dtTimeStamp.ToShortDateString)

        '*** calculate which yyyy-mm directory to place this file into
        Dim ymDir As String = Format(dtTimeStamp, "yyyy-MM")

        '*** don't try and nest to junk years
        If chkNestYear.Checked And dtTimeStamp.Year < 1980 Then Exit Sub


        If vidType.Contains(Mid(fi.Extension, 2).ToUpper) Then
            '*** video file types
            'sender.ReportProgress(0, "V: " & fi.Name & " " & fi.CreationTime & " " & fi.LastWriteTime)
            Dim diTarget As DirectoryInfo = New DirectoryInfo(txtVidDest.Text)
            If Not diTarget.Exists Then Throw New ArgumentException("video target directory does not exist " & dtTimeStamp.Year)

            '*** if we are nesting, look for nested level
            If chkNestYear.Checked Then
                diTarget = New DirectoryInfo(txtVidDest.Text & "\" & dtTimeStamp.Year & "\")
                If Not diTarget.Exists Then Throw New ArgumentException("NESTED video target directory does not exist " & dtTimeStamp.Year)
            End If

            '*** if the path does not exist, create it
            Dim wDir As New DirectoryInfo(diTarget.FullName & "\" & ymDir & " vid\")

            If Not wDir.Exists Then
                wDir.Create()
                sender.ReportProgress(0, "create: " & wDir.Name)
            End If

            '*** copy file to target, do not overwrite
            Try

                '*** we need to stream the bytes from the zipfile and create/write to a target
                Dim fTarget As FileInfo = New FileInfo(Path.Combine(wDir.FullName, fi.Name))
                If fTarget.Exists Then Throw New ArgumentException("exists")
                Dim st As Stream = z.Open()
                Dim content(z.Length) As Byte
                st.Read(content, 0, z.Length)
                '*** we can now write out this file to our target directory
                Dim fsWrite = fTarget.Create
                fsWrite.Write(content, 0, content.Length)
                fsWrite.Close()
                fTarget.CreationTime = dtTimeStamp
                fTarget.LastWriteTime = dtTimeStamp

                sender.ReportProgress(0, "V:zip " & fi.Name)
            Catch
                sender.ReportProgress(0, "zip skip: " & fi.Name)
            End Try

        Else
            '*** other (picture) file types
            Dim diTarget As DirectoryInfo = New DirectoryInfo(txtOtherDest.Text)
            If Not diTarget.Exists Then Throw New ArgumentException("picture target directory does not exist " & dtTimeStamp.Year)

            '*** if we are nesting, look for nested level
            If chkNestYear.Checked Then
                diTarget = New DirectoryInfo(txtOtherDest.Text & "\" & dtTimeStamp.Year & "\")
                If Not diTarget.Exists Then Throw New ArgumentException("NESTED picture target directory does not exist " & dtTimeStamp.Year)
            End If

            '*** if the path does not exist, create it
            Dim wDir As New DirectoryInfo(diTarget.FullName & "\" & ymDir & " pix\")
            If Not wDir.Exists Then
                wDir.Create()
                sender.ReportProgress(0, "create: " & wDir.Name)
            End If

            '*** copy file to target, do not overwrite
            Try

                '*** we need to stream the bytes from the zipfile and create/write to a target
                Dim fTarget As FileInfo = New FileInfo(Path.Combine(wDir.FullName, fi.Name))
                If fTarget.Exists Then Throw New ArgumentException("exists")
                Dim st As Stream = z.Open()
                Dim content(z.Length) As Byte
                st.Read(content, 0, z.Length)
                '*** we can now write out this file to our target directory
                Dim fsWrite = fTarget.Create
                fsWrite.Write(content, 0, content.Length)
                fsWrite.Close()
                fTarget.CreationTime = dtTimeStamp
                fTarget.LastWriteTime = dtTimeStamp

                sender.ReportProgress(0, "P:zip " & fi.Name)
            Catch
                sender.ReportProgress(0, "zip skip: " & fi.Name)
            End Try

        End If



    End Sub
    ''' <summary>
    ''' Fixes the broken date sometimes found in google photos downloads. If dt is less than year 2000 then try and pull a date from sFilename
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="sFilename"></param>
    ''' <returns></returns>
    Function fixBrokenDate(ByRef dt As DateTime, sFilename As String) As Boolean
        'https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings?redirectedfrom=MSDN

        '*** its more complex... sometimes the modified date is nonsense, and instead there are attributes that relate to 
        '*** dateTaken or mediacreated
        '*** but generally the files will have a regex of \d{8}_\{d6} in them which we can use, and if this is less than dt we should use it
        '*** and if dt is sub 2000 then definately we should use it
        Dim result As DateTime = #1/1/2000#

        'this almost works, but some files from the OLYMPUS are showing modify dates of 2000 and a date taken date of 2016




        '*** if we can match to a date in the filename, use this

        Dim m As Match = Regex.Match(sFilename, "(\d{8})_(\d{6})|(\d{8})")
        If m.Success Then
            '*** further search in filename for full date time, or fall back to date only

            If m.Groups(3).Length > 0 Then
                Debug.Print(m.Groups(3).ToString)
                DateTime.TryParseExact(m.Groups(3).ToString, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, result)
            Else
                DateTime.TryParseExact(m.Groups(0).ToString, "yyyyMMdd_HHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeLocal, result)
            End If

            End If

        '*** if dt is sub year 2000, definately use result
        '*** if dt is > result, then use result


        If (dt < #1/1/2000#) Then
            dt = result
            Return True
        End If

        If (dt > result) Then
            dt = result
            Return True
        End If

        'no change, exit
        Return False


#If False Then

        If (dt < #1/1/2000#) Then

            '20200104_123302
            Debug.Print(sFilename.Substring(0, 13))


            If DateTime.TryParseExact(sFilename.Substring(0, 13), "yyyyMMdd_HHmm", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeLocal, result) Then
                dt = result
                Return True
            End If

            '20200104
            If DateTime.TryParseExact(sFilename.Substring(0, 8), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, result) Then
                dt = result
                Return True
            End If
        End If
            Return False 'no change made
#End If

    End Function

    Private Sub chkNestYear_CheckedChanged(sender As Object, e As EventArgs) Handles chkNestYear.CheckedChanged
        SaveSetting("flspt22", "general", "nestYear", chkNestYear.Checked)
    End Sub
End Class
