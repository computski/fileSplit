
Imports System.ComponentModel
Imports System.IO
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



    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        '*** run the background process once, do not allow retriggers
        txtReport.Clear()

#If False Then
        '*** we cannot call COM from inside the BackgroundWorker (or, I cannot devote the time cracking that problem)
        '*** so instaed build a dictionary of files and their datestamps here
        Dim dicFile As New Dictionary(Of String, DateTime)
        Dim dirSource As New DirectoryInfo(txtSource.Text)
        If dirSource.Exists Then
            For Each fi As FileInfo In dirSource.GetFiles
                dicFile.Add(fi.Name, getDateFromTag(fi))
            Next
        Else
            txtReport.Text = "error: source directory not found"
            Exit Sub
        End If

        txtReport.Text = "processing files..." & vbCrLf
        Debug.WriteLine("processing files")
#End If

        '2022-02-20 correct use of COM
        Dim objShell As Object
        objShell = New Shell32.Shell


        'If (BackgroundWorker1.IsBusy = False) Then BackgroundWorker1.RunWorkerAsync(dicFile)
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

            '*** is this a zip file?
            If fi.Extension.ToUpper = ".ZIP" Then
                '*** iterate the zip
                Dim zip As New System.IO.Compression.ZipArchive(fs)
                For Each z As System.IO.Compression.ZipArchiveEntry In zip.Entries

                    Dim st As Stream = z.Open()
                    Dim content(z.Length) As Byte
                    '*** even though this is an async method, it will be blocking unless you use it in an async method 
                    '**** nope. use the non asyc method. this way we definately will read in full file then write it
                    '*** we don't need to be paralleling up lots of async file threads

                    'https://learn.microsoft.com/en-us/dotnet/api/system.io.stream.readasync?view=net-8.0

                    st.Read(content, 0, z.Length)

                    sender.ReportProgress(0, "V: " & z.Name)
                    '*** we can now write out this file to our target directory

                    Dim fart As New FileInfo("C:\Users\julia\Documents\TESTpix\thing.jpg")
                    If Not fart.Exists Then
                        '*** write to fart
                        Dim fsOut = fart.Create
                        fsOut.Write(content, 0, content.Length)
                        fsOut.Close()
                    End If


                    'st is a stream and not a filestream
                    'the file attributes are actually on z

                    'processFile(st, sender, e)
                    st.Close()
                Next

            Else
                '*** regular file
                processFile(fs, sender, e)
                sender.ReportProgress(0, "R: " & fs.Name)
            End If

            fs.Close()
        Next







        '2024-8-12 hack, to test zip files
        '  Dim fi As New FileInfo(txtSource.Text)
        '  processZipFile(fi, worker, e)


#If False Then
        Dim sourceDirectory As String = txtSource.Text
        '*** we pass a dictionary of files and their dateStamps into the worker via e.Arguement
        processFile(New DirectoryInfo(sourceDirectory), worker, e)
#End If

        worker.ReportProgress(1, vbCrLf & "+++ Complete +++")
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        txtReport.AppendText(e.UserState.ToString & vbCrLf)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If (e.Cancelled = True) Then

            txtReport.AppendText("Canceled!")

        ElseIf Not (e.Error Is Nothing) Then

            txtReport.AppendText("Error: " + e.Error.Message)

        Else
            'txtReport.Text = "Done!"
        End If
    End Sub
    ''' <summary>
    ''' Process zip files
    ''' </summary>
    Sub processZipFile(source As FileInfo, sender As Object, e As DoWorkEventArgs)
        If (source.Extension <> ".zip") Then Exit Sub

        Dim fs As IO.FileStream = New FileStream(source.FullName, FileMode.Open)

        Dim zip As New System.IO.Compression.ZipArchive(fs)
        'sb.Append("found zip file, processing contents...<br/>")
        For Each f As System.IO.Compression.ZipArchiveEntry In zip.Entries
            sender.ReportProgress(0, "V: " & f.Name)

#If False Then
            'we presumably can use streamReader as well
            'https://learn.microsoft.com/en-us/dotnet/api/system.io.compression.ziparchive.getentry?view=net-8.0
            Dim entry As System.IO.Compression.ZipArchiveEntry = zip.GetEntry(f.Name)
            Using writer As StreamWriter = New StreamWriter(entry.Open())
                writer.BaseStream.Seek(0, SeekOrigin.[End])

                writer.WriteLine("append line to file")
            End Using
#End If

            '*** read the contents.  yes this is awaitable but if we don't use it as such it becomes blocking i think





            ' Dim m As String = processPBfile(st, f.Name)
            Dim writeTarget As New FileInfo("c:\blah")
            Dim fsTarget As FileStream = writeTarget.Create()
            Dim content() As Byte
            'st.ReadAsync(content, 0, f.Length)
            'fsTarget.WriteAsync(content, 0, st.Length)  'awaitable

        Next
        fs.Close()
        'Append("zipfile processed.")

    End Sub
    ''' <summary>
    ''' Overload, process a single filestream
    ''' </summary>
    ''' <param name="fs"></param>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub processFile(fs As FileStream, sender As Object, e As DoWorkEventArgs)

    End Sub


    ''' <summary>
    ''' This routine processes individual files
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub processFile(source As DirectoryInfo, sender As Object, e As DoWorkEventArgs)
        '*** errors thrown here are captured in the BackgroundWorker1_RunWorkerCompleted routine

        If (sender.CancellationPending = True) Then e.Cancel = True : Return

        '*** we pass a dictionary of files and their dateStamps into the worker
        '*** cast e.Arguement to a Dictionary here
        'Dim dicFile As Dictionary(Of String, DateTime) = e.Argument


        Dim vidType() As String = Split(txtVidType.Text.ToUpper, ",")


        '*** note, the createTime is updated if the file is copied
        '*** the last write time is more useful as its the date the file was written and will remain correct
        '*** so long as the file is not edited

        For Each fi As FileInfo In source.GetFiles()
            If (sender.CancellationPending = True) Then e.Cancel = True : Return


            '*** based on the file suffix, we need to write it to the video or other dir
            '*** based on its date, we need to create a yyyy-mm directory to write it to


            '*** pick up the tag date the picture/video was taken
            '*** 2022-02-15 unfortunately does not work due to threading problems
            '    Dim dtTimeStamp As Date = getDateFromTag(fi)

            '*** 2022-02-15 so instead we use a dictionary of files and pass this to the worker
            '*** 2022-02-20 no need for dictionary as COM is now working
            ' Dim dtTimeStamp = dicFile.Item(fi.Name)


            '***2022-02-20 use COM to get file datestamp info
            Dim COM As Object = e.Argument
            Dim objFolder As Object = COM.NameSpace(fi.DirectoryName)  'folder

            Dim objFolderItem As Object = objFolder.ParseName(fi.Name)
            Dim dtTimeStamp As DateTime = Nothing

            '***don't need to worry about date formats because they are not stored as a culture specific format
            '***i.e. system will render them to itself in system culture and tryparse will work

            If DateTime.TryParse(Regex.Replace(objFolder.getdetailsof(objFolderItem, 208), "[^ -~]+", ""), dtTimeStamp) Then
                '208 is media created

            ElseIf DateTime.TryParse(Regex.Replace(objFolder.getdetailsof(objFolderItem, 12), "[^ -~]+", ""), dtTimeStamp) Then
                '12 is date taken
            Else
                dtTimeStamp = fi.LastWriteTime
            End If

            objFolderItem = Nothing
            objFolder = Nothing
            COM = Nothing



#If False Then
            Dim dtTimeStamp = Nothing '2022-02-15 for now, derive from the filename

            '*** Workaround is to use tryparse exact on the filename.  Samsung phone gives the file a timestamp name
            '*** yyyyMMdd_HHmmss the capitalisation is important

            If Not DateTime.TryParseExact(Path.GetFileNameWithoutExtension(fi.Name), New String() {"yyyyMMdd_HHmmss", "d-M-yyyy", "dd-MM-yyyy"}, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, dtTimeStamp) Then
                '*** fall back is to use the lastwritetime, however dumps from google will set this as Now()
                dtTimeStamp = fi.LastWriteTime
            End If
#End If

            '*** calculate what yyyy-mm directory to place this file
            Dim ymDir As String = Format(dtTimeStamp, "yyyy-MM")

            If vidType.Contains(Mid(fi.Extension, 2).ToUpper) Then
                '*** video file types
                'sender.ReportProgress(0, "V: " & fi.Name & " " & fi.CreationTime & " " & fi.LastWriteTime)
                Dim diTarget As DirectoryInfo = New DirectoryInfo(txtVidDest.Text)
                If Not diTarget.Exists Then Throw New ArgumentException("video target directory does not exist")

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
                    fTarget.CreationTime = dtTimeStamp   ''added 2024-08-12

                    sender.ReportProgress(0, "V: " & fi.Name)
                Catch
                    sender.ReportProgress(0, "skip: " & fi.Name)
                End Try

            Else
                '*** other (picture) file types
                Dim diTarget As DirectoryInfo = New DirectoryInfo(txtOtherDest.Text)
                If Not diTarget.Exists Then Throw New ArgumentException("picture target directory does not exist")

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
                    fTarget.CreationTime = dtTimeStamp  ''added 2024-08-12
                    sender.ReportProgress(0, "P: " & fi.Name)
                Catch
                    sender.ReportProgress(0, "skip: " & fi.Name)
                End Try


            End If

        Next




    End Sub

    ''' <summary>
    ''' returns MediaCreated (video files) DateTaken (picture files) or lastModified (other files)
    ''' </summary>
    ''' <param name="fi"></param>
    ''' <returns>datetime</returns>
    '''
    Function getDateFromTag(fi As FileInfo) As DateTime

        'this fails if called from within a worker process
        'https://stackoverflow.com/questions/31403956/exception-when-using-shell32-to-get-file-extended-properties

        'bummer, even adding  <STAThread()> above this function does not fix the issue
        'https://stackoverflow.com/questions/18242819/vb-net-problems-with-stathread-error-invalidoperationexception
        'https://stackoverflow.com/questions/10105518/calling-shgetfileinfo-in-thread-to-avoid-ui-freeze

        'more on this, its a complex threading issue. hmm.
        'https://stackoverflow.com/questions/54832780/how-to-access-c-com-object-in-c-sharp-background-worker
        'ergh
        'https://www.codeproject.com/Articles/990/Understanding-Classic-COM-Interoperability-With-NE



        Dim objShell As Object
        Dim objFolder As Object

        objShell = New Shell32.Shell
        objFolder = objShell.NameSpace(fi.DirectoryName)  'folder


#If False Then
        Dim detail As Object
        Dim j As Integer
        For j = 0 To 400
            detail = objFolder.GetDetailsOf(Nothing, j)
            If Not detail Is Nothing Then
                Debug.WriteLine("[" & j & "] " & detail.ToString)
            End If
        Next
#End If

        'http://www.tutorialspanel.com/different-date-conversion-string-date-vb-net/index.htm

        Dim objFolderItem As Object = objFolder.ParseName(fi.Name)
        Dim dt As DateTime = Nothing

        '***don't need to worry about date formats because they are not stored as a culture specific format
        '***i.e. system will render them to itself in system culture and tryparse will work

        If DateTime.TryParse(Regex.Replace(objFolder.getdetailsof(objFolderItem, 208), "[^ -~]+", ""), dt) Then
            '208 is media created

        ElseIf DateTime.TryParse(Regex.Replace(objFolder.getdetailsof(objFolderItem, 12), "[^ -~]+", ""), dt) Then
            '12 is date taken
        Else
            dt = fi.LastWriteTime
        End If


        objShell = Nothing
        objFolderItem = Nothing
        objFolder = Nothing
        Return dt

    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'https://www.codeproject.com/Questions/290688/Reading-Exif-Data-using-VB-Net
        'google displays by exif date-taken but when you download it sets create=modified=now
        'you need to read the DateTaken property


        'https://www.codeproject.com/Questions/815338/Inserting-GPS-tags-into-jpeg-EXIF-metadata-using-n
        'pain if you have to load it into a bitmap image first



        'https://www.codeproject.com/Articles/4956/The-ExifWorks-class

        'Dim s As String = "C:\Users\julia\Downloads\20200119_203516.jpg"
        'Dim s As String = "C:\Users\julia\Downloads\vidtest.mp4"
        Dim s As String = "C:\Users\julia\Downloads\nulist.txt"


        Dim fi As New FileInfo(s)

        MsgBox(getDateFromTag(fi))
        Exit Sub

        Debug.WriteLine(fi.DirectoryName)
        Debug.WriteLine(fi.Name)


#If False Then
        Dim EW As ExifWorks = Nothing
        Try
            EW = New ExifWorks(s)
            MsgBox(EW.DateTimeOriginal)
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If Not EW Is Nothing Then EW.Dispose()
        End Try
#End If



        'Dim fi As New FileInfo(s)


        'https://stackoverflow.com/questions/8351713/how-can-i-extract-the-date-from-the-media-created-column-of-a-video-file

        'https://stackoverflow.com/questions/62041605/updated-equivalent-code-for-using-shell-in-vb-net-with-folderitem-getdetailsof
        'Under COM import Microsoft Shell Controls And Automation 


        Dim objShell As Object
        Dim objFolder As Object

        objShell = New Shell32.Shell
        objFolder = objShell.NameSpace(fi.DirectoryName)  'folder

        'what is on offer
        Debug.WriteLine("hello")

        Dim detail As Object
        Dim j As Integer
        For j = 0 To 400
            detail = objFolder.GetDetailsOf(Nothing, j)
            If Not detail Is Nothing Then
                Debug.WriteLine("[" & j & "] " & detail.ToString)
            End If
        Next

        'http://www.tutorialspanel.com/different-date-conversion-string-date-vb-net/index.htm

        If (Not objFolder Is Nothing) Then
            Dim objFolderItem As Object

            'objFolderItem = objFolder.ParseName("vidtest.mp4") 
            objFolderItem = objFolder.ParseName(fi.Name)
            Dim dt As DateTime = Nothing

            '***long faff short, do this
            '***don't need to worry about date formats because they are not stored as a culture specific format
            '***i.e. system will render them to itself in system culture and tryparse will work

            If Not DateTime.TryParse(Regex.Replace(objFolder.getdetailsof(objFolderItem, 208), "[^ -~]+", ""), dt) Then
                '208 is media created
                If Not DateTime.TryParse(Regex.Replace(objFolder.getdetailsof(objFolderItem, 12), "[^ -~]+", ""), dt) Then
                    '12 is date taken
                    MsgBox("fallback")
                    dt = fi.LastWriteTime
                End If

            End If

            MsgBox("dt is " & dt)


            Exit Sub


            Dim x As String = objFolder.getdetailsof(objFolderItem, 208)

            '*** ok so getDetails adds in nonprintable chars which then mess up the string even if you reassign
            '*** so you need to regex it to clean it.  then you can parse it

            x = Regex.Replace(x, "[^ -~]+", "")
            Debug.WriteLine(x)
            DateTime.TryParse(x, dt)
            MsgBox(dt)
            Exit Sub





            'x = "‎5/‎07/‎2006 ‏‎3:18 AM"
            Dim z As String = "5/07/2006 3:18 AM"
            z = x
            'https://stackoverflow.com/questions/47052779/parse-date-string-with-single-digit-day-e-g-1-11-2017-as-well-as-12-11-2017
            'fiddly as 
            dt = Date.ParseExact("18-03-2016", "d-M-yyyy", Globalization.DateTimeFormatInfo.InvariantInfo)  'works
            dt = Date.ParseExact(z, "d/M/yyyy H:mm tt", Globalization.DateTimeFormatInfo.InvariantInfo)

            MsgBox(dt)
            Exit Sub
            If DateTime.TryParseExact(x, "d/M/yyyy H:m tt", Nothing, Nothing, dt) Then
                MsgBox(dt)
            Else
                MsgBox("fail " & x)
            End If
            ' dt = DateTime.ParseExact(x, "dd-MM-yyyy HH:mm tt", Nothing)

            MsgBox(objFolderItem.name & " " & objFolder.getdetailsof(objFolderItem, 208))
            MsgBox(objFolderItem.name & " " & objFolder.getdetailsof(objFolderItem, 12))
            objFolderItem = Nothing
        End If

        objFolder = Nothing
        objShell = Nothing


        'https://stackoverflow.com/questions/22382010/what-options-are-available-for-shell32-folder-getdetailsof
        '12 is date taken on win10



    End Sub
End Class
