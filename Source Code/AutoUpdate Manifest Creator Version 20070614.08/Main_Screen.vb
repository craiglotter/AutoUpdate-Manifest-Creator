Imports System.ComponentModel
Imports System.Threading
Imports System.IO

Public Class Main_Screen

    Dim precount_max As Integer
    Dim cancel_operation As Boolean
    Dim currentcount As Integer
    Dim percentComplete As Integer
    Dim fileerrorlist As String = ""
    Dim filerenamelist As String = ""

    Dim apptorun As String = """" & Application.StartupPath & "\7za.exe"" a "
    Dim filelist As ArrayList

    Dim commandlineExit As Boolean = False
    Dim commandlineApplicationName As String = ""

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\")) = True Then
                My.Computer.Audio.Play((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\"), AudioPlayMode.Background)
            End If
            Dim Display_Message1 As New Display_Message()
            Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.Message.ToString
            Display_Message1.Timer1.Interval = 1000
            Display_Message1.ShowDialog()
            If My.Computer.FileSystem.DirectoryExists((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs") = False Then
                My.Computer.FileSystem.CreateDirectory((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            End If
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub startAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startAsyncButton.Click
        Try
            If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                SaveFileDialog1.InitialDirectory = FolderBrowserDialog1.SelectedPath
                If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    cancelAsyncButton.Enabled = True
                    startAsyncButton.Enabled = False
                    Label2.Text = FolderBrowserDialog1.SelectedPath
                    ToolTip1.SetToolTip(Label2, FolderBrowserDialog1.SelectedPath)
                    Label4.Text = "Running Precount Function"
                    filelist.Clear()
                    ProgressBar1.Value = 0
                    fileerrorlist = ""
                    filerenamelist = ""
                    BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "startAsyncButton_Click")
        End Try
    End Sub


    Private Sub cancelAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelAsyncButton.Click
        Try
            cancel_operation = True
            Me.BackgroundWorker1.CancelAsync()
            cancelAsyncButton.Enabled = False
            startAsyncButton.Enabled = True
        Catch ex As Exception
            Error_Handler(ex, "cancelAsyncButton_Click")
        End Try
    End Sub

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
            precount_max = 0
            currentcount = 0
            percentComplete = 0
            cancel_operation = False
            If worker.CancellationPending Then
                e.Cancel = True
                cancel_operation = True
            End If
            Dim dinfo As DirectoryInfo = New DirectoryInfo(FolderBrowserDialog1.SelectedPath)
            Precount(dinfo)
            dinfo = Nothing

            If precount_max > 0 Then
                percentComplete = CSng(currentcount) / CSng(precount_max) * 100
            Else
                percentComplete = 100
            End If
            worker.ReportProgress(percentComplete)
            If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) = True Then
                My.Computer.FileSystem.DeleteFile(SaveFileDialog1.FileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            End If
            Dim dinfo2 As DirectoryInfo = New DirectoryInfo(FolderBrowserDialog1.SelectedPath)
            Dim filewriter As StreamWriter = New StreamWriter(SaveFileDialog1.FileName, True, System.Text.Encoding.UTF8)
            filelist.Add(SaveFileDialog1.FileName)
            filewriter.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
            If commandlineApplicationName.Length > 0 Then
                filewriter.WriteLine("<application name=""" & commandlineApplicationName & """>")
            Else
                filewriter.WriteLine("<application name=""" & dinfo2.Name & """>")
            End If

            filewriter.WriteLine(vbTab & "<manifestUpdated>" & Format(Now, "yyyyMMddHHmm") & "</manifestUpdated>")
            filewriter.Close()
            filewriter.Dispose()
            filewriter = Nothing
            FileRunner(worker, e, dinfo2)

            filewriter = New StreamWriter(SaveFileDialog1.FileName, True, System.Text.Encoding.UTF8)
            filewriter.WriteLine("</application>")
            filewriter.Close()
            filewriter.Dispose()
            filewriter = Nothing
            Dim fin As FileInfo = New FileInfo(SaveFileDialog1.FileName)
            If commandlineApplicationName.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists((fin.DirectoryName & "\" & commandlineApplicationName & " Manifest").Replace("\\", "\")) = False Then
                    My.Computer.FileSystem.CreateDirectory((fin.DirectoryName & "\" & commandlineApplicationName & " Manifest").Replace("\\", "\"))
                End If
                For Each str As String In filelist
                    Dim f As FileInfo = New FileInfo(str)
                    My.Computer.FileSystem.MoveFile(f.FullName, (fin.DirectoryName & "\" & commandlineApplicationName & " Manifest\" & f.Name).Replace("\\", "\"), True)
                    f = Nothing
                Next
            Else
                If My.Computer.FileSystem.DirectoryExists((fin.DirectoryName & "\" & dinfo2.Name & " Manifest").Replace("\\", "\")) = False Then
                    My.Computer.FileSystem.CreateDirectory((fin.DirectoryName & "\" & dinfo2.Name & " Manifest").Replace("\\", "\"))
                End If
                For Each str As String In filelist
                    Dim f As FileInfo = New FileInfo(str)
                    My.Computer.FileSystem.MoveFile(f.FullName, (fin.DirectoryName & "\" & dinfo2.Name & " Manifest\" & f.Name).Replace("\\", "\"), True)
                    f = Nothing
                Next
            End If

            dinfo2 = Nothing
            fin = Nothing
            e.Result = ""
        Catch ex As Exception
            Error_Handler(ex, "backgroundWorker1_DoWork")
        End Try
    End Sub

    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Not (e.Error Is Nothing) Then
                Error_Handler(e.Error, "backgroundWorker1_RunWorkerCompleted")
            ElseIf e.Cancelled Then
                Me.ProgressBar1.Value = 0
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\HEEY.WAV").Replace("\\", "\")) = True Then
                    My.Computer.Audio.Play((Application.StartupPath & "\Sounds\HEEY.WAV").Replace("\\", "\"), AudioPlayMode.Background)
                End If
            Else
                Me.ProgressBar1.Value = 100
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\VICTORY.WAV").Replace("\\", "\")) = True Then
                    My.Computer.Audio.Play((Application.StartupPath & "\Sounds\VICTORY.WAV").Replace("\\", "\"), AudioPlayMode.Background)
                End If
                If filerenamelist.Length > 0 Then
                    Dim disp As Display_Message = New Display_Message()
                    disp.Timer1.Interval = 1000
                    disp.Message_Textbox.Text = "The following files were first renamed and then included in the generated manifest as they contain the illegal characters '&' or '#'" & vbCrLf & vbCrLf & filerenamelist & vbCrLf & "It is suggested you check that this does not impact too heavily on the original application."
                    disp.ShowDialog()
                End If
                If fileerrorlist.Length > 0 Then
                    Dim disp As Display_Message = New Display_Message()
                    disp.Timer1.Interval = 1000
                    disp.Message_Textbox.Text = "The following files were not included in the generated manifest as they contain the illegal characters '&' or '#'" & vbCrLf & vbCrLf & fileerrorlist & vbCrLf & "It is suggested you rename these files in the original application before recreating the manifest"
                    disp.ShowDialog()
                End If
                If My.Computer.FileSystem.FileExists("C:\WINDOWS\explorer.exe") Then
                    Dim ff As FileInfo = New FileInfo(SaveFileDialog1.FileName)
                    Dim dd As String = ff.Directory.FullName
                    ff = Nothing
                    Process.Start("C:\WINDOWS\explorer.exe", dd)
                End If
            End If
            cancelAsyncButton.Enabled = False
            startAsyncButton.Enabled = True
            If commandlineExit = True Then
                Me.Close()
            End If
        Catch ex As Exception
            Error_Handler(ex, "backgroundWorker1_RunWorkerCompleted")
            If commandlineExit = True Then
                Me.Close()
            End If
        End Try
    End Sub

    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            Me.Label4.Text = currentcount & "/" & precount_max
            If e.ProgressPercentage < 100 Then
                Me.ProgressBar1.Value = e.ProgressPercentage
            Else
                Me.ProgressBar1.Value = 100
            End If

        Catch ex As Exception
            Error_Handler(ex, "backgroundWorker1_ProgressChanged")
        End Try
    End Sub

    Private Sub Precount(ByVal currentdir As DirectoryInfo)
        Try
            For Each finfo As FileInfo In currentdir.GetFiles
                precount_max = precount_max + 1
                If cancel_operation = True Then
                    finfo = Nothing
                    currentdir = Nothing
                    Exit Sub
                End If
                finfo = Nothing
            Next
            For Each dinfo As DirectoryInfo In currentdir.GetDirectories()
                If cancel_operation = True Then
                    dinfo = Nothing
                    currentdir = Nothing
                    Exit Sub
                End If
                Precount(dinfo)
                dinfo = Nothing
            Next
            currentdir = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Precount")
        End Try
    End Sub

    Private Sub FileRunner(ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs, ByVal currentdir As DirectoryInfo)
        Try
            Dim nofileerror As Boolean = True
            For Each finfo As FileInfo In currentdir.GetFiles
                nofileerror = True
                'If finfo.Name.IndexOf("&") <> -1 Then
                '    If My.Computer.FileSystem.FileExists((finfo.FullName & "\" & finfo.Name.Replace("&", "_26")).Replace("\\", "\")) = False Then
                '        finfo.MoveTo((finfo.FullName & "\" & finfo.Name.Replace("&", "_26")).Replace("\\", "\"))
                '        filerenamelist = filerenamelist & " - " & finfo.Name.Replace("_26", "&") & " -> " & finfo.Name & vbCrLf
                '    End If
                'End If
                'If finfo.Name.IndexOf("#") <> -1 Then
                '    If My.Computer.FileSystem.FileExists((finfo.FullName & "\" & finfo.Name.Replace("#", "_23")).Replace("\\", "\")) = False Then
                '        finfo.MoveTo((finfo.FullName & "\" & finfo.Name.Replace("#", "_23")).Replace("\\", "\"))
                '        filerenamelist = filerenamelist & " - " & finfo.Name.Replace("_23", "#") & " -> " & finfo.Name & vbCrLf
                '    End If
                'End If
                'If finfo.Name.IndexOf("&") <> -1 Then
                '    fileerrorlist = fileerrorlist & " - " & finfo.Name & vbCrLf
                '    nofileerror = False
                'End If
                'If finfo.Name.IndexOf("#") <> -1 Then
                '    fileerrorlist = fileerrorlist & " - " & finfo.Name & vbCrLf
                '    nofileerror = False
                'End If
                If nofileerror = True Then


                    If filelist.IndexOf(finfo.FullName) = -1 Then
                        Dim newname As String = (finfo.DirectoryName & "\" & Format(Now(), "ddMMyyyyHHmmssf") & finfo.Name.Replace("&", "").Replace("#", "")).Replace("\\", "\") & ".zip"
                        Dim newnameShort As String = (Format(Now(), "ddMMyyyyHHmmssf") & finfo.Name.Replace("&", "").Replace("#", "")).Replace("\\", "\") & ".zip"
                        Dim namecounter As Integer = 1
                        'Dim tempname, tempnameShort As String
                        'tempname = newname
                        'tempnameShort = newnameShort
                        'While My.Computer.FileSystem.FileExists(tempname) = True
                        '    tempname = newname
                        '    tempnameShort = newnameShort
                        '    tempname = tempname & namecounter
                        '    tempnameShort = tempnameShort & namecounter
                        '    namecounter = namecounter + 1
                        'End While
                        'namecounter = Nothing
                        'newname = tempname
                        'newnameShort = tempnameShort
                        'If newnameShort.ToLower.StartsWith("default") Then
                        '    MsgBox(tempname)
                        'End If
                        filelist.Add(newname)
                        If My.Computer.FileSystem.FileExists(newname) = True Then
                            My.Computer.FileSystem.DeleteFile(newname, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        End If
                        DosShellCommand(apptorun & """" & newname & """ """ & finfo.FullName & """")

                        Dim filewriter As StreamWriter = New StreamWriter(SaveFileDialog1.FileName, True, System.Text.Encoding.UTF8)
                        filewriter.WriteLine(vbTab & "<file>")
                        filewriter.WriteLine(vbTab & vbTab & "<filename>" & newnameShort & "</filename>")
                        filewriter.WriteLine(vbTab & vbTab & "<filepath>" & (finfo.DirectoryName & "\" & newnameShort).Remove(0, FolderBrowserDialog1.SelectedPath.Length) & "</filepath>")
                        filewriter.WriteLine(vbTab & vbTab & "<filepathclear>" & (finfo.DirectoryName & "\" & finfo.Name).Remove(0, FolderBrowserDialog1.SelectedPath.Length) & "</filepathclear>")
                        filewriter.WriteLine(vbTab & vbTab & "<filesize>" & finfo.Length & "</filesize>")
                        filewriter.WriteLine(vbTab & vbTab & "<filelastmodified>" & Format(finfo.LastWriteTime, "yyyyMMddHHmm") & "</filelastmodified>")
                        filewriter.WriteLine(vbTab & "</file>")
                        filewriter.Close()
                        filewriter.Dispose()
                        filewriter = Nothing
                        currentcount = currentcount + 1
                        If precount_max > 0 Then
                            percentComplete = CSng(currentcount) / CSng(precount_max) * 100
                        Else
                            percentComplete = 100
                        End If
                        worker.ReportProgress(percentComplete)
                    End If
                    If cancel_operation = True Then
                        finfo = Nothing
                        currentdir = Nothing
                        Exit Sub
                    End If
                    finfo = Nothing
                End If
            Next
            For Each dinfo As DirectoryInfo In currentdir.GetDirectories()
                If cancel_operation = True Then
                    dinfo = Nothing
                    currentdir = Nothing
                    Exit Sub
                End If
                FileRunner(worker, e, dinfo)
                dinfo = Nothing
            Next
            currentdir = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Precount")
        End Try
    End Sub

    Private Sub Main_Screen_Close(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            If FolderBrowserDialog1.SelectedPath.Length > 0 Then
                Dim dinfo As DirectoryInfo = New DirectoryInfo(FolderBrowserDialog1.SelectedPath)
                If dinfo.Exists Then
                    My.Settings.lastfolder = FolderBrowserDialog1.SelectedPath
                    My.Settings.Save()
                End If
                dinfo = Nothing
            End If

        Catch ex As Exception
            Error_Handler(ex, "Application Close")
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            'MsgBox(Format(Now(), "ddMMyyyyHHmmssf"))
            filelist = New ArrayList
            Label3.Text = "BUILD " & My.Application.Info.Version.Major & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & My.Application.Info.Version.Revision
            If Not My.Settings.lastfolder = Nothing Then
                If My.Settings.lastfolder.Length > 0 Then
                    Dim dinfo As DirectoryInfo = New DirectoryInfo(My.Settings.lastfolder)
                    If dinfo.Exists Then
                        FolderBrowserDialog1.SelectedPath = My.Settings.lastfolder
                        SaveFileDialog1.InitialDirectory = My.Settings.lastfolder
                    End If
                    dinfo = Nothing
                End If
            End If


            If My.Application.CommandLineArgs.Count = 3 Then
                Try
                    If My.Application.CommandLineArgs(0).Length > 0 And My.Application.CommandLineArgs(1).Length > 0 And My.Application.CommandLineArgs(2).Length > 0 Then
                        If My.Computer.FileSystem.DirectoryExists(My.Application.CommandLineArgs(0)) Then
                            If My.Computer.FileSystem.FileExists(My.Application.CommandLineArgs(1)) = False Then
                                cancelAsyncButton.Enabled = True
                                startAsyncButton.Enabled = False
                                Label2.Text = My.Application.CommandLineArgs(0)
                                ToolTip1.SetToolTip(Label2, My.Application.CommandLineArgs(0))
                                FolderBrowserDialog1.SelectedPath = My.Application.CommandLineArgs(0)
                                SaveFileDialog1.FileName = My.Application.CommandLineArgs(1)
                                commandlineExit = True
                                commandlineApplicationName = My.Application.CommandLineArgs(2)
                                Label4.Text = "Running Precount Function"
                                filelist.Clear()
                                ProgressBar1.Value = 0
                                fileerrorlist = ""
                                filerenamelist = ""
                                BackgroundWorker1.RunWorkerAsync()
                            End If
                        End If
                    End If

                Catch ex1 As Exception
                    Error_Handler(ex1, "Load from Command Line")
                    Application.Exit()
                End Try
            End If
        Catch ex As Exception
            Error_Handler(ex, "Load")
        End Try
    End Sub


    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False

            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If



            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()


        Catch ex As Exception
            Error_Handler(ex)
        End Try
        Return s
    End Function

End Class