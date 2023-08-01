Imports System.IO
Public Class Form1
    Dim objBaseClass As ClsBase
    Dim objFileValidate As ClsValidation
    Dim objGetSetINI As ClsShared
    Dim StrEncrpt As String = String.Empty

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try

            Timer1.Interval = 1000
            Timer1.Enabled = False

            Conversion_Process()

            Timer1.Enabled = True

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form Load", "Timer1_Tick")
        End Try
    End Sub
    Private Sub Generate_SettingFile()

        Dim strConverterCaption As String = ""
        Dim strSettingsFilePath As String = My.Application.Info.DirectoryPath & "\settings.ini"

        Try
            objGetSetINI = New ClsShared

            '-Genereate Settings.ini File-
            If Not File.Exists(strSettingsFilePath) Then

                '-General Section-
                Call objGetSetINI.SetINISettings("General", "Date", Format(Now, "dd/MM/yyyy"), strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Audit Log", My.Application.Info.DirectoryPath & "\Audit", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Error Log", My.Application.Info.DirectoryPath & "\Error", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Input Folder", My.Application.Info.DirectoryPath & "\INPUT", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Epay Output Folder", My.Application.Info.DirectoryPath & "\Output", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Report Folder", My.Application.Info.DirectoryPath & "\Report", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "Response Folder", My.Application.Info.DirectoryPath & "\Response", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "RevResponse Folder", My.Application.Info.DirectoryPath & "\RevResponse", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "Validation", My.Application.Info.DirectoryPath & "\Validation\CHIP_FUND_Validation.xls", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderSuc", My.Application.Info.DirectoryPath & "\Archive\Success", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderUnSuc", My.Application.Info.DirectoryPath & "\Archive\UnSuccess", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Temp Folder", My.Application.Info.DirectoryPath & "\Temp", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "==", "==========================================", strSettingsFilePath) 'Separator

                Call objGetSetINI.SetINISettings("General", "Res_Archived FolderSuc", My.Application.Info.DirectoryPath & "\Archive\Res_Archive\Success", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Res_Archived FolderUnSuc", My.Application.Info.DirectoryPath & "\Archive\Res_Archive\UnSuccess", strSettingsFilePath)

                'Call objGetSetINI.SetINISettings("General", "Temp Folder", My.Application.Info.DirectoryPath & "\Temp", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "Converter Caption", "CHIP FUND Converter", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Process Output File Ignoring Invalid Transactions", "N", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "File Counter", "0", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Number Of Records In Per Output File", "1500", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "==", "==========================================", strSettingsFilePath) 'Separator


                '-Client Details Section-
                Call objGetSetINI.SetINISettings("Client Details", "Client Name", "CHIP FUND Converter", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Client Code", "CHIP FUND Converter", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Input Date Format", "dd/MM/yyyy", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "==", "====================================", strSettingsFilePath) 'Separator

                'Epay details

                Call objGetSetINI.SetINISettings("Payment Details", "Debit Account Number", "", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Payment Details", "Remitter Name", "", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Payment Details", "=================================================", "==", strSettingsFilePath) 'Separator


                Call objGetSetINI.SetINISettings("FileProcessing", "Timer Interval", "2000", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("Directory Path(Encryption_Decryption)", "Encryption\Decryption_Dir", "C:", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Directory Path(Encryption_Decryption)", "Encryption\Decryption_Folder", "", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Directory Path(Encryption_Decryption)", "==", "====================================", strSettingsFilePath) 'Separator

                '-Encryption Section-
                Call objGetSetINI.SetINISettings("Encryption", "Encryption required for Epay", "N", strSettingsFilePath)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Call objGetSetINI.SetINISettings("Encryption", "Batch File Path", "C:\encrypt\Test.bat", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Encryption", "PICKDIR Path", "C:\encrypt\IN", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Encryption", "DROPDIR Path", "C:\encrypt\OUT", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Encryption", "CRCDIR Path", "C:\encrypt\CRC", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Encryption", "==", "==", strSettingsFilePath) '-Separator-

            End If

            '-Get Converter Caption from Settings-
            If File.Exists(strSettingsFilePath) Then
                strConverterCaption = objGetSetINI.GetINISettings("General", "Converter Caption", strSettingsFilePath)
                If strConverterCaption <> "" Then
                    Text = strConverterCaption.ToString() & " - Version " & Mid(Application.ProductVersion.ToString(), 1, 3)
                Else
                    MsgBox("Either settings.ini file does not contains the key as [ Converter Caption ] or the key value is blank" & vbCrLf & "Please refer to " & strSettingsFilePath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End
                End If
            End If

        Catch ex As Exception
            MsgBox("Error" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error while Generating Settings File")
            End

        Finally
            If Not objGetSetINI Is Nothing Then
                objGetSetINI.Dispose()
                objGetSetINI = Nothing
            End If

        End Try

    End Sub
    Private Sub Conversion_Process()
        Dim objfolderAll As DirectoryInfo

        Try
            If objBaseClass Is Nothing Then
                objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            End If

            '-Get Settings-
            If GetAllSettings() = True Then
                MsgBox("Either file path is invalid or any key value is left blank in settings.ini file", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error In Settings")
                Exit Sub
            End If


            '-Process Input-
            objfolderAll = New DirectoryInfo(strInputFolderPath)
            If objfolderAll.GetFiles.Length = 0 Then
                objfolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started for INPUT Files")

                For Each file As FileInfo In objfolderAll.GetFiles("*")
                    objBaseClass.isCompleteFileAvailable(file.FullName)
                    If Mid(file.FullName, file.FullName.Length - 3, 4).ToString().ToUpper() = ".CSV" Then
                        objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("INPUT File [ " & file.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)
                        Process_Each(file.FullName)

                        objfolderAll.Refresh()
                    Else
                        objBaseClass.LogEntry("Invalid File Format", False)
                    End If
                Next
            End If

            ' For Response 

            objfolderAll = Nothing

            objfolderAll = New DirectoryInfo(strResponseFolderPath)

            If objfolderAll.GetFiles.Length = 0 Then
                objfolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started for RESPONSE Files")

                For Each objFileOne As FileInfo In objfolderAll.GetFiles()
                    objBaseClass.isCompleteFileAvailable(objFileOne.FullName)
                    If Mid(objFileOne.FullName, objFileOne.FullName.Length - 3, 4).ToString().ToUpper() = ".txt".ToUpper And Mid(objFileOne.FullName, objFileOne.FullName.Length - 3, 4).ToString().ToUpper() <> ".BAK" Then
                        objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("RESPONSE File [ " & objFileOne.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)

                        Response_File(objFileOne.FullName)

                        objfolderAll.Refresh()

                    End If
                Next
            End If


        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Conversion_Process")

        Finally
            If Not objBaseClass Is Nothing Then
                objBaseClass.Dispose()
                objBaseClass = Nothing
            End If
        End Try
    End Sub
    Private Sub Response_File(ByVal strResFileName As String)

        Dim strResponseInputFile As String
        Try
            gstrResponseInputFolder = strResFileName.Substring(0, strResFileName.LastIndexOf("\"))
            gstrResponseInputFile = strResFileName.Substring(strResFileName.LastIndexOf("\"))
            strResponseInputFile = gstrResponseInputFile.Replace("\", "")

            Dim strRespFile As String = ""

            If (strResponseInputFile.ToUpper).Contains("_EOD") Then
                'Else
                '    objBaseClass.LogEntry("Invalid Response File Format", False)
                'End If

                objFileValidate = New ClsValidation(strResFileName, objBaseClass.gstrIniPath)

                If objFileValidate.CheckResponseValidateFile(strResFileName) = True Then
                    objBaseClass.LogEntry("Response File Reading Completed Successfully")

                    If (objFileValidate.DtUnSucResp.Rows.Count = 0) Or (strProceed.ToString().Trim().ToUpper() = "Y") Then

                        If objFileValidate.DtInputResp.Rows.Count > 0 Then
                            objBaseClass.LogEntry("Reverse File Generation Process Started....")

                            If Generate_Output_Response(objFileValidate.DtInputResp, strResponseInputFile) = False Then
                                objBaseClass.FileMove(strResFileName, strResArchivedFolderUnSuc & "\" & Path.GetFileName(strResFileName))
                                objBaseClass.LogEntry("Reverse File Generation process failed due to Error", True)
                                objBaseClass.LogEntry("Reverse input file :" + Path.GetFileName(strResFileName) + " Is Moved to " + strResArchivedFolderUnSuc)
                            Else
                                objBaseClass.FileMove(strResFileName, strResArchivedFolderSuc & "\" & Path.GetFileName(strResFileName))
                                objBaseClass.LogEntry("Reverse Files are Generated Successfully", False)
                                objBaseClass.LogEntry("Reverse input file :" + Path.GetFileName(strResFileName) + " Is Moved to " + strResArchivedFolderSuc)
                            End If
                        Else
                            objBaseClass.LogEntry("No Valid Record present in Response File")
                            objBaseClass.FileMove(strResFileName, strResArchivedFolderUnSuc & "\" & Path.GetFileName(strResFileName))
                            objBaseClass.LogEntry("Reverse input file :" + Path.GetFileName(strResFileName) + " Is Moved to " + strResArchivedFolderUnSuc)
                        End If
                    Else

                        objBaseClass.LogEntry("No Valid Record present in Response File")
                        objBaseClass.FileMove(strResFileName, strResArchivedFolderUnSuc & "\" & Path.GetFileName(strResFileName))
                        objBaseClass.LogEntry("Reverse input file :" + Path.GetFileName(strResFileName) + " Is Moved to " + strResArchivedFolderUnSuc)
                    End If

                    If objFileValidate.DtUnSucResp.Rows.Count > 0 Then
                        objBaseClass.LogEntry("Response File contains following Discrepancies")
                        objBaseClass.LogEntry("Writing Instruction failed for  Response File following ")

                        With objFileValidate.DtUnSucResp
                            For Each _dtRow As DataRow In .Rows
                                If _dtRow("Reason").ToString().Trim() <> "" Then
                                End If
                                objBaseClass.LogEntry(_dtRow("Reason").ToString)
                            Next
                        End With

                    End If
                Else
                    objBaseClass.LogEntry("Invalid Response File")
                    objBaseClass.FileMove(strResFileName, strResArchivedFolderUnSuc & "\" & Path.GetFileName(strResFileName))
                    objBaseClass.LogEntry("Reverse input file :" + Path.GetFileName(strResFileName) + " Is Moved to " + strResArchivedFolderUnSuc)
                End If
            End If
        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "CHIP_FUND_CONVERTER", "Response_File")
        End Try
    End Sub
    Private Sub Process_Each(ByVal strInputFileName As String)
        Dim TrnProcSuc As Boolean
        Try

            gstrInputFolder = strInputFileName.Substring(0, strInputFileName.LastIndexOf("\"))
            gstrInputFile = strInputFileName.Substring(strInputFileName.LastIndexOf("\"))
            gstrInputFile = gstrInputFile.Replace("\", "")

            '-Conversion Process-

            objBaseClass.LogEntry("", False)
            objBaseClass.LogEntry("Process Started")
            objBaseClass.LogEntry("Reading Input File " & gstrInputFile, False)

            objFileValidate = New ClsValidation(strInputFileName, objBaseClass.gstrIniPath)

            If objFileValidate.CheckValidateFile(strInputFileName) = True Then

                objBaseClass.LogEntry("Input File Reading Completed Successfully", False)

                If (objFileValidate.DtUnSucEpayOutput.Rows.Count = 0) Or (strProceed.ToString().Trim().ToUpper() = "Y") Then
                    objBaseClass.LogEntry("Input File Validated Successfully", False)

                    If objFileValidate.DtSuccEpayOutput.Rows.Count > 0 Then

                        objBaseClass.LogEntry("Output File Generation Process Started", False)

                        If GenerateOutPutFile(objFileValidate.DtSuccEpayOutput, gstrInputFile) = False Then       ''Generating Output
                            TrnProcSuc = False
                            objBaseClass.LogEntry("Output File Generation process failed due to Error", True)
                            objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                            objBaseClass.LogEntry("Input file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)
                        Else
                            TrnProcSuc = True
                            objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderSuc & "\" & gstrInputFile)
                            objBaseClass.LogEntry("Input file [" + Path.GetFileName(strInputFileName) + "] Is Moved to " + strArchivedFolderSuc)
                            objBaseClass.LogEntry("Output Files is Generated Successfully", False)
                        End If

                    Else
                        TrnProcSuc = False
                        objBaseClass.LogEntry("No Valid Record present in Input File")
                        objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                        objBaseClass.LogEntry("Input file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)
                    End If
                Else
                    TrnProcSuc = False
                    objBaseClass.LogEntry("No Valid Record present in Input File")
                    objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                    objBaseClass.LogEntry("Input file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)
                End If

                '-Write Summary Report-
                Dim strSummaryFileName As String
                strSummaryFileName = Path.GetFileNameWithoutExtension(gstrInputFile)
                objBaseClass.LogEntry("[Writing Transaction Report]")
                Call Payment_Report()
                objBaseClass.LogEntry("Transaction Report File Generated Successfully")

                If objFileValidate.DtUnSucEpayOutput.Rows.Count > 0 Then
                    objBaseClass.LogEntry("Epay File contains following Discrepancies")
                    objBaseClass.LogEntry("Writing Instruction failed for Epay file ")

                    With objFileValidate.DtUnSucEpayOutput
                        For Each _dtRow As DataRow In .Rows
                            If _dtRow("Reason").ToString().Trim() <> "" Then
                                objBaseClass.LogEntry(_dtRow("Reason").ToString)
                            End If
                        Next
                    End With
                End If
            Else
                TrnProcSuc = False
                objBaseClass.LogEntry("Invalid Input File")
                objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                objBaseClass.LogEntry("Input file :" + Path.GetFileName(strInputFileName) + " Is Moved to " + strArchivedFolderUnSuc)

            End If
            If TrnProcSuc <> False Then
                objBaseClass.LogEntry("Process Completed Successfully", False)
                objBaseClass.LogEntry("-------------------------------------------------------------------------------------", False)

            Else
                objBaseClass.LogEntry("Process Terminated", False)

                objBaseClass.LogEntry("-------------------------------------------------------------------------------------", False)
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Shapoorji_Palonji_Converter", "CmdProcess_Click")

        Finally

            If Not objFileValidate Is Nothing Then
                objBaseClass.ObjectDispose(objFileValidate.DtInput)
                objBaseClass.ObjectDispose(objFileValidate.DtUnSucInput)
                objBaseClass.ObjectDispose(objFileValidate.DtSuccEpayOutput)
                objBaseClass.ObjectDispose(objFileValidate.DtUnSucEpayOutput)
                objFileValidate.Dispose()
                objFileValidate = Nothing
            End If
        End Try
    End Sub
    Private Sub Payment_Report()
        Dim strSumFileName As String
        Dim Count_SuccRec As Integer = 0
        Dim Count_UnSuccRec As Integer = 0
        Try
            strSumFileName = "Transaction_Report_" & Path.GetFileNameWithoutExtension(gstrInputFile) & ".csv"

            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "[" & Format(Now, "dd-MM-yyyy hh:mm:ss") & "]")

            objBaseClass.WriteSummaryTxt(strSumFileName, "Transaction Report for Input File " & gstrInputFile)
            objBaseClass.WriteSummaryTxt(strSumFileName, "Beneficiary Name,Debit Account Number,Amount,IFSC CODE,Status,Reason")

            For Each row As DataRow In objFileValidate.DtSuccEpayOutput.Select()
                objBaseClass.WriteSummaryTxt(strSumFileName, Replace(row("Beneficiary Name").ToString, ",", "") & "," & Replace(row("Debit Account Number").ToString, ",", "") & "," & row("Amount").ToString & "," & row("Beneficiary Bank IFS Code").ToString & ",Successful," & row("Reason").ToString())
                Count_SuccRec += 1

            Next
            For Each row As DataRow In objFileValidate.DtUnSucEpayOutput.Select()
                objBaseClass.WriteSummaryTxt(strSumFileName, Replace(row("Beneficiary Name").ToString, ",", "") & "," & Replace(row("Debit Account Number").ToString, ",", "") & "," & row("Amount").ToString & "," & row("Beneficiary Bank IFS Code").ToString & ",UnSuccessful," & row("Reason").ToString())
                Count_UnSuccRec += 1
            Next
            objBaseClass.WriteSummaryTxt(strSumFileName, "Successful Record Count :" & Count_SuccRec)
            objBaseClass.WriteSummaryTxt(strSumFileName, "UnSuccessful Record Count :" & Count_UnSuccRec)
        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Payment_Report")

        End Try

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Timer1.Interval = 100
            Timer1.Enabled = True

            Generate_SettingFile()

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form Load", "form1_Load")
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Function GetAllSettings() As Boolean
        Try
            GetAllSettings = False

            If Not File.Exists(My.Application.Info.DirectoryPath & "\settings.ini") Then
                GetAllSettings = True
                MsgBox("Either settings.ini file does not exists or invalid file path" & vbCrLf & My.Application.Info.DirectoryPath & "\settings.ini", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            End If

            '-Audit Folder Path-
            If strAuditFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strAuditFolderPath) Then
                    Directory.CreateDirectory(strAuditFolderPath)
                    If Not Directory.Exists(strAuditFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Audit Log folder. Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", True)
                        End If
                        MsgBox("Invalid path for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                        Exit Function
                    End If
                End If
            End If

            '-Error Folder Path-
            If strErrorFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Error Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strErrorFolderPath) Then
                    Directory.CreateDirectory(strErrorFolderPath)
                    If Not Directory.Exists(strErrorFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Error Log folder. Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Error Log folder." & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Input Folder Path-
            If strInputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Input Folder " & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strInputFolderPath) Then
                    Directory.CreateDirectory(strInputFolderPath)
                    If Not Directory.Exists(strInputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Input Folder. Please check settings.ini file, the key as [ Input Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Input Folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "settings Error")
                    End If
                End If
            End If

            '-Archived Success Path-
            If strArchivedFolderSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archived Success folder" & vbCrLf & "Please check settings.ini file, the key as [ Archived Success Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderSuc) Then
                    Directory.CreateDirectory(strArchivedFolderSuc)
                    If Not Directory.Exists(strArchivedFolderSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archived Success Please check [ settings.ini ] file, the key as [ Archived Success Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archived Success Folder." & vbCrLf & "Please check settings.ini file, the key as [ Archived Success Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Archived Unsuccess Path-
            If strArchivedFolderUnSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archived Unsuccess folder" & vbCrLf & "Please check settings.ini file, the key as [ Archived Unsuccess Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderUnSuc) Then
                    Directory.CreateDirectory(strArchivedFolderUnSuc)
                    If Not Directory.Exists(strArchivedFolderUnSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archived Unsuccess Folder. Please check [ settings.ini ] file, the key as [ Archived Unsuccess Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archived Unsuccess Folder." & vbCrLf & "Please check settings.ini file, the key as [ Archived Unsuccess Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If
            ''''''''''''''Res Archived
            '-Archived Success Path-
            If strResArchivedFolderSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Response Archived Success folder" & vbCrLf & "Please check settings.ini file, the key as [ Response Archived Success Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strResArchivedFolderSuc) Then
                    Directory.CreateDirectory(strResArchivedFolderSuc)
                    If Not Directory.Exists(strResArchivedFolderSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Response Archived Success Please check [ settings.ini ] file, the key as [ Response Archived Success Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Response Archived Success Folder." & vbCrLf & "Please check settings.ini file, the key as [ Response Archived Success Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Res Archived Unsuccess Path-
            If strResArchivedFolderUnSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Response Archived Unsuccess folder" & vbCrLf & "Please check settings.ini file, the key as [ Response Archived Unsuccess Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strResArchivedFolderUnSuc) Then
                    Directory.CreateDirectory(strResArchivedFolderUnSuc)
                    If Not Directory.Exists(strResArchivedFolderUnSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Response Archived Unsuccess Folder. Please check [ settings.ini ] file, the key as [ Response Archived Unsuccess Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Response Unsuccess Folder." & vbCrLf & "Please check settings.ini file, the key as [ Response Archived Unsuccess Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Output Folder Path-
            If strEpayOutputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Epay Output Folder" & vbCrLf & "Please check settings.ini file, the key as [ Output Add Beneficiary Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strEpayOutputFolderPath) Then
                    Directory.CreateDirectory(strEpayOutputFolderPath)
                    If Not Directory.Exists(strEpayOutputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Epay Output Folder. Please check [ settings.ini ] file, the key as [ Epay Output Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Epay Output Folder." & vbCrLf & "Please check settings.ini file, the key as [ Epay Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Validation File Path-
            If strValidationPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Validation file." & vbCrLf & "Please check settings.ini file, the key as [ Validation ] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not File.Exists(strValidationPath) Then
                    GetAllSettings = True
                    If Not objBaseClass Is Nothing Then
                        objBaseClass.LogEntry("Error in settings.ini file, Validation file does not exist or invalid file path", True)
                    End If
                    MsgBox("Validation file does not exist or invalid file path" & vbCrLf & strValidationPath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                End If
            End If


            '-Temp Folder Path-
            If strTempFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Temp folder" & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strTempFolderPath) Then
                    Directory.CreateDirectory(strTempFolderPath)
                    If Not Directory.Exists(strTempFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Temp Folder. Please check [ settings.ini ] file, the key as [ Temp Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Temp Folder." & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            'Report
            If strReportFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Report folder" & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strReportFolderPath) Then
                    Directory.CreateDirectory(strReportFolderPath)
                    If Not Directory.Exists(strReportFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Report Folder. Please check [ settings.ini ] file, the key as [ Report Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Report Folder." & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Response Folder Path-
            If strResponseFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Response folder" & vbCrLf & "Please check settings.ini file, the key as [ Response Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strResponseFolderPath) Then
                    Directory.CreateDirectory(strResponseFolderPath)
                    If Not Directory.Exists(strResponseFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Output Folder. Please check [ settings.ini ] file, the key as [ Response Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Response Folder." & vbCrLf & "Please check settings.ini file, the key as [ Response Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Reverse Response Folder Path-
            If strReverseResponseFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Reverse Response folder" & vbCrLf & "Please check settings.ini file, the key as [ Reverse Response Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strReverseResponseFolderPath) Then
                    Directory.CreateDirectory(strReverseResponseFolderPath)
                    If Not Directory.Exists(strReverseResponseFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Reverse Response Folder. Please check [ settings.ini ] file, the key as [ Reverse Response Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Reverse Response Folder." & vbCrLf & "Please check settings.ini file, the key as [ Reverse Response Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

        Catch ex As Exception
            GetAllSettings = True
            'MsgBox("Error - " & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error While Getting Log Path from Settings.ini File")
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "GetAllSettings")

        End Try

    End Function


End Class
