Imports System.IO

Module GenrateOutput

    Dim objLogCls As New ClsErrLog
    Dim objGetSetINI As ClsShared
    Dim objBaseClass As ClsBase
    Dim objValidationClass As ClsValidation
    Dim SumOfAmount As Double = 0

    Public Function GenerateOutPutFile(ByRef dtOutput As DataTable, ByVal strFileName As String) As Boolean
        Dim gstrA2Afile As String = String.Empty
        Dim strMethodCalForEpay As Boolean = False
        Try
            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            objValidationClass = New ClsValidation(strFileName, objBaseClass.gstrIniPath)
            FileCounter = objBaseClass.GetINISettings("General", "File Counter", My.Application.Info.DirectoryPath & "\settings.ini")
            FileCounter = FileCounter + 1

            If Len(FileCounter) < 3 Then
                FileCounter = FileCounter.PadLeft(4, "0").Trim()
                FileCounter = FileCounter.Substring(FileCounter.Length - 3, 3)
            End If

            strFileName = (objValidationClass.IsJustAlpha(Path.GetFileNameWithoutExtension(gstrInputFile), 10, "N")).Replace(" ", "")

            Dim strOptFileName As String = ""

            '  strOptFileName = strFileName & "_" & Format(Now.Date(), "ddMMyy") & "_" & FileCounter.ToUpper

            strOptFileName = strFileName.Replace("_req", "").Replace("_REQ", "")

            Dim FileCount As Integer = dtOutput.DefaultView.ToTable(True, "File_No").Rows.Count
            For index = 1 To FileCount
                If FileCount = 1 Then
                    gstrOutputFile_EPAY = strOptFileName & ".txt"
                Else
                    gstrOutputFile_EPAY = strOptFileName & "_SP" & index & ".txt"
                End If

                If Generate_OutPut_Epay(dtOutput, gstrOutputFile_EPAY, index) = False Then
                    GenerateOutPutFile = False
                Else
                    GenerateOutPutFile = True
                    OptFile_EncryptionProc(gstrOutputFile_EPAY)

                    Call objBaseClass.SetINISettings("General", "File Counter", Val(FileCounter), My.Application.Info.DirectoryPath & "\settings.ini")
                End If
            Next


        Catch ex As Exception
            GenerateOutPutFile = False
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "GenerateOutPutFile")
        End Try


    End Function

    Private Function Generate_OutPut_Epay(ByRef _dtRBI As DataTable, ByVal strFileName As String, ByVal FileNo As Integer) As Boolean
        Dim strOutPutLine As String
        Dim objStrmWriter As StreamWriter
        Dim FileDateTime As String
        Dim TotalAmt As Double = 0
        Dim strCompanyCode As String
        Try

            If _dtRBI.Rows.Count > 0 Then
                'objStrmWriter = New StreamWriter(strEpayOutputFolderPath & "\" & strFileName)
                objStrmWriter = New StreamWriter(strTempFolderPath & "\" & strFileName)

                'FileDateTime = DateTime.Now.ToString("ddMMyyyyHHmmss")
                FileDateTime = DateTime.Now.ToString("ddMMyyyy")
                ' strCompanyCode = Pad_Length("CHIPFUN" & FileDateTime.ToString() & FileCounter.PadLeft(3, "0").Trim().PadRight(20, " "), 20)
                strCompanyCode = Pad_Length("CHIPFUN" & FileDateTime.ToString().PadRight(20, " "), 20)
                '-------H Record

                strOutPutLine = ""
                strOutPutLine = "H" & Pad_Length(Format(Now, "dd/MM/yyyy").PadRight(10, " "), 10) & strCompanyCode
                objStrmWriter.WriteLine(strOutPutLine)

                For Each drRow As DataRow In _dtRBI.Select("File_No=" & FileNo)
                    'Heder 
                    ''-------D Record
                    strOutPutLine = ""

                    For intIndex As Int32 = 0 To _dtRBI.Columns.Count - 6
                        strOutPutLine = strOutPutLine & drRow(intIndex).ToString()
                    Next
                    TotalAmt = TotalAmt + Convert.ToDouble(drRow("Amount").ToString().Trim())
                    objStrmWriter.WriteLine(strOutPutLine, strFileName)
                Next

                '-------F Record
                strOutPutLine = "F" & (_dtRBI.Rows().Count).ToString().PadLeft(5, "0") & TotalAmt.ToString(".00").PadLeft(14, "0")
                objStrmWriter.WriteLine(strOutPutLine)

                objBaseClass.LogEntry("Epay Output file [" & strFileName & "] is generated successfully", False)

                Generate_OutPut_Epay = True
            Else
                objBaseClass.LogEntry("Epay Record Not Found")
                Generate_OutPut_Epay = False
            End If
        Catch ex As Exception
            Generate_OutPut_Epay = False
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Generate_Output", "GenerateEpayOutPutFile")
        Finally
            If Not objStrmWriter Is Nothing Then
                objStrmWriter.Close()
                objStrmWriter.Dispose()

            End If
        End Try
    End Function

    Public Sub OptFile_EncryptionProc(ByVal strEpayOptFileName As String)
        Dim stremWriter As StreamWriter = Nothing
        Try
            If strYBLEncryptioniRequired.ToString().Trim().ToUpper() = "Y" Then
                objBaseClass.LogEntry("Encryption process start for file " & strEpayOptFileName)
                objBaseClass.LogEntry("Check Batch File is Available for processing " & strYBLBatchFilePath)

                System.Threading.Thread.Sleep(2000)

                Dim maxRetry = 5
                Dim strOpenFileClosed As Boolean = False

                Dim file As FileInfo = New FileInfo(strYBLBatchFilePath)
                If IsFileOpen(file) = True Then
                    objBaseClass.LogEntry("file is being used by another process " & file.Name)
                    For retry As Integer = 0 To maxRetry - 1
                        System.Threading.Thread.Sleep(2000)
                        If IsFileOpen(file) = False Then
                            strOpenFileClosed = True
                            Exit For
                        End If
                    Next
                    If strOpenFileClosed = False Then
                        If Not stremWriter Is Nothing Then
                            stremWriter.Close()
                            stremWriter.Dispose()
                            objBaseClass.LogEntry("StremWriter object closed for encryption" & file.Name)
                        End If
                    End If
                End If


                ''Encryption Payment Output File
                If IsFileOpen(file) = False Then

                    'objBaseClass.LogEntry("Encryption process start for file " & striGTBOptFileName)
                    objBaseClass.FileMove(strTempFolderPath & "\" & strEpayOptFileName, strYBLPICKDIRPath & "\" & strEpayOptFileName.Replace(" ", ""))
                    objBaseClass.LogEntry(strEpayOptFileName & " file move in " & strYBLPICKDIRPath)
                    strEpayOptFileName = strEpayOptFileName.Replace(" ", "")

                    System.Threading.Thread.Sleep(2000)
                    ''-Encryption
                    objBaseClass.FileDelete(strYBLBatchFilePath & "\" & "Test.bat")
                    stremWriter = New StreamWriter(strYBLBatchFilePath) ''''

                    objBaseClass.LogEntry("Batch File Generated  Successfully " & strYBLBatchFilePath)
                    stremWriter.WriteLine("cd\")
                    stremWriter.WriteLine(strEncryption_Decryption_Dir)
                    stremWriter.WriteLine("cd " & strEncryption_Decryption_Folder & "encrypt")
                    stremWriter.WriteLine("encrypt " & strYBLPICKDIRPath & "\" & strEpayOptFileName & " " & strYBLDROPDIRPath & "\" & strEpayOptFileName & ".enc" & " " & strYBLCRCDIRPath & "\" & strEpayOptFileName & ".crc")

                    stremWriter.WriteLine("END")

                    'objBaseClass.ObjectDispose(stremWriter)

                    If Not stremWriter Is Nothing Then
                        stremWriter.Close()
                        stremWriter.Dispose()
                        objBaseClass.LogEntry("StremWriter object closed for encryption_1" & file.Name)
                    End If
                    ''''''''''''''''''''''''''

                    System.Threading.Thread.Sleep(4000)
                    objBaseClass.LogEntry("Batch call" & strYBLBatchFilePath)

                    objBaseClass.Execute_Batch_file_E()
                    objBaseClass.LogEntry("YBL Encrypting file " & strEpayOptFileName & " is Completed by YBL.")
                    ''-
                    objBaseClass.FileMove(strYBLDROPDIRPath & "\" & strEpayOptFileName & ".enc", strEpayOutputFolderPath & "\" & strEpayOptFileName & ".enc")

                    Threading.Thread.Sleep(4000)

                    objBaseClass.FileDelete(strYBLPICKDIRPath & "\" & strEpayOptFileName)
                    objBaseClass.FileDelete(strYBLCRCDIRPath & "\" & strEpayOptFileName & ".crc")
                Else
                    If Not stremWriter Is Nothing Then
                        stremWriter.Close()
                        stremWriter.Dispose() 'objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("StremWriter object closed for encryption_2" & file.Name)
                    End If

                End If ''-
            Else
                objBaseClass.FileMove(strTempFolderPath & "\" & strEpayOptFileName, strEpayOutputFolderPath & "\" & strEpayOptFileName)
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "OptFile_EncryptionProc")
        Finally
            If Not stremWriter Is Nothing Then
                stremWriter.Close()
                stremWriter.Dispose()
            End If
        End Try

    End Sub

    Public Function Generate_Output_Response(ByRef _dtRes As DataTable, ByVal strFileName As String) As Boolean

        Dim strData As String = ""
        Dim objStrmWriter As StreamWriter = Nothing
        Dim ColDisplay As String() = {"RecIdentifier", "File name", "TXN_NO", "File_NO", "Reason", "Additional Field2", "Additional Field3", "CheckerID"}
        Dim displayColAsText As String() = {"CUSTOMERREFNUM", "AMOUNT", "REFERENCENO", "CREDITDATE"}

        Try
            If (_dtRes IsNot Nothing) Then
                If (_dtRes.Rows.Count > 0) Then

                    If objBaseClass Is Nothing Then
                        objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
                    End If
                    objValidationClass = New ClsValidation(strFileName, objBaseClass.gstrIniPath)

                    gstrResOutputfile = (objValidationClass.IsJustAlpha(Path.GetFileNameWithoutExtension(strFileName), 10, "N")).Replace(" ", "").Replace("_req", "").Replace("_REQ", "").Replace("_EOD", "").Replace("_eod", "") & ".csv"

                    objStrmWriter = New StreamWriter(strReverseResponseFolderPath & "\" & gstrResOutputfile)
                    objBaseClass.LogEntry("Reverse Output File generating process Started....")

                    strData = ""
                    Dim Counter As Integer
                    Counter = 0
                    For Each drRow As DataRow In _dtRes.Rows
                        Counter += 1
                        strData = ""
                        For Inti As Int32 = 0 To drRow.ItemArray.Length - 4
                            If Not ColDisplay.Contains(_dtRes.Columns(Inti).ColumnName) Then
                                'If displayColAsText.Contains(_dtRes.Columns(Inti).ColumnName.ToUpper) Then 'Commented by swati dtd 2023-04-25
                                '    strData = strData & "'" & (drRow.ItemArray(Inti).ToString()) & ","
                                'Else
                                '    strData = strData & (drRow.ItemArray(Inti).ToString()) & ","
                                'End If
                                ''''''''''''''''''''''''''''''''''''
                                strData = strData & (drRow.ItemArray(Inti).ToString()) & ","
                            End If
                        Next
                        strData = strData.Substring(0, strData.Length - 1)

                        If Counter < _dtRes.Rows.Count Then
                            objStrmWriter.WriteLine(strData, strFileName)
                        ElseIf Counter = _dtRes.Rows.Count Then
                            objStrmWriter.Write(strData, strFileName)
                        End If
                    Next

                    If Not objStrmWriter Is Nothing Then
                        objStrmWriter.Close()
                        objStrmWriter.Dispose()
                    End If
                    objBaseClass.LogEntry("Reverse Response Output File [" & strFileName & "] is  Generated Successfully")

                    Generate_Output_Response = True
                Else
                    objBaseClass.LogEntry("No Records Found to Generate Reverse Response Output File.")
                    Generate_Output_Response = False
                End If
            End If
        Catch ex As Exception
            Generate_Output_Response = False
            Call objBaseClass.Handle_Error(ex, "GenrateOutput", Err.Number, "Generate_Output_Response")
        End Try


    End Function

    Public Function IsFileOpen(ByVal file As FileInfo) As Boolean
        Dim stream As FileStream = Nothing
        IsFileOpen = True
        Try
            stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
            Return False
        Catch ex As Exception
            Return True
        End Try
    End Function
    Public Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then

                ' Check_Comma = Chr(34) & strTemp & Chr(34) & ","
                Check_Comma = strTemp
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Payment", "Check_Comma")

        End Try
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function


    Function RemoveCharacter(ByVal stringToCleanUp As String)
        Dim characterToRemove As String = ""
        characterToRemove = Chr(34) + "=~^!#$%&'()*+,-@`/\:{}[]"

        Dim firstThree As Char() = characterToRemove.Take(30).ToArray()
        For index = 0 To firstThree.Length - 1
            stringToCleanUp = stringToCleanUp.ToString.Replace(firstThree(index), "")
        Next
        Return stringToCleanUp
    End Function
End Module
