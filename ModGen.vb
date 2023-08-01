Option Explicit On

Module ModGen


    Public blnErrorLog As Boolean
    Public strAuditFolderPath As String
    Public strErrorFolderPath As String

    Public strInputFolderPath As String
    Public gstrInputFile As String
    Public gstrInputFolder As String

    Public gstrOutputFile_EPAY As String

    Public strEpayOutputFolderPath As String
    Public strReportFolderPath As String
    Public strTempFolderPath As String             ' Temp folder path
    ''Archive
    Public strArchivedFolderSuc As String
    Public strArchivedFolderUnSuc As String

    ''Res_Archive
    Public strResArchivedFolderSuc As String
    Public strResArchivedFolderUnSuc As String
    ''''''''''''''''''

    ''Res
    Public gstrResOutputfile As String
    Public gstrResponseInputFolder As String
    Public gstrResponseInputFile As String

    Public strResponseFolderPath As String             ' Response folder path
    Public strReverseResponseFolderPath As String            ' RevResponse folder path

    Public strProceed As String
    Public strInvalidTrans As String
    Public FileCounter As String
    Public strTransactionNo As String
    Public strValidationPath As String

    '-Client Details-
    Public strClientCode As String
    Public strClientName As String
    Public strInputDateFormat As String

    ''''Epay details
    Public strDebitAccNo As String
    Public strRemitter_Name As String
    'Encryption

    Public strYBLEncryptioniRequired As String

    Public strEncrypt As String
    Public strYBLBatchFilePath As String
    Public strYBLPICKDIRPath As String
    Public strYBLDROPDIRPath As String
    Public strYBLCRCDIRPath As String

    Public strEncryption_Decryption_Dir As String
    Public strEncryption_Decryption_Folder As String

    Public intTimerIntrvl As Integer

End Module


