'This Code Written with Visual Studio 6 - VBA - Outlook Macro
' (1) copies one file from one folder to another folder with the VBA FileSystemObject
' (2) requires a reference to the object library "Microsoft Scripting Runtime" under Options > Tools > References in the VBE.
' (3) requires a referance to the object library "ActiveX Data Object 6.1 Lib"
' (4) requires a referance to the object library "OLE Automation"

Option Explicit
Private WithEvents inboxItems As Outlook.Items
'Creating a FileSystemObject
Public FSO As New FileSystemObject
Private Sub Application_Startup()
  Dim outlookApp As Outlook.Application
  Dim objectNS As Outlook.NameSpace

  Set outlookApp = Outlook.Application
  Set objectNS = outlookApp.GetNamespace("MAPI")
  Set inboxItems = objectNS.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub inboxItems_ItemAdd(ByVal Item As Object)
On Error GoTo ErrorHandler
Dim Msg As Outlook.MailItem
Dim MessageInfo
Dim Result
Dim lngCount As Long
Dim objAttachments As Outlook.Attachments
Dim i As Long
Dim strFile As String
Dim CUSTOMER As String   'Customer characters "_Customer_"'
Dim isCustomer As Boolean
Dim strFolderpath As String

    CUSTOMER = "_Customer_"  'Define Customer characters'
    
    ' Get the path to your My Documents folder
    strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)

    ' Set the Attachment folder.
    strFolderpath = strFolderpath & "\FileFolder\"
    
    If (FSO.FolderExists(strFolderpath)) Then
        'File exist DO NOTHÝNG
        
    Else
        'File not exist
        FSO.CreateFolder (strFolderpath)
    End If
    MsgBox (strFolderpath)
    
    
  
If TypeName(Item) = "MailItem" Then

    Dim MailId As String
    'We define to store mail id
    MailId = Format(Now(), "yyyy-MM-dd-hh-mm-ss") & "-" & Item.SenderEmailAddress
    'EntryID
     'Item is mail object, Item.SenderEmailAddress is sender mail : lorem@outlook.com
     'Item.EntryID is uniq mail id like this 00000ERS0123412410000......
     '.SentOn , .RecievedTime , .Subject , .Size , .Body
     

    
     
    Set objAttachments = Item.Attachments 'Outlook Atactment Object
    lngCount = objAttachments.Count

    If lngCount > 0 Then
    
    ' Use a count down loop for removing items
    ' from a collection. Otherwise, the loop counter gets
    ' confused and only every other item is removed.
    
    For i = lngCount To 1 Step -1
    
    
    
    ' Get the file name.
    strFile = "\" & objAttachments.Item(i).FileName
    MsgBox (strFile)
    'We will check if Filename contain _Customer_ return True else return false
    isCustomer = isContain(strFile, CUSTOMER)

    If isCustomer = True Then
    strFolderpath = strFolderpath & MailId & "\"
    
    If (FSO.FolderExists(strFolderpath)) Then
        'File exist DO NOTHÝNG
        
    Else
        'File not exist
        FSO.CreateFolder (strFolderpath)
    End If
    MsgBox (strFolderpath)
    
        strFile = strFolderpath & strFile
        ' Save the attachment as a file.
        objAttachments.Item(i).SaveAsFile strFile
        
        Call saveAsBinarySQL(strFolderpath, objAttachments.Item(i).FileName, MailId)
    End If
    
    Next i
    
    'If want to delete mail after save computer and MSSQL server, you can remove comment Item.Delete
    'Item.Delete
    
    End If

End If
ExitNewItem:
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume ExitNewItem
End Sub

Function isContain(ByVal i As String, ByVal k As String) As Boolean
    'i is File name string, k is Customer String'
Dim tempNumber As Integer

    For tempNumber = 1 To (Len(i) - Len(k))
    
        If Mid(i, tempNumber, Len(k)) = k Then
            isContain = True
        End If
    
    Next tempNumber
    
End Function

Sub saveAsBinarySQL(FileLocation As String, FileName As String, MailIds As String)
'To save a file in a table as binary
'-----------------------------------------
' About SQL TABLE
'id -Primarly Key And Auto Increment - INT
'binaryFile - VARBINARY(MAX) - XLS or some file format convert to Binary File we store binary format in table
'name - NVARCCHAR(MAX) - XLS Document name
'date - DateTime - Recieve Mail Time - Maybe we sorting Time or sominting
'mailId - NVARCHAR(MAX) - Uniq mail id
'-----------------------------------------

    Dim adoStream       As Object
    Dim adoCmd          As Object
    Dim strFilePath     As String
    Dim adoCon          As Object
    Const strDB         As String = "Krautz"  'Database name
    Const strServerName As String = "DESKTOP-G9CS4CF\SQLEXPRESS"  'Server Name


    Set adoCon = CreateObject("ADODB.Connection")
    Set adoStream = CreateObject("ADODB.Stream")
    Set adoCmd = CreateObject("ADODB.Command")
    
    '--Open Connection to SQL server ------------------
    adoCon.CursorLocation = adUseClient
    adoCon.Open "Provider=SQLOLEDB;Data Source=" & strServerName & ";Initial Catalog = " & strDB & ";Integrated Security=SSPI;"
    '--------------------------------------------------
    
    
    strFilePath = FileLocation & FileName
    
    
    adoStream.Type = adTypeBinary
    adoStream.Open
    adoStream.LoadFromFile strFilePath 'It fails if file is open
        
    'Need To Enter Table Name
    With adoCmd
        .CommandText = "INSERT INTO SaveToBinary VALUES (?,?,?,?)" ' Query
        .CommandType = adCmdText
        
        '---adding parameters
        .Parameters.Append .CreateParameter("@binaryFile", adVarBinary, adParamInput, adoStream.Size, adoStream.Read)
        .Parameters.Append .CreateParameter("@name", adVarChar, adParamInput, 200, FileName)
        .Parameters.Append .CreateParameter("@date", adDBDate, adParamInput, 133, Date)
        .Parameters.Append .CreateParameter("@mailId", adVarChar, adParamInput, 200, MailIds)
        '---
    End With
    
    adoCmd.ActiveConnection = adoCon
    adoCmd.Execute
        
    adoCon.Close
    
    ' Dim MailFolder As String
    
    ' MailFolder = FileLocation & MailIds
    
    ' 'We create uniq mail folder, same all mail attactment saved uniq folder
    ' If (FSO.FolderExists(MailFolder)) Then
    '     'File Exist - Do Nothing cause we define
    ' Else
    '     'File does not exist
    '     FSO.CreateFolder (MailFolder)
    ' End If
    
    'FSO.CopyFile strFilePath, MailFolder
    'Call CopyFileWithFSOBasic(strFilePath, MailFolder, False)
    
    'FSO.DeleteFile strFilePath, True
    
End Sub

Sub CopyFileWithFSOBasic(SourceFilePath As String, DestPath As String, OverWrite As Boolean)
    Call FSO.CopyFile(SourceFilePath, DestPath, OverWrite)
    MsgBox ("Baþarýlý")

End Sub
