# ReVBShell
## Files
* server.py - Interactive Python shell, listening on port 8080 for clients
* client.vbs - Visual Basic Script client which connectes to the IP/port specified and periodically fetches commands

## Components
### Server
_Interactive Python shell_

**Supported commands**
```
- CD [directory]     - Change directory. Shows current directory when without parameter.
- DOWNLOAD [path]    - Download the file at [path] to the .\Downloads folder.
- GETUID             - Get shell user id.
- GETWD              - Get working directory. Same as CD.
- HELP               - Show this help.
- IFCONFIG           - Show network configuration.
- KILL               - Stop script on the remote host.
- PS                 - Show process list.
- PWD                - Same as GETWD and CD.
- SET [name] [value] - Set a variable, for example SET LHOST 192.168.1.77.
                       When entered without parameters, it shows the currently set variables.
- SHELL [command]    - Execute command in cmd.exe interpreter;
                       When entered without command, switches to SHELL context.
- SHUTDOWN           - Exit this commandline interface (does not shutdown the client).
- SYSINFO            - Show sytem information.
- SLEEP [ms]         - Set client polling interval;
                       When entered without ms, shows the current interval.
- UNSET [name]       - Unset a variable
- UPLOAD [localpath] - Upload the file at [path] to the remote host.
                       Note: Variable LHOST is required.
- WGET [url]         - Download file from url.
```

### Client
_VBS client_
Configuration can be set in the .vbs file itself.
* strHost - IP of host to connect back to; should be the IP of the host where server.py is running
* strPort - Listening port on the above host
* intSleep - Default delay between the polls to the server

**Default settings**
```
strHost = "127.0.0.1"
strPort = "8080"
intSleep = 5000
```
