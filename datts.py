#!/usr/bin/python2
# coding: utf8

'''
    datts - download and save attachments from IMAP server.

    For help type: ./datts.py --help

    v0.2.0
'''

import os
import sys
import socket
from signal import signal, SIGINT
import Queue
from threading import Thread, Lock
from time import time, gmtime, strftime
from getopt import getopt, GetoptError
from getpass import getpass, GetPassWarning
from imaplib import IMAP4, IMAP4_SSL
from email import message_from_string
from email.header import decode_header

def handler(signum, frame):
    '''
    Handler function for SIGINT.
    '''

    q_mail_in.put(SENTINEL)
    print '\n- Interrupted by Ctrl+C, please wait...'

signal(SIGINT, handler)

# Default port for IMAP over SSL.
SERVER_PORT = 993
# Max number of workers
MAX_THREAD_POOL_SIZE = 10
# Triggers exit in the threads.
SENTINEL = object()
# Stats
count_mail = 0
count_att = 0

q_mail_in = Queue.LifoQueue(maxsize=0)
q_mail_out = Queue.LifoQueue(maxsize=0)

def main():
    '''
    Main function.
    '''

    global q_mail_in, q_mail_out

    # Number of workers.
    thread_pool_size = 1

    if len(sys.argv) < 2:
        usage()
        sys.exit(0)

    # True if option is required.
    options = [
        ['login=', True],
        ['server=', True],
        ['dir=', True],
        ['mbox=', True],
        ['n=', False],
        ['delete', False],
        ['help', False],
        ['dump', False],
        ['noinline', False],
        ['c=', False]
    ]

    try:
        opts, _ = getopt(sys.argv[1:], '', [options[x][0] for x in range(len(options))])
    except GetoptError as err:
        print str(err)
        sys.exit(1)

    missing = []
    gopts = [opts[i][0].replace('--', '') for i in range(len(opts))]
    for i in range(len(options)):
        if 'help' in gopts:
            usage()
            sys.exit(0)
        if options[i][0].replace('=', '') not in gopts and options[i][1] is True:
            missing.append(options[i][0])

    if missing:
        print 'Required options: {}'.format(missing)
        sys.exit(1)

    # Options ready to use.
    ok_opts = dict()
    for opt, val in opts:
        if opt == '--login':
            ok_opts['username'] = val
        if opt == '--server':
            ok_opts['server_name'] = val
        if opt == '--dir':
            ok_opts['directory'] = unicode(val, 'utf8')
        if opt == '--mbox':
            ok_opts['mbox'] = val
            ok_opts['mbox_encoded'] = utf7mod_encode(unicode(val, 'utf8'))
        if opt == '--n':
            ok_opts['how_many'] = val
        if opt == '--delete':
            ok_opts['delete'] = True
        if opt == '--dump':
            ok_opts['dump'] = True
        if opt == '--noinline':
            ok_opts['noinline'] = True
        if opt == '--c':
            thread_pool_size = val

    if 'dump' in ok_opts:
        print '\nOptions dump:\n'
        for opt, val in opts:
            if opt == '--dump':
                continue
            else:
                print '{:10} = {}'.format(opt, val)
        print '\n'
        sys.exit(0)

    if not os.path.isdir(ok_opts['directory']):
        print 'No such directory: {}'.format(ok_opts['directory'])
        sys.exit(1)

    if 'how_many' in ok_opts:
        try:
            ok_opts['how_many'] = int(ok_opts['how_many'])
            if ok_opts['how_many'] < 0:
                raise ValueError
        except ValueError:
            print 'Option of messages(--n) must be positive number.'
            sys.exit(1)
    else:
        ok_opts['how_many'] = None

    try:
        thread_pool_size = int(thread_pool_size)
        if thread_pool_size < 1 or thread_pool_size > MAX_THREAD_POOL_SIZE:
            raise ValueError
    except ValueError:
        print 'Option of connections(--c) must be positive number and <= {}'.format(MAX_THREAD_POOL_SIZE)
        sys.exit(1)

    try:
        ok_opts['password'] = getpass()
    except GetPassWarning as err:
        print str(err)

    print_banner()

    #start_time = time()
    print '- Starting...'

    # Fetch UIDs
    uids, msg_count = get_uids(SERVER_PORT, **ok_opts)

    print '- Number of messages in {} mailbox: {}'.format(ok_opts['mbox'], msg_count[0])
    print '- Starting download:'

    # Fill queue with UIDs.
    for mail_uid in uids[:ok_opts['how_many']]:
        q_mail_in.put(mail_uid)

    lock = Lock()

    threads = [Thread(target=proccess_message, args=(q_mail_in,
                                                     q_mail_out,
                                                     lock,
                                                     SERVER_PORT),
                      kwargs=ok_opts) for _ in range(thread_pool_size)]

    for thread in threads:
        thread.start()

    # Start main thread loop.
    progress_loop()

    for thread in threads:
        thread.join(1.0)

# End of function main().

def get_uids(server_port, **ok_opts):
    '''
    Obtain message UIDs from IMAP server.
    Returns UIDs(list) and total count of messages in the mailbox.
    '''

    server = connect(server_port, **ok_opts)

    print '- Connected to server'

    select_return_code, msg_count = server.select(ok_opts['mbox_encoded'])
    if select_return_code == 'OK':

        _, uids_data = server.uid('search', None, 'ALL')
        uids = uids_data[0].split()

        if not uids:
            print '- No messages found in: {}'.format(ok_opts['mbox'])
            server.close()
            server.logout()
            sys.exit(0)

        # SORT command may not be available in all IMAP servers. rfc3501, rfc5256.
        # Instead we are sorting UIDs locally.
        uids = [int(x) for x in uids]
        uids.sort(reverse=True)

        # Main thread disconnects from server.
        server.close()
        server.logout()
    else:
        print '- Cannot select given mailbox: {}'.format(ok_opts['mbox'])
        server.logout()
        sys.exit(1)

    return uids, msg_count

def proccess_message(q_in, q_out, lock, server_port, **ok_opts):
    '''
    Worker function.
    '''

    # Variables, shared between threads.
    global count_mail, count_att

    server = connect(server_port, **ok_opts)

    select_return_code, _ = server.select(ok_opts['mbox_encoded'])
    if select_return_code == 'OK':

        while not q_in.empty():
            try:
                message_id = q_in.get()
            except Queue.Empty:
                server.close()
                server.logout()
                break

            if message_id is SENTINEL:
                # Ctrl+C pressed, propagate to other threads.
                q_in.put(SENTINEL)
                server.close()
                server.logout()
                break

            _, mail_data = server.uid('fetch', message_id, '(RFC822)')
            mail = message_from_string(mail_data[0][1])

            subject = mail.get('subject')
            if subject:
                subject_text = decode_header(subject)
                subject_text = reconstruct_subject(subject_text)
            else:
                subject_text = '-'

            mail_all_parts = []
            for part in mail.walk():

                content = part.get('Content-Disposition')
                if content:
                    if 'noinline' in ok_opts and 'inline' in content:
                        continue
                filename = part.get_filename()

                if filename is not None:

                    # Filename can be UTF8, quoted-printable.
                    name, charset = decode_header(filename)[0]
                    if charset:
                        filename = name.decode(charset)

                    file_path = os.path.join(ok_opts['directory'], filename)
                    file_path = generate_unique_filename(ok_opts['directory'], file_path)
                    file_path = file_path.encode('utf-8')

                    try:
                        with open(file_path, 'wb') as output_file:
                            output_file.write(part.get_payload(decode=True))
                    except IOError as err:
                        file_error = err
                    else:
                        file_error = None

                    with lock:
                        count_att += 1
                else:
                    file_path = None
                    file_error = None

                # Collect all parts of the message...
                thread_msg_part = (message_id, subject_text, file_path, file_error)
                mail_all_parts.append(thread_msg_part)
            # and put them into the q_out queue so they arrive in order.
            q_out.put(mail_all_parts)

            # Delete message.
            if 'delete' in ok_opts and file_error is None:
                server.uid('store', message_id, '+FLAGS', '(\\Deleted)')
                server.expunge()

            with lock:
                count_mail += 1

            if file_error is not None:
                break
    else:
        print '- Cannot select given mailbox: {}'.format(ok_opts['mbox'])
        server.logout()
        sys.exit(1)

def progress_loop():
    '''
    Print progress messages.
    '''

    start_time = time()

    # Main thread loop.
    while True:

        message_id_prev = ''
        attachment_found_count = 0

        try:
            work = q_mail_out.get(timeout=5.0)

            for message_id, subject_text, file_path, file_error in work:

                # Print subject only once.
                if message_id != message_id_prev:
                    print 'Message: {}'.format(subject_text)
                    message_id_prev = message_id

                if file_path is not None:
                    attachment_found_count += 1

                if isinstance(file_error, Exception):
                    try:
                        raise file_error
                    except IOError as err:
                        print '- File saving error: {}, '.format(err)

                if file_error is None and file_path is not None:
                    print '- Saving file to: {}'.format(file_path)

            if attachment_found_count == 0:
                print '- Attachment not found.'

        except Queue.Empty:
            print '- Finished downloading'
            break

    end_time = strftime('%H:%M:%S', gmtime((time() - start_time)))

    print '----------------------------------------'
    print 'Summary'
    print '----------------------------------------'
    print 'Number of messages processed: {:>10}'.format(count_mail)
    print 'Number of attachments: {:>17}'.format(count_att)
    print 'Time taken: {:>28}'.format(end_time)

def connect(server_port, **ok_opts):
    '''
    Connect to IMAP server.
    '''

    try:
        server = IMAP4_SSL(ok_opts['server_name'], server_port)
        server.login(ok_opts['username'], ok_opts['password'])
    except socket.error:
        print 'Unable to connect to: {}'.format(ok_opts['server_name'])
        sys.exit(1)
    except IMAP4.error as err:
        print 'Unable to login: {} {}'.format(err, 'Check your credentials.')
        sys.exit(1)

    return server

def generate_unique_filename(directory, file_path):
    '''
    Avoid overwriting files by adding sequence number to their names.
    '''

    sequence = 1

    while os.path.isfile(file_path):
        new_filename = os.path.basename(file_path).split('.')[0] + '(' + str(sequence) + ')'
        new_filename = new_filename + '.' + '.'.join(os.path.basename(file_path).split('.')[1:])
        file_path = os.path.join(directory, new_filename)
        sequence += 1

    return file_path

def utf7mod_encode(text):
    '''
    Encode text in modified UTF7.
    IMAP4rev1 international mailbox names are encoded with modified UTF7, rfc3501 #section-5.1.3
    '''

    encoded = ''

    for char in text:

        if char >= u'\x20' and char <= u'\x7e':
            if char == '&':
                encoded += '&-'
            else:
                encoded += char
        else:
            encoded += char

    return encoded.encode('utf7').replace('/', ',').replace('+', '&')

def reconstruct_subject(subject):
    '''
    Reconstruct subject from internal representation.
    Can be split into individual parts with encoding charset type for each part.
    '''

    complete_subject = ''

    for subject_part in subject:

        if subject_part[1] is not None:
            complete_subject += subject_part[0].decode(subject_part[1]) + ' '
        else:
            complete_subject += subject_part[0] + ' '

    return complete_subject.lstrip().encode('utf-8')

def usage():
    '''
    Print usage.
    '''

    print '\ndatts - download and save attachments from IMAP server'
    print '\nUsage: ', sys.argv[0],
    print '''--login --mbox --dir --server [--n] [--c] [--delete] [--dump] [--noinline] [--help]

    Option      Argument    Description
    -------------------------------------------
    --help                  show this help  
    --login     string      login to your account
    --server    string      server name
    --mbox      string      remote folder with attachments
    --dir       string      local folder for storing attachments
    --n         number      how many messages to download? Default is all of them.
    --c         number      how many connections to start? Default is 1, max is 10.
    --delete                should we delete message after download? Default is to NOT delete.
    --dump                  print provided options and exit
    --noinline              skip attachments embedded in message body text
    '''

def print_banner():
    '''
    Print nice banner.
    '''

    # ASCII art generated using `Text to ASCII Art Generator` at http://patorjk.com

    print r'''
           __      __  __      
      ____/ /___ _/ /_/ /______
     / __  / __ `/ __/ __/ ___/
    / /_/ / /_/ / /_/ /_(__  ) 
    \__,_/\__,_/\__/\__/____/ v0.2.0
    '''

if __name__ == '__main__':
    main()
