#!/usr/bin/python2
# coding: utf8

'''
    datts - download and save attachments from IMAP server.

    For help type: ./datts.py --help

    v0.1.0
'''

import os
import sys
import socket
from time import time, gmtime, strftime
from getopt import getopt, GetoptError
from getpass import getpass, GetPassWarning
from imaplib import IMAP4, IMAP4_SSL
from email import message_from_string
from email.header import decode_header

# Default port for IMAP over SSL.
SERVER_PORT = 993

def main():
    '''
    Main function.
    '''

    delete = False
    dump = False
    how_many = None
    noinline = False

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
        ['noinline', False]
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

    for opt, val in opts:
        if opt == '--login':
            username = val
        if opt == '--server':
            server_name = val
        if opt == '--dir':
            directory = unicode(val, 'utf8')
        if opt == '--mbox':
            mbox = val
            mbox_encoded = utf7mod_encode(unicode(val, 'utf8'))
        if opt == '--n':
            how_many = val
        if opt == '--delete':
            delete = True
        if opt == '--dump':
            dump = True
        if opt == '--noinline':
            noinline = True

    if dump:
        print '\nOptions dump:\n'
        for opt, val in opts:
            if opt == '--dump':
                continue
            else:
                print '{:10} = {}'.format(opt, val)

        print '\n'
        sys.exit(0)

    if not os.path.isdir(directory):
        print 'No such directory: {}'.format(directory)
        sys.exit(1)

    if how_many is not None:
        try:
            how_many = int(how_many)
            if how_many < 0:
                raise ValueError
        except ValueError:
            print 'Number of messages(--n) must be a positive value.'
            sys.exit(1)

    try:
        password = getpass()
    except GetPassWarning as err:
        print str(err)

    banner()

    start = time()
    print '- Starting...'

    try:
        server = IMAP4_SSL(server_name, SERVER_PORT)
        server.login(username, password)
    except socket.error:
        print 'Unable to connect to: {}'.format(server_name)
        sys.exit(1)
    except IMAP4.error as err:
        print 'Unable to login: {}'.format(err)
        sys.exit(1)

    print '- Connected to server'

    select_return_code, msg_count = server.select(mbox_encoded)
    if select_return_code == 'OK':

        _, uids_data = server.uid('search', None, 'ALL')
        uids = uids_data[0].split()

        if not uids:
            print 'No messages found in: {}'.format(mbox)
            server.close()
            server.logout()
            sys.exit(0)

        # SORT command may not be available in all IMAP servers. rfc3501, rfc5256.
        # Instead we are sorting UIDs locally.
        uids = [int(x) for x in uids]
        uids.sort(reverse=True)

        print '- Number of messages in {} mailbox: {}'.format(mbox, msg_count[0])

        # Stats
        count_mail = 0
        count_att = 0

        print '- Starting download:'

        for mail_id in uids[:how_many]:

            attachment_present = False

            _, mail_data = server.uid('fetch', mail_id, '(RFC822)')
            mail = message_from_string(mail_data[0][1])

            subject = mail.get('subject')
            if subject:
                subject_text = decode_header(subject)
                print 'Message: {}'.format(reconstruct_subject(subject_text))
            else:
                print 'Message: -'

            for part in mail.walk():

                if part.get_content_maintype() == 'multipart':
                    continue
                if noinline:
                    content = part.get_all('Content-Disposition')
                    if content is None:
                        continue
                    else:
                        content = content[0].split(';')
                        if content[0] == 'inline':
                            continue

                filename = part.get_filename()

                if filename is None:
                    continue
                else:
                    attachment_present = True

                    # Filename can be UTF8, quoted-printable.
                    name, charset = decode_header(filename)[0]
                    if charset:
                        filename = name.decode(charset)

                    file_path = os.path.join(directory, filename)
                    file_path = generate_unique_filename(directory, file_path)

                    print 'Saving file to: {}'.format(file_path.encode('utf-8'))

                    try:
                        output_file = open(file_path, 'wb')
                        output_file.write(part.get_payload(decode=True))
                    except IOError as err:
                        print 'Cannot write file: {}'.format(err)
                        server.close()
                        server.logout()
                        sys.exit(1)
                    else:
                        output_file.close()

                    count_att += 1

            if not attachment_present:
                print 'No attachment'

            # Delete message.
            if delete:
                server.uid('store', mail_id, '+FLAGS', '(\\Deleted)')
                server.expunge()

            count_mail += 1

        end = strftime('%H:%M:%S', gmtime((time() - start)))

        print '- Finished downloading'
        print '----------------------------------------'
        print 'Summary'
        print '----------------------------------------'
        print 'Number of messages processed: {:>10}'.format(count_mail)
        print 'Number of attachments: {:>17}'.format(count_att)
        print 'Time taken: {:>28}'.format(end)

        server.close()
        server.logout()
    else:
        print 'Cannot select given mailbox: {}'.format(mbox)
        server.logout()
        sys.exit(1)

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

    return complete_subject.encode('utf-8')

def usage():
    '''
    Print usage.
    '''

    print '\ndatts - download and save attachments from IMAP server'
    print '\nUsage: ', sys.argv[0],
    print '''--login --mbox --dir --server [--n] [--delete] [--dump] [--noinline] [--help]

    Option      Argument    Description
    -------------------------------------------
    --help                  show this help  
    --login     string      login to your account
    --server    string      server name
    --mbox      string      remote folder with attachments
    --dir       string      local folder for storing attachments
    --n         number      how many messages to download? Default is all of them.
    --delete                should we delete message after download? Default is to NOT delete.
    --dump                  print provided options and exit
    --noinline              skip attachments embedded in message body text
    '''

def banner():
    '''
    Print nice banner.
    '''

    # ASCII art generated using `Text to ASCII Art Generator` at http://patorjk.com

    print r'''

      ____/ /___ _/ /_/ /______
     / __  / __ `/ __/ __/ ___/
    / /_/ / /_/ / /_/ /_(__  ) 
    \__,_/\__,_/\__/\__/____/  v0.1.0
    '''

if __name__ == '__main__':

    main()
