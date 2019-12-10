#!/usr/bin/python

'''
	datts - download and save attachments from IMAP server.

	For help type: ./datts --help
'''

import os
import sys
import socket
from time import time, gmtime, strftime
from getopt import getopt, GetoptError
import imaplib
from email import message_from_string
from email.header import decode_header


''' defult port for IMAP over SSL '''
SERVER_PORT =  993


def main():

	delete = False
	dump = False
	how_many = None
	noinline = False

	if len(sys.argv) < 2:
		usage()
		sys.exit(0)

	''' True if option is required '''
	options = [
			['login=', True],
			['password=',True],
			['server=',True],
			['dir=',True],
			['mbox=',True],
			['n=',False],
			['delete',False],
			['help',False],
			['dump',False],
			['noinline',False]
	]

	try:
		opts, args = getopt(sys.argv[1:], '', [options[x][0] for x in range(len(options))])
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

	for o,v in opts:
		if o == '--login':
			username = v 
		if o == '--password':
			password = v
		if o == '--server':
			server_name = v
		if o == '--dir':
			directory = v
		if o == '--mbox':
			mbox = v
		if o == '--n':
			how_many = v
		if o == '--delete':
			delete = True
		if o == '--dump':
			dump = True
		if o == '--noinline':
			noinline = True

	if dump:
		print '\nOptions dump:\n'
		for o,v in opts:
			if o == '--dump':
				continue
			else:
				print '{:10} = {}'.format(o,v)
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
	
	banner()
	
	start = time()
	print '- Starting...'

	try:
		server = imaplib.IMAP4_SSL(server_name, SERVER_PORT)
		server.login(username, password)
	except socket.error:
		print 'Unable to connect to: {}'.format(server_name)
		sys.exit(1)
	except imaplib.IMAP4.error as err:
		print 'Unable to login: {}'.format(err)
		sys.exit(1)

	print '- Connected to server'

	select_return_code, msg_count = server.select(mbox)
	if select_return_code == 'OK':

		uids_return_code, uids_data = server.uid('search', None, 'ALL')
		uids = uids_data[0].split()

		if not uids:
			print 'No messages found in: {}'.format(mbox)
			server.close()
			server.logout()
			sys.exit(0)

		'''
		SORT command may not be available in all IMAP servers. rfc3501, rfc5256
		Instead we are sorting UIDs locally.
		'''
		uids = [int(x) for x in uids]
		uids.sort(reverse=True)

		print '- Number of messages in {} mailbox: {}'.format(mbox,msg_count[0])

		''' stats '''
		count_mail = 0
		count_att = 0

		print '- Starting download:'

		for mail_id in uids[:how_many]:

			attachment_present = False

			print '- Message #: {}'.format(mail_id)
			mail_return_code, mail_data = server.uid('fetch', mail_id, '(RFC822)')
			mail = message_from_string(mail_data[0][1])

			for part in mail.walk():
				if part.get_content_maintype() == 'multipart':
					continue

				if noinline:
					content = part.get_all('Content-Disposition')
					if content is None:
						pass
					else:
						content = content[0].split(';')
						if content[0] == 'inline':
							continue

				filename = part.get_filename()

				if filename is None:
					continue
				else:
					attachment_present = True

					''' filename can be UTF8, quoted-printable '''
					name, charset = decode_header(filename)[0]
					if charset:
						filename = name.decode(charset)

					file_path = os.path.join(directory, filename)

					''' avoid overwriting files by adding sequence number '''
					sequence = 1
					while os.path.isfile(file_path):
						new_filename = os.path.basename(file_path).split('.')[0] + '(' + str(sequence) + ')'
						new_filename = new_filename + '.' + '.'.join(os.path.basename(file_path).split('.')[1:])
						file_path = os.path.join(directory, new_filename)
						sequence += 1
					
					print 'Saving file to: {}'.format(file_path.encode('utf-8'))

					try:
						fd = open(file_path, 'wb')
						fd.write(part.get_payload(decode=True))
					except IOError as err:
						print 'Cannot write file: {}'.format(err)
						server.close()
						server.logout()
						sys.exit(1)
					else:
						fd.close()
				
					count_att += 1

			if not attachment_present:
				print 'No attachment'

			''' delete message '''
			if delete:
				server.uid('store', mail_id, '+FLAGS', '\\Deleted')
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
		
def usage():

	print '\ndatts - download and save attachments from IMAP server'
	print '\nUsage: ',  sys.argv[0], '''--login --password --mbox --dir --server [--n] [--delete] [--dump] [--noinline] [--help]

	Option		Argument	Description
	-------------------------------------------
	--help				show this help	
	--login		string		login to your account
	--password	string		password to your account
	--server	string		server name
	--mbox		string		remote folder with attachments
	--dir		string		local folder for storing attachments
	--n		number		how many messages to download? Default is all of them.
	--delete			should we delete message after download? Default is to NOT delete.
	--dump				print provided options and exit
	--noinline			skip attachments embedded in message body text

	'''

def banner():

	''' ASCII art generated using `Text to ASCII Art Generator` at http://patorjk.com '''

	print '''
	       __      __  __      
	  ____/ /___ _/ /_/ /______
	 / __  / __ `/ __/ __/ ___/
	/ /_/ / /_/ / /_/ /_(__  ) 
	\__,_/\__,_/\__/\__/____/  v.0.1   
	'''

if __name__ == '__main__':
	main()

