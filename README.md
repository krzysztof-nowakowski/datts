	       __      __  __      
	  ____/ /___ _/ /_/ /______
	 / __  / __ `/ __/ __/ ___/
	/ /_/ / /_/ / /_/ /_(__  ) 
	\__,_/\__,_/\__/\__/____/
Download and save attachments from IMAP server

## Info

This simple tool lets you download and save attachments from your e-mails. It should work with any mail service providers that support IMAP protocol, including free gmail.com, outlook.com, yahoo.com etc.
It supports simultaneous connections and you can interrupt download process by pressing Ctrl+C - it will exit gracefully.

## How to install

```sh
 git clone https://github.com/krzysztof-nowakowski/datts.git
```
or hit Download button

## How to use

```sh
./datts.py --help

Usage:  ./datts.py --login --mbox --dir --server [--n] [--c] [--delete] [--dump] [--noinline] [--help]

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

```

## Examples

1. You want to download all attachments from fun_stuff folder to your local folder called backup/. Your e-mail service provider
is gmail.com.

```sh
./datts.py --login name --server imap.gmail.com --mbox fun_stuff --dir backup/ 
```
2. You want to process 100 messages from default folder(inbox) and save attachments to backup/ and delete those 100 messages after that. Your e-mail provider is outlook.com.

```sh
./datts.py --login name --server outlook.office365.com --mbox inbox --dir backup/ --n 100 --delete
```
3. You have lots of files and want to speed up downloading process by starting 5 simultaneous connections

```sh
./datts.py --login name --server imap.gmail.com --mbox fun_stuff --dir backup/ --c 5
```

## Q&A

#### Q: Where do I get all those IMAP server names?

A: From your mail service provider, also check your account settings or maybe on one of this sites:
* https://www.arclab.com/en/kb/email/list-of-smtp-and-imap-servers-mailserver-list.html
* https://www.cubexsoft.com/help/imap-server.html
* take a look at providers.txt

#### Q: Will datts delete my messages in mailbox?

A: Default is to leave them in mailbox unless you specify `--delete` option. Keep in mind that this is permanent deletion, you won't find deleted emails in your Trash/Bin folder.

#### Q: How about data being overwritten?

A:  If you don't download all of the attachments at first run and run datts again, it will download the same messages next time and write them to disk with slighty changed names. This is to avoid data overwriting as you may have more than one attachmet named the same name.

#### Q: What does the `--noinline` option do?

A: When you create email and insert for eg. image in the text, it will become "inline attachment". Sometimes you don't want to download this type of files as often they are logos, banners or similiar garbage.

## Note: 
- this tool assumes the default port 993, if your is different you need to change `SERVER_PORT =  993` in the code
to whatever is needed.

## TODO:

- [ ] move to Python 3

