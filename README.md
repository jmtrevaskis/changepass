Simple VB Script tool to change local passwords of Windows servers,

Will work fine on Windows XP/7/2003/2008/2012

Syntax: cscript randompass.vbs computers.txt

The script will pick a random 20 character password and output to passwords.csv in CSV format. Makes use of the pstools pspasswd.exe because I am lazy. Put your list of computers in computers.txt and when prompted enter an administrator username/password. 