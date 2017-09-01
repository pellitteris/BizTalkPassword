This little tool allows you to set BizTalk send ports and receive locations passwords. You can use it to create automated scripts during deployment operations.

It is just an example that currently manages the adapter FILE, FTP, POP3 and WCF *. Of course you can access the source code in order to manage other kind of adapters.

Example:

Microsys.EAI.Framework.PasswordManager.exe -set -receive -name:MyReceiveLocation -user:John -password:@Passw0rd1

Microsys.EAI.Framework.PasswordManager.exe -set -send -name:MySendPort -user:John -password:@Passw0rd1

Parameters:

-list -application:[application name]
-get -receive -name:[receive location name]
-get -send -name:[send port name]
-set -receive -name:[receive location name] -user:[username] -password:[password]
-set -send -name:[send port name] -user:[username] -password:[password]
