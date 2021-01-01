Reason why i developed this tool.

There might be several reasons where you end up with duplicate emails in outlook. in my case I migrated emails from a "windows live mail" via export/import to outlook and ended up with duplicate emails.
Outlook has a build in solution to remove duplicate emails but it doesn't seem to work as expected as many duplicate emails remained in the mailbox and more people seem to have the same problem without a solution.
After googleing for a solution I downloaded a trial of a commercial product but only 10 duplicate emails could be removed and after installation I received a spear phishing email, so no go for that solution.
I decided to create my own solution. See release section for the compiled executable.

What it does.

This tool will move all duplicate emails of your outlook instance to the "deleted items folder". 

How it does this.

The tool will connect to outlook and calculate a hash of the body of all emails and save it in a local db.
Some automated e-mails have the same body but diferent attachments. Please note, they will be compared as equal if you dont including the "sent date" or the "size" of the e-mail as a comparing criteria. 
Therefore it is recommended to keep the default setting which includes the "sent date", when running the tool.
Ofcourse the usage of this tool is at your own risk.

How to use it.

Download the source code and compile it in VS 2019 OR copy OutlookDupMails.exe and the dupmails.mdb file from the 'release' section to the same folder on your computer. 
Hit the start button. 
note: Do not close outlook while the tool is running as it might interrupt it's processsing.
