# Redirected Folder Health

> _"You should PowerShell a way to find broken folder redirections..."_

![Redirection.Check.png](https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/Redirection.Check.png)

## <img src="https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/shell.prompt.icon2.png" width="25"/> About
We all love the functionality and idea behind folder redirections on Windows platforms.  The idea of user data on servers is an attractive one, providing both security and backup protection.  At the same time, folder redirection works behind the scenes, allowing users to work on and save data in places they are already familiar with using.

## <img src="https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/shell.prompt.icon2.png" width="25"/> The Problem
As good as folder redirections are, this Group Policy based solution comes with absolutely no system of alerting when things go wrong, leaving admins blind to issues where user data is no longer being redirected.  The issue can be a time bomb, silently waiting until the day comes where the user's device fails, a file gets corrupted, or some other type of data integrity event occurs.  As a result, data loss events occur out of left field and can be catastrophic in certain situations.

## <img src="https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/shell.prompt.icon2.png" width="25"/> A Solution
![PowerShell Icon](https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/PowerShell_Core_6.0_icon.png)
The initial design of the solution was specific to our work case at the time, but I have since published this code base as a means of controlling a far more re-useable version.  The Get-RFH.ps1 script can be used in a standalone method for quick, one-off checks for individual or bulk machines, or it can be called on in constructors for more versatility such as automating a regular event where Get-RedirectedFolderGPO updates which libraries are configured to be redirected and passes that information to the Get-RFH script where the results are emailed to a ticketing solution for an engineer to investigate further.

## <img src="https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/shell.prompt.icon2.png" width="25"/> Get-RFH.ps1 Examples
Taken directly from the comment-based help:

- Returns the Desktop, Documents, and Downloads path for all users logged into the specified machine, SL-COMPUTER-001.  The function is piped to Select-Object so that only the desired object members are shown.

#### Input
```
Get-RFH -ComputerName 'SL-COMPUTER-001' -Library D,O,W | Select-Object ComputerName,User,Desktop,Documents,Downloads
```

#### Output
```
ComputerName : SL-COMPUTER-001
User         : Administrator
Desktop      : C:\Users\Administrator\Desktop
Documents    : C:\Users\Administrator\Documents
Downloads    : C:\Users\Administrator\Downloads

ComputerName : SL-COMPUTER-001
User         : user1
Desktop      : \\SL-DC-01\RedirectedFolders\user1\Desktop
Documents    : \\SL-DC-01\RedirectedFolders\user1\Documents
Downloads    : \\SL-DC-01\RedirectedFolders\user1\Downloads
```

- Gets the content of a text file for a list of computer names and runs the cmdlet against those machines.  The Administrator account is excluded from the gathered information by using the -ExcludeAccount parameter.

#### Input
```
Get-RFH -ComputerName (Get-Content C:\test\computers.txt) -Library D,O,P -ExcludeAccount Administrator | Select-Object ComputerName,User,Desktop,Documents,Pictures
```

#### Output
```
ComputerName : sl-computer-001
User         : user1
Desktop      : \\SL-DC-01\RedirectedFolders\user1\Desktop
Documents    : \\SL-DC-01\RedirectedFolders\user1\Documents
Pictures     : \\SL-DC-01\RedirectedFolders\user1\Pictures

ComputerName : sl-computer-002
User         : User-002
Desktop      : C:\Users\user-002\Desktop
Documents    : C:\Users\user-002\Documents
Pictures     : C:\Users\user-002\Pictures
```

- Runs the check using the -ShowHost parameter.  Information about the script's operation and findings are displayed color-coded on the console, and the script returns the object containing its findings at the end.

#### Input
```
Get-RFH -ComputerName 'SL-COMPUTER-001','SL-COMPUTER-002' -Library D,O -ShowHost | Select-Object ComputerName,User,Desktop,Documents
```

#### Output
```
RFH script started on 12/25/2021 10:52:46.
 Library(ies) being checked:
    Desktop
    Documents
Checking for redirections loaded on SL-COMPUTER-001...
   Desktop path for user Administrator on machine SL-COMPUTER-001 is not redirected!
   Documents path for user Administrator on machine SL-COMPUTER-001 is not redirected!
   Desktop path for user user1 on machine SL-COMPUTER-001 is redirected.
   Documents path for user user1 on machine SL-COMPUTER-001 is redirected.
Checking for redirections loaded on SL-COMPUTER-002...
   Desktop path for user Administrator on machine SL-COMPUTER-002 is not redirected!
   Documents path for user Administrator on machine SL-COMPUTER-002 is not redirected!
   Desktop path for user User-002 on machine SL-COMPUTER-002 is not redirected!
   Documents path for user User-002 on machine SL-COMPUTER-002 is not redirected!
RFH script completed on 12/25/2021 10:52:49 after 0 hour(s), 0 minute(s), and 2 second(s) for Library(ies) Desktop Documents
ComputerName    User          Desktop                                    Documents
------------    ----          -------                                    ---------
SL-COMPUTER-001 Administrator C:\Users\Administrator\Desktop             C:\Users\Administrator\Documents
SL-COMPUTER-001 user1         \\SL-DC-01\RedirectedFolders\user1\Desktop \\SL-DC-01\RedirectedFolders\user1\Documents
SL-COMPUTER-002 Administrator C:\Users\Administrator\Desktop             C:\Users\Administrator\Documents
SL-COMPUTER-002 User-002      C:\Users\user-002\Desktop                  C:\Users\user-002\Documents
```

-  The first part of the commnand gets the string formatted names of all computers located in a specific OU.  These names are sent down the pipeline where they are used as input for the -ComputerName parameter of Get-RFH.  Get-RFH checks each of those computers for the Desktop, Documents, Music, Pictures, and Video library paths of any logged in user and writes only results where a path is not redirected to a log file.  The log file is then sent as an attachment in an email report and then deleted from disk.

#### Input
```
(Get-ADComputer -Filter * -SearchBase 'OU=SL_Computers,OU=SavyLabs,DC=savylabs,DC=local').Name | Get-RFH -Library D,O,M,P,V -LogError C:\test\errors.csv -SendEmail user@mycompany.com -From no-reply@mycompany.com -SmtpServer mail.mycompany.com -Port 25
```

- Checks a number of libraries for users on an RDS server to look for redirection problems and send them to a supervisor for review.

#### Input
```
Get-RFH -ComputerName SL-RDS-03 -Library D,O,M,P,V,A,W -LogError $env:TEMP\temp.csv -SendEmail boss@mycompany.com -Cc me@mycompany.com -From no-reply@mycompany.com -SmtpServer mail.mycompany.com -Port 25
```

## <img src="https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/shell.prompt.icon2.png" width="25"/> Known Issues and Limitations
- Target machines are checked in serial
- System initiating the operation must have the ActiveDirectory PS module
- Remote PS connections must be allowed on client endpoints
- Only logged in user sessions will be detected (users who are logged in but have their screen locked, are idle, are inactive, or are disconneted (in the case of a connection to a remote desktop) **_will_** be detected)

## <img src="https://raw.githubusercontent.com/drummermanrob20/Misc/main/resources/shell.prompt.icon2.png" width="25"/> Future Goals
As a means of self-criticism, the biggest thing that bothers me about the way my tool works is that it operates in serial.  If you need to check a few hundred machines, you're going to be waiting for each of them to be checked one at a time.  While this methodology may save on bandwidth and not matter for use in an automated check, the solution currently fails to utilize one of the best aspects about PowerShell and therefore creates limitations when scaling.  Leveraging more resouces on the client end and getting the script to work in parallel is ultimately what needs to be done next.

> **Other future goals for the project included below:**
- Add a progress bar (though moot if the operation can run in parallel)
- New function for estimating the total size redirections will consume on a server once implemented
- New function for actively changing broken redirections when found