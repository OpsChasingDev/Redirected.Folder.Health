Function Get-RFH {
    <#
    .SYNOPSIS
        Runs against a computer to check the path of specified user libraries for any user logged into that machine.
    .DESCRIPTION
        Runs against a computer to check the path of specified user libraries for any user logged into that machine.
        
        This function is written to fulfill multiple use cases.  As a tool, it can be used to remotely and quickly determine the path of a user library without the need for registry browsing or even knowing who is logged into the target system.
        
        As an automated solution, the cmdlet can be set up to regularly investigate systems and libraries to report on problems found where libraries that should be redirected are in fact not so remediation can take place, helping to avoid data loss and problems with data accessibility.
        
        Problem detection currently supports traditional folder redirections to a UNC path as well as OneDrive redirections.  Any local path for a specified user library will be identified as a problem.

        The function writes entries to the Application log when an operation is started (Informational), an operation has completed (Informational), and a check mid-operation has detected a library not redirected (Warning).  Event ID: 13, Source: RedirectedFolderHealth

        Restrictions and Limitations:
        - The system running the function must have access to the ActiveDirectory module.
        - PowerShell remoting must be enabled on target machines.
        - A user session must be logged in (either active or innactive) for data to be gathered on that user's library paths.
        - The function currently works in serial.
    .EXAMPLE
        PS C:\> Get-RFH -Library D

        Output:
ComputerName : SL-DC-01
User         : Administrator
Desktop      : C:\Users\Administrator\Desktop
Documents    :
Downloads    :
Music        :
Pictures     :
Video        :
Favorites    :
AppData      :
StartMenu    :
Contacts     :
Links        :
Searches     :
SavedGames   :

        Returns the Desktop path for all users logged into the local machine.  The object returned includes libraries that were not checked and leaves them as null.
    .EXAMPLE
        PS C:\> Get-RFH -ComputerName 'SL-COMPUTER-001' -Library D,O,W | Select-Object ComputerName,User,Desktop,Documents,Downloads

        Output:
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

        Returns the Desktop, Documents, and Downloads path for all users logged into the specified machine, SL-COMPUTER-001.  The function is piped to Select-Object so that only the desired object members are shown.
    .EXAMPLE
        PS C:\> Get-RFH -ComputerName (Get-Content C:\test\computers.txt) -Library D,O,P -ExcludeAccount Administrator | Select-Object ComputerName,User,Desktop,Documents,Pictures

        Output:
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

        Gets the content of a text file for a list of computer names and runs the cmdlet against those machines.  The Administrator account is excluded from the gathered information by using the -ExcludeAccount parameter.
    .EXAMPLE
        PS C:\> (Get-ADComputer -Filter * -SearchBase 'OU=SL_Computers,OU=SavyLabs,DC=savylabs,DC=local').Name | Get-RFH -Library D,V,A -ExcludeAccount Administrator

        Output:
Warning: The computer SL-COMPUTER-003 could not be contacted!

ComputerName : SL-COMPUTER-001
User         : user1
Desktop      : \\SL-DC-01\RedirectedFolders\user1\Desktop
Documents    :
Downloads    :
Music        :
Pictures     :
Video        : \\SL-DC-01\RedirectedFolders\user1\Videos
Favorites    :
AppData      : C:\Users\user1\AppData\Roaming
StartMenu    :
Contacts     :
Links        :
Searches     :
SavedGames   :

ComputerName : SL-COMPUTER-002
User         : User-002
Desktop      : C:\Users\user-002\Desktop
Documents    :
Downloads    :
Music        :
Pictures     :
Video        : C:\Users\user-002\Videos
Favorites    :
AppData      : C:\Users\user-002\AppData\Roaming
StartMenu    :
Contacts     :
Links        :
Searches     :
SavedGames   :

        Checks the Desktop, Video, and AppData libraries for all users on a list of computers obtained by piping Get-ADComputer to Get-RFH.  The Administrator account is excluded using the -ExcludeAccount parameter.  A warning message is presented for SL-COMPUTER-003 as it could not be contacted.
    .EXAMPLE
        PS C:\> Get-RFH -ComputerName 'SL-COMPUTER-001','SL-COMPUTER-002' -Library D,O -ShowHost | Select-Object ComputerName,User,Desktop,Documents

        Output:
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

        Runs the check using the -ShowHost parameter.  Information about the script's operation and findings are displayed color-coded on the console, and the script returns the object containing its findings at the end.
    .EXAMPLE
        PS C:\> Get-RFH -ComputerName 'SL-COMPUTER-001','SL-COMPUTER-002' -Library D | Where-Object {$_.Desktop -match "SL-DC-01"} | Select-Object User,Desktop

        Output:
User  Desktop
----  -------
user1 \\SL-DC-01\RedirectedFolders\user1\Desktop

        Checks specified computers to see if the Desktop path for the users logged into those machines is redirected to the server SL-DC-01.  The function first gathers the Desktop value for all user sessions found on the specified machines.  The returned information is piped to Where-Object which only passes the objects where the Desktop property matched "SL-DC-01".
    .EXAMPLE
        PS C:\> Get-RFH -ComputerName 'SL-COMPUTER-001','SL-COMPUTER-002' -Library D | Where-Object {$_.Desktop -notlike "*\\*"} | Select-Object User,Desktop

        Output:
User          Desktop
----          -------
Administrator C:\Users\Administrator\Desktop
Administrator C:\Users\Administrator\Desktop
User-002      C:\Users\user-002\Desktop

        Checks specified computers to see if any logged in users have their Desktop path pointed to a non-UNC location.
    .EXAMPLE
        PS C:\Windows\System32> (Get-ADComputer -Filter * -SearchBase 'OU=SL_Computers,OU=SavyLabs,DC=savylabs,DC=local').Name | Get-RFH -Library D,O,M,P,V -LogAll C:\test\logall.csv

        Output: Writes all findings to the CSV file specified in -LogAll.  Any machines that could not be contacted for a check will be displayed on the console as a warning message.
    .EXAMPLE
        PS C:\Windows\System32> (Get-ADComputer -Filter * -SearchBase 'OU=SL_Computers,OU=SavyLabs,DC=savylabs,DC=local').Name | Get-RFH -Library D,O,M,P,V -LogError C:\test\errors.csv

        Output: Writes findings where paths were not redirected to the CSV file specified in -LogError.  Any machines that could not be contacted for a check will be displayed on the console as a warning message.
    .EXAMPLE
        Get-RFH -ComputerName (Get-Content C:\test\computers.txt) -Library D,O,M,P,V -LogError C:\test\errors.csv -LogAll \\SL-DC-01\TestShares\log.csv

        The cmdlet checks the Desktop, Documents, Music, Pictures, and Video library paths for any user logged into the machines listed in the computers.txt file.  Only users who have one or more of those paths without redirection will be logged to the errors.csv file, whereas all findings regardles of path value will be stored in the log.csv file saved to a network share.
    .EXAMPLE
        PS C:\Windows\System32> (Get-ADComputer -Filter * -SearchBase 'OU=SL_Computers,OU=SavyLabs,DC=savylabs,DC=local').Name | Get-RFH -Library D,O,M,P,V -LogError C:\test\errors.csv -SendEmail user@mycompany.com -From no-reply@mycompany.com -SmtpServer mail.mycompany.com -Port 25
        
        The first part of the commnand gets the string formatted names of all computers located in a specific OU.  These names are sent down the pipeline where they are used as input for the -ComputerName parameter of Get-RFH.  Get-RFH checks each of those computers for the Desktop, Documents, Music, Pictures, and Video library paths of any logged in user and writes only results where a path is not redirected to a log file.  The log file is then sent as an attachment in an email report and then deleted from disk.
    .EXAMPLE
        PS C:\Windows\System32> (Get-ADComputer -Filter * -SearchBase 'OU=SL_Computers,OU=SavyLabs,DC=savylabs,DC=local').Name | Get-RFH -Library D,O,M,P,V -LogError C:\test\errors.csv -SendEmail user@mycompany.com -Bcc reportcollection@mycompany.com -From no-reply@mycompany.com -SmtpServer mail.mycompany.com -Port 2525 -UseSSL
        Performs the same functionality as the previous example but also blind copies the report to another recipient, uses a custom SMTP server port, and elects to use SSL.
    .EXAMPLE
        Get-RFH -ComputerName SL-RDS-03 -Library D,O,M,P,V,A,W -LogError $env:TEMP\temp.csv -SendEmail boss@mycompany.com -Cc me@mycompany.com -From no-reply@mycompany.com -SmtpServer mail.mycompany.com -Port 25
        Checks a number of libraries for users on an RDS server to look for redirection problems and send them to a supervisor for review.
    .INPUTS
        System.String[]
        This cmdlet accepts computer names as strings in order to specify the systems it runs on.
    .OUTPUTS
        CSV
        If using either the -LogAll or -LogError parameters, output will be stored in a CSV file.

        Email
        If using the -SendEmail parameter, output stored in a CSV file will be sent as an attachment on an email.

        System.Management.Automation.PSCustomObject
        By default, a PSCustomObject is returned by the cmdlet with the below members:

Name         MemberType   Definition
----         ----------   ----------
Equals       Method       bool Equals(System.Object obj)
GetHashCode  Method       int GetHashCode()
GetType      Method       type GetType()
ToString     Method       string ToString()
AppData      NoteProperty System.String
ComputerName NoteProperty System.String
Contacts     NoteProperty System.String
Desktop      NoteProperty System.String
Documents    NoteProperty System.String
Downloads    NoteProperty System.String
Favorites    NoteProperty System.String
Links        NoteProperty System.String
Music        NoteProperty System.String
Pictures     NoteProperty System.String
SavedGames   NoteProperty System.String
Searches     NoteProperty System.String
StartMenu    NoteProperty System.String
User         NoteProperty System.String
Video        NoteProperty System.String
    .NOTES
        Author: Robert Stapleton
        Version: 2.0.0
        Date: 2022-01-11
    .LINK
        https://github.com/drummermanrob20/Redirected.Folder.Health
    #>
    [Cmdletbinding(DefaultParameterSetName = 'General')]
    [Alias('Get-RedirectedFolderHealth')]
    Param (

        <# Specifies the computer on which you want to run the folder redirection check.
        If nothing is specified, the cmdlet runs against the localhost returned by $env:COMPUTERNAME #>
        [Parameter(ParameterSetName = 'General',
                    ValueFromPipeline = $true,
                    Position = 0)]
        [Parameter(ParameterSetName = 'Email-LogAll',
                    ValueFromPipeline = $true,
                    Position = 0)]
        [Parameter(ParameterSetName = 'Email-LogError',
                    ValueFromPipeline = $true,
                    Position = 0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $ComputerName = $env:COMPUTERNAME,

        <# Indicates the user libraries you wish to have inspected, denoted by a single letter per library specified.
        User libraries are represented by a letter specified in the legend below:
        D = Desktop
        O = Documents
        W = Downloads
        M = Movies
        P = Pictures
        V = Videos
        F = Favorites
        A = AppData (roaming)
        S = Start Meu
        C = Contacts
        L = Links
        H = Searches
        G = Saved Games #>
        [Parameter(ParameterSetName = 'General',
                    Mandatory,
                    Position = 1,
                    HelpMessage = 'Enter the letter corresponding to the library you want checked:
                    D = Desktop
                    O = Documents
                    W = Downloads
                    M = Movies
                    P = Pictures
                    V = Videos
                    F = Favorites
                    A = AppData (roaming)
                    S = Start Meu
                    C = Contacts
                    L = Links
                    H = Searches
                    G = Saved Games')]
        [Parameter(ParameterSetName = 'Email-LogAll',
                    Mandatory,
                    Position = 1,
                    HelpMessage = 'Enter the letter corresponding to the library you want checked:
                    D = Desktop
                    O = Documents
                    W = Downloads
                    M = Movies
                    P = Pictures
                    V = Videos
                    F = Favorites
                    A = AppData (roaming)
                    S = Start Meu
                    C = Contacts
                    L = Links
                    H = Searches
                    G = Saved Games')]
        [Parameter(ParameterSetName = 'Email-LogError',
                    Mandatory,
                    Position = 1,
                    HelpMessage = 'Enter the letter corresponding to the library you want checked:
                    D = Desktop
                    O = Documents
                    W = Downloads
                    M = Movies
                    P = Pictures
                    V = Videos
                    F = Favorites
                    A = AppData (roaming)
                    S = Start Meu
                    C = Contacts
                    L = Links
                    H = Searches
                    G = Saved Games')]
        [ValidateSet("D","O","W","M","P","V","F","A","S","C","L","H","G")]
        [String[]]
        $Library,

        <# Specifies an account to exclude, such as an administrative account that does not use folder redirections.
        If the specified account has a logged in session on a machine being checked, this account's libraries will not be inspected, nor will they be included in any report.
        Input an account in the form of a 'username' such as JDoe or John.Doe; do not include the domain in any format such as a UPN. This will be the equivalent of the SamAccountName property for an object returned by Get-ADUser. #>
        [Parameter(ParameterSetName = 'General')]
        [Parameter(ParameterSetName = 'Email-LogAll')]
        [Parameter(ParameterSetName = 'Email-LogError')]
        [String[]]
        $ExcludeAccount,

        <# Specifies the full path including the file name of a CSV file that results will be saved to; all results will be saved whether libraries are redirected or not.
        If used with the -SendEmail parameter, the file will be sent as an attachment and then deleted from its location on disk.
        If using the -SendEmail parameter, you cannot use both this parameter and the -LogError parameter in the same syntax. #>
        [Parameter(ParameterSetName = 'General')]
        [Parameter(ParameterSetName = 'Email-LogAll',
                    Mandatory,
                    HelpMessage = 'Enter the full path including the file name of the CSV file you want generated.
                    This will be attached to the email report and include all results.')]
        [ValidatePattern('.\.csv$')]
        [String]
        $LogAll,

        <# Specifies the full path including the file name of a CSV file that results will be saved to; only findings where libraries are not redirected will be reported.
        If used with the -SendEmail parameter, the file will be sent as an attachment and then deleted from its location on disk.
        If using the -SendEmail parameter, you cannot use both this parameter and the -LogAll parameter in the same syntax. #>
        [Parameter(ParameterSetName = 'General')]
        [Parameter(ParameterSetName = 'Email-LogError',
                    Mandatory,
                    HelpMessage = 'Enter the full path including the file name of the CSV file you want generated.
                    This will be attached to the email report and include only findings where libraries are not redirected.')]
        [ValidatePattern('.\.csv$')]
        [String]
        $LogError,

        # Displays the progress and results of checks using Write-Host; this is disabled by default.  Use -Verbose for a more complete view of the operations at play.
        [Parameter(ParameterSetName = 'General')]
        [Parameter(ParameterSetName = 'Email-LogAll')]
        [Parameter(ParameterSetName = 'Email-LogError')]
        [Switch]
        $ShowHost = $false,

        <# Specifies the email address you want the report sent to.
        This parameter must be used with either the -LogAll or -LogError parameter and will follow the syntax of the Email-LogAll or Email-LogError parameter set, respectively.
        You cannot use -LogAll and -LogError in the same syntax together when using -SendEmail. #>
        [Parameter(ParameterSetName = 'Email-LogAll',
                    Mandatory,
                    HelpMessage = 'Enter the email address(es) you want receiving the email report.')]
        [Parameter(ParameterSetName = 'Email-LogError',
                    Mandatory,
                    HelpMessage = 'Enter the email address(es) you want receiving the email report.')]
        [MailAddress[]]
        $SendEmail,

        <# Specifies the email address you want the report to come from.  The email address specified will be appended to a display name and show like the below example:
        -From no-reply@mycompany.com
        "Redirected Folder Health <no-reply@mycompany.com>" #>
        [Parameter(ParameterSetName = 'Email-LogAll',
                    Mandatory,
                    HelpMessage = 'Enter the email address you want the email report to come from.')]
        [Parameter(ParameterSetName = 'Email-LogError',
                    Mandatory,
                    HelpMessage = 'Enter the email address you want the email report to come from.')]
        [MailAddress]
        $From,

        # Specifies the email addresses to which a carbon copy (CC) of the email report is sent.
        [Parameter(ParameterSetName = 'Email-LogAll')]
        [Parameter(ParameterSetName = 'Email-LogError')]
        [MailAddress[]]
        $Cc,

        # Specifies the email addresses that receive a copy of the report but are not listed as recipients of the email.
        [Parameter(ParameterSetName = 'Email-LogAll')]
        [Parameter(ParameterSetName = 'Email-LogError')]
        [MailAddress[]]
        $Bcc,

        # Specifies the name of the SMTP server that sends the email report.
        [Parameter(ParameterSetName = 'Email-LogAll',
                    Mandatory,
                    HelpMessage = 'Enter the SMTP server address by name.')]
        [Parameter(ParameterSetName = 'Email-LogError',
                    Mandatory,
                    HelpMessage = 'Enter the SMTP server address by name.')]
        [String]
        $SmtpServer,

        # Specifies the port on the SMTP server.  No default value is set.
        [Parameter(ParameterSetName = 'Email-LogAll',
                    Mandatory,
                    HelpMessage = 'Enter the port number for the receiving SMTP server.')]
        [Parameter(ParameterSetName = 'Email-LogError',
                    Mandatory,
                    HelpMessage = 'Enter the port number for the receiving SMTP server.')]
        [Int]
        $Port,

        # The Secure Sockets Layer (SSL) protocol is used to establish a secure connection to the remote computer to send mail. By default, SSL is not used.
        [Parameter(ParameterSetName = 'Email-LogAll')]
        [Parameter(ParameterSetName = 'Email-LogError')]
        [Switch]
        $UseSSL

    )
    
    BEGIN {
        Write-Verbose "         Script Started        "
        Write-Verbose "   >> BEGIN BLOCK STARTED <<   "
        
        # getting current date/time and the number of libraries being checked
        $DateStart = Get-Date
        $LCount = $Library.Count

        # collecting the full names of libraries being checked for logging and output purposes based on the input provided to the -Library parameter
        $LFullCollection = @()
        $i = 0
        Do {
            switch ($Library[$i]) {
                "D" {$LFull = "Desktop"}
                "O" {$LFull = "Documents"}
                "W" {$LFull = "Downloads"}
                "M" {$LFull = "Music"}
                "P" {$LFull = "Pictures"}
                "V" {$LFull = "Video"}
                "F" {$LFull = "Favorites"}
                "A" {$LFull = "AppData"}
                "S" {$LFull = "Start Menu"}
                "C" {$LFull = "Contacts"}
                "L" {$LFull = "Links"}
                "H" {$LFull = "Searches"}
                "G" {$LFull = "Saved Games"}
            }
            $LFullCollection += "$LFull"
            $i += 1
            }
            While ($i -lt $LCount)

        # used as a property/column selector for the CSV output generated by the -LogError parameter
        $PropertyOutput = @("ComputerName","User")
        $PropertyOutput += $LFullCollection

        # writing script beginning to the application event log
        eventcreate /ID 13 /L APPLICATION /T INFORMATION /SO RedirectedFolderHealth /D "RFH script started on $DateStart for Library(ies) $LFullCollection" > $null

        # prep work by gathering enabled AD user objects with their name and SID as well as creating the empty collection arrays
        $ADUser = Get-ADUser -Filter 'enabled -eq $true'
                Write-Verbose "Stored enabled Active Directory user objects"
        $UserCollection = @()
                Write-Verbose "Created empty array for custom user object information"
        $ResultCollection = @()
                Write-Verbose "Created empty array for custom result object information"

        ForEach ($a in $ADUser) {
            $ObjUser = New-Object -TypeName psobject
            $ObjUser | Add-Member -MemberType NoteProperty -Name "Name" -Value $a.SamAccountName
            $ObjUser | Add-Member -MemberType NoteProperty -Name "SID" -Value ($a.SID).Value

            If ($ExcludeAccount -notcontains $a.SamAccountName) {
                $UserCollection += $ObjUser
            }
        }
            Write-Verbose "Custom user object information gathered and completed"

        # gets the full name of the libraries selected to show script progress and information on the console when specifying the -ShowHost switch parameter
        If ($ShowHost) {
            Write-Host -ForegroundColor Yellow "RFH script started on $DateStart." `n "Library(ies) being checked:"
            $j = 0
            Do {
                switch ($Library[$j]) {
                    "D" {$LFull = "Desktop"}
                    "O" {$LFull = "Documents"}
                    "W" {$LFull = "Downloads"}
                    "M" {$LFull = "Music"}
                    "P" {$LFull = "Pictures"}
                    "V" {$LFull = "Videos"}
                    "F" {$LFull = "Favorites"}
                    "A" {$LFull = "AppData"}
                    "S" {$LFull = "Start Menu"}
                    "C" {$LFull = "Contacts"}
                    "L" {$LFull = "Links"}
                    "H" {$LFull = "Searches"}
                    "G" {$LFull = "Saved Games"}
                }
                Write-Host -ForegroundColor Yellow "   "$LFull
                $j += 1
            }
            While ($j -lt $LCount)
        }

        Write-Verbose "   >> BEGIN BLOCK FINISHED <<   "
    }
    PROCESS {
        ForEach ($c in $ComputerName) {
            Write-Verbose "   >> PROCESS BLOCK STARTED FOR $c <<   "

            Try {
                # creates the session to the computer
                $Session = New-PSSession -ComputerName $c -Name "PSS_$c" -ErrorAction Stop
                If ($ShowHost) {Write-Host -ForegroundColor Yellow "Checking for redirections loaded on $c..."}
                    Write-Verbose "$c -   [SESSION ESTABLISHED]"

                # gets the user accounts to check on the computer
                $LocalUserFolder = Invoke-Command -Session $Session {Get-ChildItem C:\users}
                $LocalUserName = $LocalUserFolder.Name
                    Write-Verbose "$c -   Stored users"
                
                # stores the SID of each user that was found on the machine
                $LocalUserFull = $UserCollection | Where-Object {$LocalUserName -match $_.Name} | Select-Object Name,SID
                    Write-Verbose "$c -   Obtained SID information for stored users"

                # operates the data gathering for each user that was found on the computer
                ForEach ($l in $LocalUserFull) {
                    $CurrentUserSID = $l.SID
                    $CurrentUserName = $l.Name
                        Write-Verbose "$c - Checking $CurrentUserName"
                    
                    # store the "D" desktop path value for the user
                    If ($Library -eq "D") {
                        $DesktopPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name Desktop -ErrorAction SilentlyContinue).Desktop
                        }
                        If ($DesktopPath) {Write-Verbose "$c -   Desktop value stored as $DesktopPath"}

                        # conditions for logging and reporting
                        If ($DesktopPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Desktop path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($DesktopPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Desktop path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($DesktopPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Desktop path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Desktop path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "O" documents path value for the user
                    If ($Library -eq "O") {
                        $DocumentsPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name Personal -ErrorAction SilentlyContinue).Personal
                        }
                        If ($DocumentsPath) {Write-Verbose "$c -   Documents value stored as $DocumentsPath"}

                        # conditions for logging and reporting
                        If ($DocumentsPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Documents path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($DocumentsPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Documents path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($DocumentsPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Documents path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Documents path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "W" downloads path value for the user
                    If ($Library -eq "W") {
                        $DownloadsPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{374DE290-123F-4565-9164-39C4925E467B}" -ErrorAction SilentlyContinue)."{374DE290-123F-4565-9164-39C4925E467B}"
                        }
                        If ($DownloadsPath) {Write-Verbose "$c -   Downloads value stored as $DownloadsPath"}

                        # conditions for logging and reporting
                        If ($DownloadsPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Downloads path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($DownloadsPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Downloads path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($DownloadsPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Downloads path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Downloads path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "M" music path value for the user
                    If ($Library -eq "M") {
                        $MusicPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "My Music" -ErrorAction SilentlyContinue)."My Music"
                        }
                        If ($MusicPath) {Write-Verbose "$c -   Music value stored as $MusicPath"}

                        # conditions for logging and reporting
                        If ($MusicPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Music path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($MusicPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Music path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($MusicPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Music path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Music path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "P" pictures path value for the user
                    If ($Library -eq "P") {
                        $PicturesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "My Pictures" -ErrorAction SilentlyContinue)."My Pictures"
                        }
                        If ($PicturesPath) {Write-Verbose "$c -   Pictures value stored as $PicturesPath"}

                        # conditions for logging and reporting
                        If ($PicturesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Pictures path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($PicturesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Pictures path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($PicturesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Pictures path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Pictures path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "V" video path value for the user
                    If ($Library -eq "V") {
                        $VideoPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "My Video" -ErrorAction SilentlyContinue)."My Video"
                        }
                        If ($VideoPath) {Write-Verbose "$c -   Video value stored as $VideoPath"}

                        # conditions for logging and reporting
                        If ($VideoPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Video path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($VideoPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Video path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($VideoPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Video path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Video path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "F" favorites path value for the user
                    If ($Library -eq "F") {
                        $FavoritesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name Favorites -ErrorAction SilentlyContinue).Favorites
                        }
                        If ($FavoritesPath) {Write-Verbose "$c -   Favorites value stored as $FavoritesPath"}

                        # conditions for logging and reporting
                        If ($FavoritesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Favorites path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($FavoritesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Favorites path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($FavoritesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Favorites path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Favorites path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "A" roaming appdata path value for the user
                    If ($Library -eq "A") {
                        $AppDataPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name AppData -ErrorAction SilentlyContinue).AppData
                        }
                        If ($AppDataPath) {Write-Verbose "$c -   AppData value stored as $AppDataPath"}

                        # conditions for logging and reporting
                        If ($AppDataPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   AppData path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($AppDataPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   AppData path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($AppDataPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   AppData path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "AppData path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "S" start menu path value for the user
                    If ($Library -eq "S") {
                        $StartMenuPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "Start Menu" -ErrorAction SilentlyContinue)."Start Menu"
                        }
                        If ($StartMenuPath) {Write-Verbose "$c -   StartMenu value stored as $StartMenuPath"}

                        # conditions for logging and reporting
                        If ($StartMenuPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   StartMenu path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($StartMenuPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   StartMenu path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($StartMenuPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   StartMenu path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "StartMenu path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "C" contacts path value for the user
                    If ($Library -eq "C") {
                        $ContactsPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{56784854-C6CB-462B-8169-88E350ACB882}" -ErrorAction SilentlyContinue)."{56784854-C6CB-462B-8169-88E350ACB882}"
                        }
                        If ($ContactsPath) {Write-Verbose "$c -   Contacts value stored as $ContactsPath"}

                        # conditions for logging and reporting
                        If ($ContactsPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Contacts path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($ContactsPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Contacts path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($ContactsPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Contacts path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Contacts path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "L" links path value for the user
                    If ($Library -eq "L") {
                        $LinksPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}" -ErrorAction SilentlyContinue)."{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}"
                        }
                        If ($LinksPath) {Write-Verbose "$c -   Links value stored as $LinksPath"}

                        # conditions for logging and reporting
                        If ($LinksPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Links path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($LinksPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Links path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($LinksPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Links path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Links path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "H" searches path value for the user
                    If ($Library -eq "H") {
                        $SearchesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}" -ErrorAction SilentlyContinue)."{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}"
                        }
                        If ($SearchesPath) {Write-Verbose "$c -   Searches value stored as $SearchesPath"}

                        # conditions for logging and reporting
                        If ($SearchesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Searches path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($SearchesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Searches path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($SearchesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Searches path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Searches path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "G" saved games path value for the user
                    If ($Library -eq "G") {
                        $SavedGamesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}" -ErrorAction SilentlyContinue)."{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}"
                        }
                        If ($SavedGamesPath) {Write-Verbose "$c -   SavedGames value stored as $SavedGamesPath"}

                        # conditions for logging and reporting
                        If ($SavedGamesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   SavedGames path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($SavedGamesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   SavedGames path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($SavedGamesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   SavedGames path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "SavedGames path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # creates and writes members to the object for the result of user on the computer based on selected libraries
                    $ObjResult = New-Object -TypeName psobject
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$c"
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "User" -Value "$CurrentUserName"

                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Desktop" -Value $DesktopPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Documents" -Value $DocumentsPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Downloads" -Value $DownloadsPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Music" -Value $MusicPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Pictures" -Value $PicturesPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Video" -Value $VideoPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Favorites" -Value $FavoritesPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "AppData" -Value $AppDataPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "StartMenu" -Value $StartMenuPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Contacts" -Value $ContactsPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Links" -Value $LinksPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Searches" -Value $SearchesPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "SavedGames" -Value $SavedGamesPath

                    # appends the resulting object of the computer's redirection check to the collection of results only if any of the paths contain values
                    # because an object is made for every user account found on the target machine regardless of being logged in, this prevents objects with blank path values from being returned
                    If($ObjResult.Desktop -or 
                        $ObjResult.Documents -or
                        $ObjResult.Downloads -or
                        $ObjResult.Music -or
                        $ObjResult.Pictures -or
                        $ObjResult.Video -or
                        $ObjResult.Favorites -or
                        $ObjResult.AppData -or
                        $ObjResult.StartMenu -or
                        $ObjResult.Contacts -or
                        $ObjResult.Links -or
                        $ObjResult.Searches -or
                        $ObjResult.SavedGames) 
                    {
                        $ResultCollection += $ObjResult
                        Write-Verbose "$c -   Object added to the collection array"
                    }
                }

                # removes the session to the computer
                Remove-PSSession -Session $Session
                    Write-Verbose "$c -   [SESSION REMOVED]"
            }
            Catch {
                Write-Output "Warning: The computer $c could not be contacted!"
            }

            Write-Verbose "   >> PROCESS BLOCK FINISHED FOR $c <<   "
        }
    }
    END {
        Write-Verbose "   >> END BLOCK STARTED <<   "

        # outputs the collection of results as specified, containing all computers and all users found on each computer
        If ($LogAll) {
            # writes all findings to csv
            Write-Output $ResultCollection | Select-Object $PropertyOutput | Export-Csv -Path $LogAll -NoTypeInformation
        }
        If ($LogError) {
            # writes only problems to csv
            ForEach ($r in $ResultCollection) {
                If (
                    $r.Desktop -and $r.Desktop -notlike "\\*" -and $r.Desktop -notlike "*OneDrive*" -or
                    $r.Documents -and $r.Documents -notlike "\\*" -and $r.Documents -notlike "*OneDrive*" -or
                    $r.Downloads -and $r.Downloads -notlike "\\*" -and $r.Downloads -notlike "*OneDrive*" -or
                    $r.Music -and $r.Music -notlike "\\*" -and $r.Music -notlike "*OneDrive*" -or
                    $r.Pictures -and $r.Pictures -notlike "\\*" -and $r.Pictures -notlike "*OneDrive*" -or
                    $r.Video -and $r.Video -notlike "\\*" -and $r.Video -notlike "*OneDrive*" -or
                    $r.Favorites -and $r.Favorites -notlike "\\*" -and $r.Favorites -notlike "*OneDrive*" -or
                    $r.AppData -and $r.AppData -notlike "\\*" -and $r.AppData -notlike "*OneDrive*" -or
                    $r.StartMenu -and $r.StartMenu -notlike "\\*" -and $r.StartMenu -notlike "*OneDrive*" -or
                    $r.Contacts -and $r.Contacts -notlike "\\*" -and $r.Contacts -notlike "*OneDrive*" -or
                    $r.Links -and $r.Links -notlike "\\*" -and $r.Links -notlike "*OneDrive*" -or
                    $r.Searches -and $r.Searches -notlike "\\*" -and $r.Searches -notlike "*OneDrive*" -or
                    $r.SavedGames -and $r.SavedGames -notlike "\\*" -and $r.SavedGames -notlike "*OneDrive*"
                )
                {
                    Write-Output $r | Select-Object $PropertyOutput | Export-Csv -Path $LogError -NoTypeInformation -Append
                }
            }
        }

        # obtaining info to report on elapsed time taken for the script to complete
        $DateEnd = Get-Date
        $DateDiff = $DateEnd - $DateStart
        $Hour = $DateDiff.Hours
        $Minute = $DateDiff.Minutes
        $Second = $DateDiff.Seconds
        
        # handles actions taken to send an email report if specified
        If ($SendEmail) {
                Write-Verbose "Preparing email..."
            $DomainName = (Get-ADDomain).DNSRoot
            $From = $From.ToString()
            [string]$FromAddress = "Redirected Folder Health <$From>"

            # splatting to construct parameters and values for Send-MailMessage
            $EmailSplat = @{
                To = $SendEmail
                From = $FromAddress
                Subject = "Redirected Folder Health Report for $DomainName"
                SmtpServer = $SmtpServer
                Port = $Port
            }
            If ($Cc) {$EmailSplat += @{Cc = $Cc}}
            If ($Bcc) {$EmailSplat += @{Bcc = $Bcc}}
            If ($UseSSL) {$EmailSplat += @{UseSSL = $true}}
            
            # add body and peform actions if -LogAll is specified
            If ($LogAll) {
                $Body = "See the attachment for folder redirection details.
                Script completed on $env:COMPUTERNAME at $DateEnd after $Hour hour(s), $Minute minute(s), and $Second second(s) for Library(ies) $LFullCollection."
                $EmailSplat += @{Body = $Body}
                $EmailSplat += @{Attachments = $LogAll}
                Send-MailMessage @EmailSplat -WarningAction SilentlyContinue
                    Write-Verbose "Email sent."
                Remove-Item -Path $LogAll -Force
                    Write-Verbose "Log file $LogAll removed."
            }

            # add body and perform actions if -LogError is specified
            If ($LogError) {
                $Body = "See the attachment for folder redirection details.
                The local paths shown need to be addressed so they are redirected and protected against data loss.
                Script completed on $env:COMPUTERNAME at $DateEnd after $Hour hour(s), $Minute minute(s), and $Second second(s) for Library(ies) $LFullCollection."
                $EmailSplat += @{Body = $Body}
                $EmailSplat += @{Attachments = $LogError}
                Send-MailMessage @EmailSplat -WarningAction SilentlyContinue
                    Write-Verbose "Email sent."
                Remove-Item -Path $LogError -Force
                    Write-Verbose "Log file $LogError removed."
            }
        }

        # writes all findings and full object info to the pipeline if no logging options are specified
        If (!$LogAll -and !$LogError) {
            Write-Output $ResultCollection
        }

        # writing completion to the application event log and out to host
        eventcreate /ID 13 /L APPLICATION /T INFORMATION /SO RedirectedFolderHealth /D "RFH script completed on $DateEnd after $Hour hour(s), $Minute minute(s), and $Second second(s) for Library(ies) $LFullCollection." > $null
        If ($ShowHost) {Write-Host -ForegroundColor Yellow "RFH script completed on $DateEnd after $Hour hour(s), $Minute minute(s), and $Second second(s) for Library(ies) $LFullCollection"}

        Write-Verbose "   >> END BLOCK FINISHED <<   "
        Write-Verbose "       Script Completed       "
    }
}