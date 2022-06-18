# Explanation of the Get-RFH Function

## Begin
- Collecting AD User information in a custom object
- Preparing output formating
- Configuring the array of libraries to check
- Configuring the array of results from checks

## Process
- Handles the function's main logic
- Creates PSSessions to each computer
- Gets the list of users potentially logged in based on the folder names in C:\Users\
- Compares list of users with AD collection to return SamAccountNames and SIDs
- Gets the value of each library path specified for each user
- About 80% of code in the blocks for each library is used for providing information to $ShowHost (not necessary)
- Creates the custom object used for output collection
- Closes the PSSessions to each computer

## End
- Responsible for delivering selective output based on what the user specified
- Also handles email sending
- Also handles visually segregated results on screen

# Running in Parallel

## Potential Problems
- Method 1: Sending bulk information to each computer simultaneously (ADUsers and libraries to check)
- Method 2: Letting each remote machine query AD for users simultaneously
- Method 3: Find a way to use PowerShell to get logged in users in a deterministic fashion

## Potential Solution
The script currently operates by getting all enabled user account info in the domain and matching the results of the child items in C:\Users to find the SIDs to check on each computer.  However, it might be possible on each machine to simply get the child items of the HKEY_USERS hive and scan all of them for redirections.  We can bypass the need to make any calls to AD or send bulk user information to remote machines.  In order for this to work, we need to find answers to the following questions:
- Does "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" actually match all SIDs with user accounts for the machine? **YES**
- Does "HKU\" load all information about all user accounts on the machine when accessed from another session? **NO** (only loads info about logged in users)
- Does "HKLM\SYSTEM\CurrentControlSet\Control\hivelist\" show only logged in users? (we have to only check logged in users) **YES**

## Plan Notes for Parallel Operation
- still only check logged in users (checking any other users on the system may be wasteful as there is no way to know if those user accounts still belong to people actively employed at the organization)
- only input to remote machines will be what libraries to check
- only output from remote machines will be a custom PS object with the user/computer/library information
- host will initiate remote code and gather all the output (collection of custom PS objects)
- host will then do all logic necessary to return the user desired output
- create a wrapper for Invoke-Command so users can set a throttle limit on the concurrency
- get rid of Write-Host options and code entirely