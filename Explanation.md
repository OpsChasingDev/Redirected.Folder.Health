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
The script currently operates by getting all enabled user account info in the domain and matching the results of the child items in C:\Users to find the SIDs to check on each computer.  However, it might be possible on each machine to simply get the child items of the HKEY_USERS hive and scan all of them for redirections.  In order for this to work, we need to find answers to the following questions:
- Does "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" actually match all SIDs with user accounts for the machine?
- Does "HKU\" load all information about a user account on the machine when accessed from another session?
- Does "HKLM\SYSTEM\CurrentControlSet\Control\hivelist\" show only logged in users? (we have to only check logged in users)