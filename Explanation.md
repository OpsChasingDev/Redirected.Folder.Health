# Plain explanation of the RFH script

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
