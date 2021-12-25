# Redirected.Folder.Health
Initiative for automating the monitoring and reporting of folder redirections.

Credit towards Dylan Brandes for planting the seed for this idea.  Thanks, bud!

Prior to starting this project in GitHub, this tool was already in production at my place of work for several of our clients and has been able to report on an unexpectedly large number of broken folder redirections that we would've otherwise not been able to catch, helping to both maintain user functionality and act as a line of defense against data loss.

My goal with this project is to refine the tool with a full re-work of the code and the breakdown of the components invovled so they will be more reusable, more easily readable, more easily debuggable, and ultimately be something that will not only be a full start-to-finish solution, but also be easy to give other members of the Team so they can reap the benefits of its power.

Publicly storing my code is also something I have found is important to me so that I can be exposed to criticism and feedback.

Future goals for the project that make this different from the original (only ordered by how they came to my mind):

	1) (DONE) Reduce time taken for RFH to run
		- (DONE) only fetch SIDs once
		- (DONE) utilizing PSSessions for Invoke-Command calls
		- (DONE) write results to custom objects that can be used in the pipeline
2) Find the redirection GPO by actual settings, not matching string values with wildcards
3) Scheduled health checks using PowerShell Scheduled Jobs, not Scheduled Tasks
	4) (DONE) Add support for OneDrive redirections
	5) (DONE) Enhanced flexibility and options for the email notification such as authentication types and the use of SSL
6) New function to check the configuration of the monitoring system
7) Cleaned up version of Move-Module that has options to move the module file to any user-chosen PSModule path
	8) (DONE) Added functionality to check more user library locations for redirections such as appdata
	9) (DONE) Added messaging for running as Verbose in RFH
10) New function for actively changing redirection settings in the user's registry
	11) (DONE) Allow multiple accounts to be exluded
12) Change function name of Schedule-RedirectedFolderHealth to Set-RFHSchedule
	13) (DONE) Change parameter ComputerName to use the ValidateNotNullOrEmpty instead of Mandatory and then set the default value of ComputerName to $env:ComputerName
	14) (DONE) Add parameter validation and defaults to other parameters so that defaults are accepted that can be run on the localhost as a quick local tool
	15) (DONE) Add ValidateScript parameter option to the regular expression defining an email address so that a meaningful error message is spit out instead of the 	regular expression's syntax which nobody will know how to read
	** OR **
	16) (DONE) Just set the data type of the email address parameters to [mailaddress]
17) New function to estimate the total size broken down by library that a machine will take up on a server once redirection is enabled
	18) (DONE (not doing this)) Remove in-line comments
	19) (DONE) Look at replacing parameter data types like [string] into [System.IO.FileInfo] to further restrict user input options
20) Find a more efficient way of getting the domain info for the email report to replace the use of Get-ComputerInfo; this only adds a few seconds, but cmon -_-
21) Add a progress bar?

** Next big change will be another re-work to run the check on all provided machines IN PARALLEL **
