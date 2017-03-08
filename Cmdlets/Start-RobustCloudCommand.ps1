<#
 
.SYNOPSIS
Generic wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

Works well with intense operations that may cause throttling

.DESCRIPTION
Wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

It accomplishs this by doing the following:
* Monitors the health of the Remote powershell session and restarts it as needed.
* Restarts the session every X number seconds to ensure a valid connection.
* Attempts to work past session related errors and will skip objects that it can't process.
* Attempts to calculate throttle exhaustion and sleep a sufficient time to allow throttle recovery

.PARAMETER Agree
Verifies that you have read and agree to the disclaimer at the top of the script file.

.PARAMETER AutomaticThrottle
Calculated value based on your tenants powershell recharge rate.
You tenant recharge rate can be calculated using a Micro Delay Warning message.

Look for the following line in your Micro Delay Warning Message
Balance: -1608289/2160000/-3000000 

The middle value is the recharge rate.
Divide this value by the number of milliseconds in an hour (3600000)
And subtract the result from 1 to get your AutomaticThrottle value

1 - (2160000 / 3600000) = 0.4

Default Value is .25

.PARAMETER Credential
Credential object for logging into Exchange Online Shell.
Prompts if there is non provided.

.PARAMETER IdentifyingProperty
What property of the objects we are processing that will be used to identify them in the log file and host
If the value is not set by the user the script will attempt to determine if one of the following properties is present
"DisplayName","Name","Identity","PrimarySMTPAddress","Alias","GUID"

If the value is not set and we are not able to match a well known property the script will generate an error and terminate.

.PARAMETER LogFile
Location and file name for the log file.

.PARAMETER ManualThrottle
Manual delay of X number of milliseconds to sleep between each cmdlets call.
Should only be used if the AutomaticThrottle isn't working to introduce sufficent delay to prevent Micro Delays

.PARAMETER NonInteractive
Suppresses output to the screen.  All output will still be in the log file.

.PARAMETER Recipients
Array of objects to operate on. This can be mailboxes or any other set of objects.
Input must be an array!
Anything comming in from the array can be accessed in the script block using $input.property

.PARAMETER ResetSeconds
How many seconds to run the script block before we rebuild the session with O365.

.PARAMETER ScriptBlock
The script that you want to robustly execute against the array of objects.  The Recipient objects will be provided to the cmdlets in the script block
and can be accessed with $input as if you were pipelining the object.

.LINK
http://EHLO.Link

.OUTPUTS
Creates the log file specified in -logfile.  Logfile contains a record of all actions taken by the script.

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\mbx.csv
$mbx = import-csv c:\temp\mbx.csv
$cred = get-Credential
.\Start-RobustCloudCommand.ps1 -Agree -Credential $cred -recipients $mbx -logfile C:\temp\out.log -ScriptBlock {Set-Clutter -identity $input.PrimarySMTPAddress.tostring() -enable:$false}

Gets all mailboxes from the service returning only Displayname,Identity, and PrimarySMTPAddress.  Exports the results to a CSV
Imports the CSV into a variable
Gets your O365 Credential
Executes the script setting clutter to off

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\recipients.csv
$recipients = import-csv c:\temp\recipients.csv
$cred = Get-Credential
.\Start-RobustCloudCommand.ps1 -Agree -Credential $cred -recipients $recipients -logfile C:\temp\out.log -ScriptBlock {Get-MobileDeviceStatistics -mailbox $input.PrimarySMTPAddress.tostring() | Select-Object -Property @{Name = "PrimarySMTPAddress";Expression={$input.PrimarySMTPAddress.tostring()}},DeviceType,LastSuccessSync,FirstSyncTime | Export-Csv c:\temp\stats.csv -Append }

Gets All Recipients and exports them to a CSV (for restartability)
Imports the CSV into a variable
Gets your O365 Credentials
Executs the script to gather EAS Device statistics and output them to a csv file


#>

function Start-RobustCloudCommand
{
    [CmdletBinding()]

    Param(
	    [switch]$Agree,
	    [Parameter(Mandatory=$true)]
	    [string]$LogFile,
	    [Parameter(Mandatory=$true)]
	    $Recipients,
	    [Parameter(Mandatory=$true)]
	    [ScriptBlock]$ScriptBlock,
	    $Credential,
	    [int]$ManualThrottle=0,
	    [double]$ActiveThrottle=.25,
	    [int]$ResetSeconds=870,
	    [string]$IdentifyingProperty,
	    [Switch]$NonInteractive
    )

    Process {
        # Writes output to a log file with a time date stamp
        Function Write-Log {
	        Param ([string]$string)
	
	        # Get the current date
	        [string]$date = Get-Date -Format G
		
	        # Write everything to our log file
	        ( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	
	        # If NonInteractive true then supress host output
	        if (!($NonInteractive)){
		        ( "[" + $date + "] - " + $string) | Write-Host
	        }
        }

        # Sleeps X seconds and displays a progress bar
        Function Start-SleepWithProgress {
	        Param([int]$sleeptime)

	        # Loop Number of seconds you want to sleep
	        For ($i=0;$i -le $sleeptime;$i++){
		        $timeleft = ($sleeptime - $i);
		
		        # Progress bar showing progress of the sleep
		        Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		
		        # Sleep 1 second
		        start-sleep 1
	        }
	
	        Write-Progress -Completed -Activity "Sleeping"
        }

        # Setup a new O365 Powershell Session
        Function New-CleanO365Session {
	
	        # If we don't have a credential then prompt for it
	        $i = 0
	        while (($Credential -eq $Null) -and ($i -lt 5)){
		        $script:Credential = Get-Credential -Message "Please provide your Exchange Online Credentials"
		        $i++
	        }
	
	        # If we still don't have a credentail object then abort
	        if ($Credential -eq $null){
		        Write-log "[Error] - Failed to get credentials"
		        Write-Error -Message "Failed to get credentials" -ErrorAction Stop
	        }

	        Write-Log "Removing all PS Sessions"

	        # Destroy any outstanding PS Session
	        Get-PSSession | Remove-PSSession -Confirm:$false
	
	        # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
	        [System.GC]::Collect()
	
	        # Sleep 15s to allow the sessions to tear down fully
	        Write-Log ("Sleeping 15 seconds for Session Tear Down")
	        Start-SleepWithProgress -SleepTime 15

	        # Clear out all errors
	        $Error.Clear()
	
	        # Create the session
	        Write-Log "Creating new PS Session"
	
	        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
		
	        # Check for an error while creating the session
	        if ($Error.Count -gt 0){
	
		        Write-Log "[ERROR] - Error while setting up session"
		        Write-log $Error
		
		        # Increment our error count so we abort after so many attempts to set up the session
		        $ErrorCount++
		
		        # if we have failed to setup the session > 3 times then we need to abort because we are in a failure state
		        if ($ErrorCount -gt 3){
		
			        Write-log "[ERROR] - Failed to setup session after multiple tries"
			        Write-log "[ERROR] - Aborting Script"
			        exit
		
		        }
		
		        # If we are not aborting then sleep 60s in the hope that the issue is transient
		        Write-Log "Sleeping 60s so that issue can potentially be resolved"
		        Start-SleepWithProgress -sleeptime 60
		
		        # Attempt to set up the sesion again
		        New-CleanO365Session
	        }
	
	        # If the session setup worked then we need to set $errorcount to 0
	        else {
		        $ErrorCount = 0
	        }
	
	        # Import the PS session
	        $null = Import-PSSession $session -AllowClobber
	
	        # Set the Start time for the current session
	        Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
        }

        # Verifies that the connection is healthy
        # Goes ahead and resets it every $ResetSeconds number of seconds either way
        Function Test-O365Session {
	
	        # Get the time that we are working on this object to use later in testing
	        $ObjectTime = Get-Date
	
	        # Reset and regather our session information
	        $SessionInfo = $null
	        $SessionInfo = Get-PSSession
	
	        # Make sure we found a session
	        if ($SessionInfo -eq $null) { 
		        Write-Log "[ERROR] - No Session Found"
		        Write-log "Recreating Session"
		        New-CleanO365Session
	        }	
	        # Make sure it is in an opened state if not log and recreate
	        elseif ($SessionInfo.State -ne "Opened"){
		        Write-Log "[ERROR] - Session not in Open State"
		        Write-log ($SessionInfo | fl | Out-String )
		        Write-log "Recreating Session"
		        New-CleanO365Session
	        }
	        # If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	        elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		        Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
		        Write-Log "Rebuilding Connection"
		
		        # Estimate the throttle delay needed since the last session rebuild
		        # Amount of time the session was allowed to run * our activethrottle value
		        # Divide by 2 to account for network time, script delays, and a fudge factor
		        # Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		        [int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		        # If the delay is >15s then sleep that amount for throttle to recover
		        if ($DelayinSeconds -gt 0){
		
			        Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
			        Start-SleepWithProgress -SleepTime $DelayinSeconds
		        }
		        # If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
		        else {
			        Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		        }
				
		        # new O365 session and reset our object processed count
		        New-CleanO365Session
	        }
	        else {
		        # If session is active and it hasn't been open too long then do nothing and keep going
	        }
	
	        # If we have a manual throttle value then sleep for that many milliseconds
	        if ($ManualThrottle -gt 0){
		        Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
		        Start-Sleep -Milliseconds $ManualThrottle
	        }
        }

        # If the $identifyingProperty has not been set then we attempt to locate a value for tracking modified objects
        Function Get-ObjectIdentificationProperty {
	        Param($object)
	
	        Write-Log "Trying to identify a property for displaying per object progress"
	
	        # Common properties to check
	        [array]$PropertiesToCheck = "DisplayName","Name","Identity","PrimarySMTPAddress","Alias","GUID"
	
	        # Set our counter to 0
	        $i = 0
	
	        # While we haven't found an ID property continue checking
	        while ([string]::IsNullOrEmpty($IdentifyingProperty))
	        {
	
		        # If we have gone thru the list then we need to throw an error because we don't have Identity information
		        # Set the string to bogus just to ensure we will exit the while loop
		        if ($i -gt ($PropertiesToCheck.length -1))
		        {
			        Write-Log "[ERROR] - Unable to find a common identity parameter in the input object"
			
			        # Create an error message that has all of the valid property names that we are looking for
			        $PropertiesToCheck | foreach { [string]$PropertiesString = $PropertiesString + "`"" + $_ + "`", " }
			        $PropertiesString = $PropertiesString.TrimEnd(", ")
			        [string]$errorstring = "Objects does not contain a common identity parameter " + $PropertiesString + " please use -IdentifyingProperty to set the identity value"
			
			        # Throw error
			        Write-Error -Message $errorstring -ErrorAction Stop
		        }
		
		        # Get the property we are testing out of our array
		        [string]$Property = $PropertiesToCheck[$i]
	
		        # Check the properties of the object to see if we have one that matches a well known name
		        # If we have found one set the value to that property
		        if ($object.$Property -ne $null)
		        { 
			        Write-log ("Found " + $Property + " to use for displaying per object progress")
			        Set-Variable -Scope script -Name IdentifyingProperty -Value $Property
		        }
		
		        # Increment our position counter
		        $i++
		
	        }
        }

        # Gather and print out information about how fast the script is running
        Function Get-EstimatedTimeToCompletion {
	        param([int]$ProcessedCount)
	
	        # Increment our count of how many objects we have processed
	        $ProcessedCount++
	
	        # Every 100 we need to estimate our completion time and write that out
	        if (($ProcessedCount % 100) -eq 0){
	
		        # Get the current date
		        $CurrentDate = Get-Date
		
		        # Average time per object in seconds
		        $AveragePerObject = (((($CurrentDate) - $ScriptStartTime).totalseconds) / $ProcessedCount)
		
		        # Write out session stats and estimated time to completion
		        Write-Log ("[STATS] - Total Number of Objects:     " + $ObjectCount)
		        Write-Log ("[STATS] - Number of Objects processed: " + $ProcessedCount)
		        Write-Log ("[STATS] - Average seconds per object:  " + $AveragePerObject)
		        Write-Log ("[STATS] - Estimated completion time:   " + $CurrentDate.addseconds((($ObjectCount - $ProcessedCount) * $AveragePerObject)))
	        }
	
	        # Return number of objects processed so that the variable in incremented
	        return $ProcessedCount
        }

        ####################
        # Main Script
        ####################

        # Force use of at least version 3 of powershell https://technet.microsoft.com/en-us/library/hh847765.aspx
        #Requires -version 3

        # Turns on strict mode https://technet.microsoft.com/library/03373bbe-2236-42c3-bf17-301632e0c428(v=wps.630).aspx
        Set-StrictMode -Version 2

        # Write creation date of script for version information
        Write-Log "Created 05/10/2016"

        # Statement to ensure that you have looked at the disclaimer or that you have removed this line so you don't have too
        if ($Agree -ne $true){ Write-Error "Please run the script with -Agree to indicate that you have read and agreed to the sample script disclaimer at the top of the script file" -ErrorAction Stop }
        else { Write-log "Agreed to Disclaimer" }

        # Log the script block for debugging purposes
        Write-log $ScriptBlock

        # Setup our first session to O365
        $ErrorCount = 0
        New-CleanO365Session

        # Get when we started the script for estimating time to completion
        $ScriptStartTime = Get-Date

        # Get the object count and write it out to be used in esitmated time to completion + logging
        [int]$ObjectCount = $Recipients.count
        [int]$ObjectsProcessed = 0

        # If we don't have an identifying property then try to find one
        if ([string]::IsNullOrEmpty($IdentifyingProperty))
        {
	        # Call our function for finding an identifying property and pass in the first recipient object
	        Get-ObjectIdentificationProperty -object $Recipients[0]
        }

        # Go thru each recipient object and execute the script block
        foreach ($object in $Recipients)
        {
	
	        # Set our initial while statement values
	        $TryCommand = $true
	        $errorcount = 0
	
	        # Try the command 3 times and exit out if we can't get it to work
	        # Record the error and restart the session each time it errors out
	        while ($TryCommand)
	        {
		        Write-log ("Running scriptblock for " + ($object.$IdentifyingProperty).tostring())
		
		        # Clear all errors
		        $Error.Clear()
	
		        # Test our connection and rebuild if needed
		        Test-O365Session
	
		        # Invoke the script block
		        Invoke-Command -InputObject $object -ScriptBlock $ScriptBlock
		
		        # Test for errors
		        if ($Error.Count -gt 0) 
		        {
			        # Write that we failed
			        Write-Log ("[ERROR] - Failed on object " + ($object.$IdentifyingProperty).tostring())
			        write-log $Error
			
			        # Increment error count
			        $errorcount++
			
				        # Handle if we keep failing on the object
				        if ($errorcount -ge 3)
				        {
					        Write-Log ("[ERROR] - Oject has failed three times " + ($object.$IdentifyingProperty).tostring())
					        Write-Log ("[ERROR] - Skipping Object")
					
					        # Increment the object processed count / Estimate time to completion
					        $ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed
					
					        # Set trycommand to false so we abort the while loop
					        $TryCommand = $false
				        }
				        # Otherwise try the command again
				        else 
				        {
					        Write-Log ("Rebuilding session and trying again")
					        # Create a new session in case the error was due to a session issue
					        New-CleanO365Session
				        }
		        }
		        else 
		        {
			        # Since we didn't get an error don't run again
			        $TryCommand = $false
			
			        # Increment the object processed count / Estimate time to completion
			        $ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed
		        }
	        }
        }

        Write-Log "Script Complete Destroying PS Sessions"
        # Destroy any outstanding PS Session
        Get-PSSession | Remove-PSSession -Confirm:$false
    }
}