#20241015
#changed popup to not display if it is the initial poll.

Add-Type -AssemblyName System.Windows.Forms    
Add-Type -AssemblyName System.Drawing

$NewLine = [Environment]::NewLine
$bDebug = $False

$Dock_Top = 1
$Dock_Bottom = 2
$Dock_Left = 3
$Dock_Right = 4
$Dock_Fill = 5

$Color_Red = "Red"
$Color_Orange = "Orange"
$Color_Green = "LightGreen"

$RunspaceStateNotStarted = 'NotStarted'
$RunspaceStateRunning = 'Running'
$RunspaceStateCompleted = 'Completed'

$ListView_Details = 1

$Splitter_Orientation_Horizontal = 0
$Splitter_Orientation_Vertical = 1

$TextAlign_Left_Middle = 16
$TextAlign_Center_Middle = 32

$SNMPPollPeriod = 60
$OIDsysUptime = ".1.3.6.1.2.1.1.3.0"

$global:htHostsAck = @{} #Acknowledgement of host down. Default $False. $True if dialog was shown for host down.
$global:htPingHosts = @{}
$global:sSNMPCommString = ''
$global:dtSNMPLastPoll = '' #Datetime stamp of last poll
$global:sLog = ''

$SNMPPoller = New-Object -ComObject olePrn.OleSNMP

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console {
$consolePtr = [Console.Window]::GetConsoleWindow()
#5 show
[Console.Window]::ShowWindow($consolePtr, 5) | Out-Null
}

function Hide-Console {
$consolePtr = [Console.Window]::GetConsoleWindow()
#0 hide
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
}

Function GetStamp()
{
	return Get-Date -UFormat "%I:%M:%S"
}

Function WriteToLog
{
	param($LogEntry)
	if ($bDebug -eq $False) {Return}
	
	$Stamp = GetStamp
	#Try to write out. Form may be closed.
	$global:sLog = $global:sLog + $Stamp + ' ' + $LogEntry + $NewLine
	try{$txtDebug.text = $global:sLog} catch {}
}

Function frmMain_Resize()
{
	$ButtonEvents.Left = $frmMain.Width - ($ButtonEvents.Width + 20)
}

Function frmLoad_Resize()
{
	$LoadButtonOk.Left = $frmLoad.Width - ($LoadButtonOk.Width + 20)
}

Function LoadButtonParseIPS_Click()
{
	$Results = ''
	$IPAddressPattern = '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'	
	$RegExMatches = $txtLoadHosts.Text | Select-String  -Pattern $IPAddressPattern -AllMatches
	
	foreach($TestValue in $RegExMatches.Matches.Value)
		{
			$bValid = $True
			$Octets = $TestValue.Split('.')
			ForEach ($Octet in $Octets)				
				{
					try
					{
						if ([int]$Octet -gt 255)
						{
							$bValid = $False
						}
					}
					catch
					{
						$bValid = $False
					}
				}
			if ($bValid -eq $True)
				{
				if ($Results.Length -gt 0) {$Results = $Results + $NewLine}
				$Results = $Results + $TestValue
				}
		}
	if ($Results.Length -gt 0) {$txtLoadHosts.Text = $Results}
}

Function frmMain_KeyUp
{
	if($_.KeyCode -eq 'F1')
	{
		#Show help
		$helptext = ''
		$helptext = $helptext + 'F1 - Help.' + $NewLine
		$helptext = $helptext + 'F5 - Start polling.' + $NewLine
		$helptext = $helptext + 'F6 - Stop polling.' + $NewLine
		$helptext = $helptext + 'F8 - Set SNMP community string.' + $NewLine
		$helptext = $helptext + 'F9 - Hide Details/Events.' + $NewLine
		$helptext = $helptext + 'F10 - Show Details.' + $NewLine
		$helptext = $helptext + 'F11 - Show Events.' + $NewLine
		if ($bDebug -eq $True) {$helptext = $helptext + 'F12 - Debug.' + $NewLine}
		$helptext = $helptext + 'Insert - load hosts.' + $NewLine
		[System.Windows.Forms.MessageBox]::Show($helptext, 'psPing Help')
	}
	
	if($_.KeyCode -eq 'F5') {SetStateStart}
	if($_.KeyCode -eq 'F6') {SetStateStop}
	if($_.KeyCode -eq 'F8') {SetSNMPCommunities}
	if($_.KeyCode -eq 'F9') {HideDetailsEvents}
	if($_.KeyCode -eq 'F10') {ShowDetails}
	if($_.KeyCode -eq 'F11') {ShowEvents}
	if($_.KeyCode -eq 'F12') {DebugShow}	
	if($_.KeyCode -eq 'Insert') {ButtonLoad_Click}
}


Function lvHosts_Click()
{	
	#Show with a single click
	if($lvHosts.SelectedItems.Count -gt 0)
		{
			$lvHost = $lvHosts.SelectedItems[0]
        	 	$hostName = $lvHost.Text
        
        		# Update the title bar of the main form with the selected host
        		$frmMain.Text = "Selected Host: $hostName - Status: $($lvHost.SubItems[1].Text)"
		}
}


Function lvHosts_DoubleClick
{
	#Node will be selected to doubleclick it.
	if($lvHosts.SelectedItems.Count -gt 0)
		{
			$lvHost = $lvHosts.SelectedItems[0]			
			
			#Create Details Form
			$frmHost = New-Object system.Windows.Forms.Form
			$frmHost.Text = $lvHost.Text
			$frmHost.Width = 500
			$frmHost.Height = 350
			$frmHost.MinimumSize = new-object System.Drawing.Size(250,200)
			
			$txtHost = New-Object System.Windows.Forms.TextBox			
			$txtHost.Font = New-Object System.Drawing.Font ("Courier New",10)
			$txtHost.Multiline = $true
			$txtHost.Scrollbars = 2 #Vertical Only
			$txtHost.Dock = $Dock_Fill
			$frmHost.Controls.Add($txtHost)
			$frmHost.Show($frmMain)
			
			$htHost = @{}
			$htHost.Add('Host',$lvHost.Text)
			$htHost.Add('Textbox',$txtHost)
			$htHost.Add('Display','Details')
			$global:htDetailEvents.Add($lvHost.Text,$htHost)
		}	
}

Function LoadHosts()
{
	param($sHosts)
	
	$lvHosts.Items.Clear()
	$txtResults.text = ''
	
	#Clear Runspace and Powershell instances
	ForEach ($htHost in $global:htPingHosts)
	{
		try
		{
			$htHost.PSInstance.Dispose()
			$htHost.RunSpace.Close()
			$htHost.RunSpace.Dispose()
		}
		catch
		{
			#Instances not set.
		}
	}
	
	$global:htPingHosts = @{}

	$Hosts = $sHosts.Split($NewLine)
	ForEach($HostEntry in $Hosts)
	{
		try
		{
			$TryPoint = ''
			if ($HostEntry -ne '')
			{
				$TryPoint = 'Building hash table entry'
				#Add $global:htPingHosts Entry
				$htHost = @{}
				$htHost.HostName = $HostEntry
				$htHost.sDetails = ''
				$htHost.sEvents = GetStamp + ' ' + 'Starting'
				$htHost.PreviousState = 'Starting'
				$htHost.PreviousDateStamp = Get-Date
				$htHost.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
				$htHost.PSInstance = [Powershell]::Create()
				$htHost.Runspace.Open()
				$htHost.PSInstance.Runspace = $htHost.Runspace
				$global:htPingHosts.Add($HostEntry,$htHost)	
				$global:htHostsAck.Add($HostEntry,$False)	
				
				#Add to UI
				$TryPoint = 'Adding to UI'
				$li = New-Object System.Windows.Forms.ListViewItem($HostEntry)
				$li.subitems.Add('') | Out-Null
				$li.subitems.Add('') | Out-Null	
				$li.subitems.Add('') | Out-Null
				$lvHosts.Items.Add($li) | Out-Null
			}
		}
		catch
		{
			WriteToLog ('[LoadHosts]Host ' + $HostEntry + ' - Failure Building Entry:' + $TryPoint)			
		}
	}
}

Function LoadButtonOk_Click
{
	#No Hosts
	if ($txtLoadHosts.Text -eq "") {$FrmLoad.Close();Return}
	
	#Stop timer
	SetStateStop
	
	#Clear Hosts
	$lvHosts.Items.Clear()
		
	$Entries = $txtLoadHosts.Text.Split($NewLine)
	LoadHosts($Entries)	
	
	$FrmLoad.Close()
}

Function ButtonLoad_Click()
{
	$frmLoad = New-Object system.Windows.Forms.Form
	$frmLoad.Text = "Load Host"
	$frmLoad.Width = 500
	$frmLoad.Height = 350
	$frmLoad.MinimumSize = new-object System.Drawing.Size(250,200)
	$frmLoad.Add_Resize({frmLoad_Resize})
		
	$LoadSplitterMain = New-Object System.Windows.Forms.SplitContainer
	$LoadSplitterMain.Orientation = $Splitter_Orientation_Horizontal
	$LoadSplitterMain.Dock = $Dock_Fill
	$frmLoad.Controls.Add($LoadSplitterMain)
	$LoadSplitterMain.SplitterDistance = 25
		
	$txtLoadHosts = New-Object System.Windows.Forms.TextBox
	$LoadSplitterMain.Panel2.Controls.Add($txtLoadHosts)
	$txtLoadHosts.Font = New-Object System.Drawing.Font ("Courier New",10)
	$txtLoadHosts.Multiline = $true
	$txtLoadHosts.Scrollbars = 2 #Vertical Only
	$txtLoadHosts.Dock = $Dock_Fill
	
	$LoadButtonOk = New-Object System.Windows.Forms.Button
	$LoadButtonOk.Text = "Ok"
	$LoadButtonOk.Height = 20
	$LoadButtonOk.Width = 50
	$LoadButtonOk.TextAlign = $TextAlign_Center_Middle
	$LoadButtonOk.Top = 5
	$LoadButtonOk.Left = $frmLoad.Width - $LoadButtonOk.Width - 20
	$LoadButtonOk.Enabled = $true
	$LoadSplitterMain.Panel1.Controls.Add($LoadButtonOk)
	$LoadButtonOk.Add_Click({LoadButtonOk_Click})

	$LoadButtonParseIPS = New-Object System.Windows.Forms.Button
	$LoadButtonParseIPS.Text = "Parse IPs"
	$LoadButtonParseIPS.Height = 20
	$LoadButtonParseIPS.Width = 100
	$LoadButtonParseIPS.TextAlign = $TextAlign_Center_Middle
	$LoadButtonParseIPS.Top = 5
	$LoadButtonParseIPS.Left = 5
	$LoadButtonParseIPS.Enabled = $true
	$LoadSplitterMain.Panel1.Controls.Add($LoadButtonParseIPS)
	$LoadButtonParseIPS.Add_Click({LoadButtonParseIPS_Click})
	
	$LoadSplitterMain.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1	
	$frmLoad.ShowDialog()
}

Function SetStateStart()
{
	$timerPoll.Enabled = $True
	$timerPoll.Start()
	$frmMain.Text = 'Running'
}

Function SetStateStop()
{
	$timerPoll.Stop()
	$timerPoll.Enabled = $False
	$frmMain.Text = 'Stopped'
}

Function SetSNMPCommunities()
{
	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
	
	$Title = 'SNMP Credentials'
	$Prompt = 'Input readonly or greater SNMP community string.'
    $Result = [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title)
	if ($Result -ne '') {$global:sSNMPCommString = $Result}	
}

Function HideDetailsEvents()
{
	$SplitterMain.Panel2.Hide()
	$SplitterMain.Panel2Collapsed = $True
}

Function ShowDetails()
{
	$SplitterMain.Panel2.Show()
	$SplitterMain.Panel2Collapsed = $False
	$txtResults.Tag = "Details"
}

Function ShowEvents()
{
	$SplitterMain.Panel2.Show()
	$SplitterMain.Panel2Collapsed = $False
	$txtResults.Tag = "Events"
}

Function GetTimeDiff()
{
	param($tsDiff) #Timespan
	
	$TryPoint = ''	
	$diff = ''
	try
	{
		$TryPoint = 'Calulate Days'
		if ($tsDiff.Days -ne 0) {$diff = $diff + $TSpan.Days.ToString() + 'D '}				
		$TryPoint = 'Calulate Hours'
		if ($tsDiff.Hours -ne 0) {$diff = $diff + $TSpan.Hours.ToString() + 'H '}
		$TryPoint = 'Calulate Minutes'
		if ($tsDiff.Minutes -ne 0) {$diff = $diff + $TSpan.Minutes.ToString() + 'M '}
		$TryPoint = 'Calulate Seconds'
		$diff = $diff + $tsDiff.Seconds.ToString() + 'S'
		
		#Set Stamp
		$TryPoint = 'Send value'
		return $diff		
	}
	catch
	{
		WriteToLog ('GetTimeDiff Failed: ' + $TryPoint)
		return ''
	}
}

Function timerPoll_Tick()
{
	ForEach($liHost in $lvHosts.Items)
	{
		$Entry = $liHost.Text
		$DTStamp = Get-Date
		
		$htHost = $global:htPingHosts[$Entry]
		$PingAsync = $htHost.PingAsync		
		$bLoadJob = $false
		
		$Results = @{}
		$Results.Status = ''
		$Results.ReplyFrom = ''
		$Results.Bytes = ''
		$Results.Time = ''
		$Results.TTL = ''
		
		#Recheck Invoke to determine state of execution in other intance
		$PollState = $htHost.PSInstance.InvocationStateInfo.State
				
		#Process States
		#Job is not started. Initial state
		if ($PollState -eq $RunspaceStateNotStarted)
		{
			$Results.Status = 'Initial'
			$bLoadJob = $True
		}
		
		#Job is still running.
		if ($PollState -eq $RunspaceStateRunning){$Results.Status = 'Running'}
		
		#Job is complete. Pull data
		if ($PollState -eq $RunspaceStateCompleted)
		{
			$TryPoint = ''
			try
			{
				#Pull Values.
				$TryPoint = 'Pulling invoke results'
				$PingResults = $htHost.PSInstance.EndInvoke($htHost.PSInvoke)
				
				$TryPoint = 'Getting Status'
				$Results.Status = $PingResults.Status
				
				if ($Results.Status -eq 'Success')
				{
					#Pull Values. If it fails, it will pull defaults from above.
					$TryPoint = 'Getting ReplyFrom'
					$Results.ReplyFrom = $PingResults.Address
					$TryPoint = 'Getting Bytes'
					$Results.Bytes = $PingResults.Buffer.length.ToString()
					$TryPoint = 'Getting Time'
					$Results.Time = $PingResults.RoundtripTime.ToString()
					$TryPoint = 'Getting TTL'
					$Results.TTL = $PingResults.Options.TTL.ToString()
				}
				
				$TryPoint = 'Setting LoadJob'
				$bLoadJob = $True
			}
			catch
			{
				WriteToLog ('[timerPoll_Tick]Host ' + $htHost.HostName + ' - Load Ping Status Failed:' + $TryPoint)
				$Results.Status = "Error"
				#Not much can be done to fix this.
				#Logging to write-output may spam the screen
			}
		}
		
		#Reload job if needed
		if($bLoadJob -eq $True)
		{
			$TryPoint = ''
			try
			{
				#Add Ping Script
				$TryPoint = 'Clear Old Scripts'
				#Remove Old Commands
				$htHost.PSInstance.Commands.Clear()
				
				$TryPoint = 'Adding Script'
				$htHost.PSInstance.AddScript({
				param($sHostname)
				
				$pinger = new-object system.net.networkinformation.ping
				$Results = $Pinger.Send($sHostName,1000)
				$Results
				})
				
				#Load Hostname Parameter
				$TryPoint = 'Adding Parameter'
				$htHost.PSInstance.AddParameter('sHostname',$htHost.HostName)
				
				#Push invoke back in for next cycle.
				$TryPoint = 'Invoking'
				$htHost.PSInvoke = $htHost.PSInstance.BeginInvoke()
			}
			catch
			{
				WriteToLog ('Host ' + $htHost.HostName + ' - LoadJob Failed:' + $TryPoint)
				#Not much can be done to fix this.
				#Logging to write-output may spam the screen
			}
		}
		
		#Use
		if ($PollState -ne 'Running')
		{
			$liHost.Subitems[1].text = $Results.Status
			$HostAck = $global:htHostsAck[$Entry]
			
			#Pull Data
			$CurrentStatus = $Results.Status #Will compare with previous status later
			
			if($CurrentStatus -eq 'Success')
			{
				#Ping Suceeded
				$Details = GetStamp
				$Details = $Details + " Reply from "
				$Details = $Details + $Results.ReplyFrom
				$Details = $Details + ": bytes="
				$Details = $Details + $Results.Bytes + " "
				$Details = $Details + "time=" + $Results.Time + "ms"
				$Details = $Details + " TTL=" + $Results.TTL
				
				$liHost.Backcolor = $Color_Green				
				if ($HostAck -eq $True)
				{
					$HostAck = $global:htHostsAck[$Entry] = $False					
				}
			}
			else
			{				
				$Details = GetStamp
				$Details = $Details + " Request Failed"
				
				$liHost.Backcolor = $Color_Red				
				if ($HostAck -eq $False)
				{
					$HostAck = $global:htHostsAck[$Entry] = $True
					#confirm device did not start offline. Otherwise the popup will be a false positive.
					if ($CurrentStatus -ne 'Initial') {[System.Windows.Forms.MessageBox]::Show("Host $($Entry) is down!", "Host Alert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)}
				}				
			}	
			
			#Update Details
			$htHost.Details = $htHost.Details + $Details + $NewLine
			
			#Trim count down to 50
			$CRLFSearch = Select-String -InputObject $htHost.Details -Pattern $NewLine -AllMatches
			if ($CRLFSearch.Matches.Count -gt 75)
				{
					#Reduce to 50
					try
					{
						$htHost.Details = $htHost.Details.SubString($CRLFSearch.Matches[50].Index + 1)
					}
					Catch
					{
						WriteToLog ('Host ' + $htHost.HostName + ' - Trim details count')
					}
				}
			
			#Check for State Change
			if ($CurrentStatus -ne $htHost.PreviousState)
			{
				#Changed
				
				#Get Total Difference
				if ($htHost.PreviousState -ne 'Initial')
				{
					try
					{
						$TSpan = $DTStamp - $htHost.PreviousStateStamp
						$diff = GetTimeDiff($TSpan)
						
						$TotalDiff = ' '					
						$TotalDiff = $TotalDiff + '('					
						$TotalDiff = $TotalDiff + $diff
						$TotalDiff = $TotalDiff + ')'
						$TotalDiff = $TotalDiff + ' '
					}
					catch
					{
						$TotalDiff = ' '
					}
				}
				else
				{
					#If it blows up, use blank
					$TotalDiff = ' '
				}
				
				#Update change markers
				$htHost.PreviousState = $CurrentStatus
				$htHost.PreviousStateStamp = $DTStamp
				#Write to Events
				$Event = $DTStamp.ToString()				
				$Event = $Event + $TotalDiff				
				$Event = $Event + $CurrentStatus
				$Event = $Event + $NewLine
				#Write event but only if it is not the Starting value
				if ($CurrentStatus -ne 'Initial') {$htHost.Events = $htHost.Events + $Event}
								
			}
			
			#Update timestamp in ListViewItem
			$diff = ''
			$TSpan = $DTStamp - $htHost.PreviousStateStamp
			$diff = GetTimeDiff($TSpan)
			$liHost.Subitems[2].text = $Diff
			
		}		
		#Update htHost
		$global:htPingHosts[$Entry] = $htHost
		
		#Update text pane
		if($lvHosts.SelectedItems.Count -gt 0)
		{
			$lvHost = $lvHosts.SelectedItems[0]
			$htHost = $htHost = $global:htPingHosts[$lvHost.Text]
			if($txtResults.Tag -eq "Details")
			{
				$txtResults.Text = $htHost.Details
				$txtResults.SelectionStart = $txtResults.Text.Length
				$txtResults.ScrollToCaret()
			}
			
			if($txtResults.Tag -eq "Events")
			{
				$txtResults.Text = $htHost.Events
				$txtResults.SelectionStart = $txtResults.Text.Length
				$txtResults.ScrollToCaret()
			}
			
		}
		else
		{
			$txtResults.Text = ''
		}
		#SNMP Poll
		if ($global:sSNMPCommString -ne ''){ProcessSNMP}
	}	
}

Function ProcessSNMP()
{
	#Called each timer tick. Only called if a community string is present
	#Check States
	$CurrentdtStamp = Get-Date
	
	#Check if dtSNMPLastPoll is blank. Set it. This will delay initial poll to allow for a few cycles so the form does not initally lock up. 
	if ($global:dtSNMPLastPoll -eq '') {$global:dtSNMPLastPoll = Get-Date}
	
	#Check if it is time to poll
	$ElapsedTime = [datetime]$CurrentdtStamp - [datetime]$global:dtSNMPLastPoll		
	if ($ElapsedTime.TotalSeconds -gt $SNMPPollPeriod)
	{
		#Reset time stamp
		$global:dtSNMPLastPoll = Get-Date
		
		#Time to poll
		$SNMPPoller = New-Object -ComObject olePrn.OleSNMP
		
		#Build hosts string.				
		ForEach($liHost in $lvHosts.Items)
		{
			$sHost = $liHost.Text
			$sUptime = ''
			WriteToLog ('[ProcessSNMP]Host - Processing: ' + $sHost)
			try
				{				
					$SNMPPoller.open($sHost,$sSNMPKey,2,1000) #Set retry to 2. 3 Was too long.
					$SNMPPollResult = [TimeSpan]::FromSeconds(($SNMPPoller.Get(".1.3.6.1.2.1.1.3.0"))/100)					
					if ($SNMPPollResult.Days -gt 0) {$sUptime = $sUptime + $SNMPPollResult.Days.ToString() + 'd ' }
					if ($SNMPPollResult.Hours -gt 0) {$sUptime = $sUptime + $SNMPPollResult.Hours.ToString() + 'h ' }
					if ($SNMPPollResult.Minutes -gt 0) {$sUptime = $sUptime + $SNMPPollResult.Minutes.ToString() + 'm ' }
					if ($SNMPPollResult.Seconds -gt 0) {$sUptime = $sUptime + $SNMPPollResult.Seconds.ToString() + 's' }
					$SNMPPoller.Close()				
				}
				catch
				{
					$sUptime = 'n/a'				
				}
			$liHost.Subitems[3].text = $sUptime
		}				
	}	
}

Function DebugShow()
{
	if ($bDebug -ne $True) {return}
	$frmDebug = New-Object system.Windows.Forms.Form
	$frmDebug.Text = "Debug"
	$frmDebug.Width = 1000
	$frmDebug.Height = 600
	$frmDebug.MinimumSize = new-object System.Drawing.Size(500,400)
	
	$txtDebug = New-Object System.Windows.Forms.TextBox
	$txtDebug.Font = New-Object System.Drawing.Font ("Courier New",10)
	$txtDebug.Multiline = $true
	$txtDebug.Scrollbars = 2 #Vertical Only
	$txtDebug.Dock = $Dock_Fill
	$frmDebug.Controls.Add($txtDebug)

	$txtDebug.text = $global:sLog
	$frmDebug.show($frmMain)
}

$timerPoll = New-Object System.Windows.Forms.Timer
$timerPoll.Interval = 1000
$timerPoll.Enabled = $False
$timerPoll.add_tick({timerPoll_Tick})

$frmMain = New-Object system.Windows.Forms.Form
$frmMain.Text = "Hit F1 for Help"
$frmMain.Width = 395
$frmMain.Height = 150
$frmMain.KeyPreview = $True
$frmMain.MinimumSize = new-object System.Drawing.Size(295,100)
$frmMain.Add_KeyUp({frmMain_KeyUp})

$SplitterMain = New-Object System.Windows.Forms.SplitContainer  #Toolbar/Hosts
$SplitterMain.Orientation = $Splitter_Orientation_Horizontal
$SplitterMain.Dock = $Dock_Fill
$frmMain.Controls.Add($SplitterMain)
$SplitterMain.SplitterDistance = 25

$lvHosts = New-Object System.Windows.Forms.ListView
$lvHosts.View = 1 #Details
$lvHosts.FullRowSelect = $true	
$lvHosts.Dock = $Dock_Fill	
$lvHosts.Columns.Add('Host',100) | Out-Null
$lvHosts.Columns.Add('State',75) | Out-Null
$lvHosts.Columns.Add('Duration',100) | Out-Null
$lvHosts.Columns.Add('SNMP Uptime',100) | Out-Null
$lvHosts.Add_DoubleClick({lvHosts_DoubleClick})
$SplitterMain.Panel1.Controls.Add($lvHosts)
$lvHosts.Add_Click({lvHosts_Click})

$txtResults = New-Object System.Windows.Forms.TextBox
$SplitterMain.Panel2.Controls.Add($txtResults)
$txtResults.Font = New-Object System.Drawing.Font ("Courier New",10)
$txtResults.Multiline = $true
$txtResults.Scrollbars = 2 #Vertical Only
$txtResults.Dock = $Dock_Fill

$SplitterMain.Panel2.Hide()
$SplitterMain.Panel2Collapsed = $True

#Process Command line arguments.
#Can be 0,1,or 2.
#If it is 1, they will be a comma seperated list of hosts to preload.
#If 2 then load the csv of host if 1 is 'debug'. Turn on debugging
$StartParameters = $Args

if ($StartParameters.count -eq 1) {LoadHosts($StartParameters[0])}
if ($StartParameters.count -eq 2) {if ($StartParameters[0] -eq 'debug') {$bDebug = $True;LoadHosts($StartParameters[1])}}

Hide-Console
$frmMain.ShowDialog() | out-Null
$timerPoll.Dispose()

ForEach ($htHost in $global:htPingHosts)
	{
		try
		{
			$htHost.PSInstance.Dispose()
			$htHost.RunSpace.Close()
			$htHost.RunSpace.Dispose()
		}
		catch
		{
		}
	}

Show-Console
