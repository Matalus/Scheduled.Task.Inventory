

#Programatically documents all scheduled tasks that meet certain filters 
$ErrorActionPreference = "Stop"
$RunDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$jsonpath = "$RunDir\servers.txt"
if(Test-Path $jsonpath){
    [array]$Servers = Get-Content $jsonpath
}else{
    Write-Error "Missing servers.txt in root directory"
}


Clear-Host
"Scheduled.Task.Inventory"
"This Script will use NTLM and Invoke-Command to remotely connect to the servers on the same domain and inventory scheduled tasks"
""
"Servers: $($Servers.Count) "
$Servers | ForEach-Object {"$_"}
""

if((Read-Host -Prompt "If you want to run in update mode type Y or press enter to continue in report mode") -like "*Y*"){
    $update = $true # default false : set to true to write changes from modified report
    Write-Host -ForegroundColor Yellow "Running in Update Mode"
}else{
    $update = $false
    Write-Host -ForegroundColor Magenta "Running in Report Only Mode"
}

if($update){
    [array]$Files = Get-ChildItem $RunDir -Filter "*.csv"
    if($Files.Count -ne 1){
        $Files = $Files | Out-GridView -Title "Select Update File" -PassThru  
    }
    $updateCSV = Import-Csv -Path $Files.FullName -Encoding UTF8
}


Function Get-OrdinalNumber {
    Param(
        [Parameter(Mandatory=$true)]
        [int64]$num
    )

    $Suffix = Switch -regex ($Num) {
        '1(1|2|3)$' { 'th'; break }
        '.?1$'      { 'st'; break }
        '.?2$'      { 'nd'; break }
        '.?3$'      { 'rd'; break }
        default     { 'th'; break }
    }
    Write-Output "$Num$Suffix"
}

Function Task-Match ($Csv, $Name, $Server){
    if ($Csv -and $Name -and $Server){
        [array]$Task = $Csv | Where-Object {
            $_.Name -eq $Name -and
            $_.Server -eq $Server
        }
        if($Task.Count -eq 1){
            Return $Task
        }else{
            Return Write-Error "Found $($Task.Count) Tasks that match this description `nName: $Name`nServer: $Server"
        }
    }else{
        Return Write-Error "Unable to validate inputs"
    }
}

# Formats JSON in a nicer format than the built-in ConvertTo-Json does.
function Format-Json([Parameter(Mandatory, ValueFromPipeline)][String] $json) {
    $indent = 0;
    ($json -Split '\n' |
      % {
        if ($_ -match '[\}\]]') {
          # This line contains  ] or }, decrement the indentation level
          $indent--
        }
        $line = (' ' * $indent * 2) + $_.TrimStart().Replace(':  ', ': ')
        if ($_ -match '[\{\[]') {
          # This line contains [ or {, increment the indentation level
          $indent++
        }
        $line
    }) -Join "`n"
  }



$TaskInventory = @() #Array of PSObjects for Task information

#loop through servers
ForEach($Server in $Servers){

    #Collect Task Information
    $tasks = Invoke-Command -ScriptBlock{
        Get-ScheduledTask | Where-Object {
            $_.Actions[0].Execute -like "*PowerShell*" -and 
            $_.Actions[0].Arguments -notlike "*System*"
        }
    } -ComputerName $Server

    ForEach($Task in $Tasks){
        "TaskName: $($Task.TaskName) : $($Task.PSComputerName)"
        
        $Triggers = $Task.Triggers
        $TriggerType = $Triggers[0].CimClass.CimClassName.Replace("MSFT_Task","").Replace("Trigger","")
        if($TriggerType.length -lt 1){
            $TriggerType = "Monthly"
        }

        #$TriggerType
        #$Trigger
        [array]$Trigger_Desc = $null
        ForEach($Trigger in $Triggers){
            $Trigger_Desc += if($TriggerType -like "*Weekly*"){
                [string]$dow = ($Trigger.DaysOfWeek | ForEach-Object {[DayOfWeek]$_}) -join ", "
                [string]$tod = ([datetime]$Trigger.StartBoundary).TimeOfDay.ToString()
                "$($dow) @ $tod"
            }elseif($TriggerType -like "*Daily*"){
                if($Trigger.Repetition.Interval -like "*H"){               
                    [REGEX]::Matches($Trigger.Repetition.Interval,"\d+").value + " Hours"
                }elseif($Trigger.Repetition.Interval -like "*M"){
                    [REGEX]::Matches($Trigger.Repetition.Interval,"\d+").value + " Minutes"
                }elseif($Trigger.Repetition.Interval -eq $null){
                    [string]$tod = ([datetime]$Trigger.StartBoundary).TimeOfDay.ToString()
                    "@ $tod"
                }else{
                    "Custom Daily"
                }
            }else{
                $schserv = new-object -com("schedule.service");
                $schserv.connect($task.PSComputerName);
                $comxml = $schserv.getfolder("\").gettask($task.TaskName);
                ### This will list out the xml for the entire task.
                #convert to string for parsing
                [string]$rawxml = $comxml.xml
                [xml]$taskxml = $rawxml
                [string]$tod = ([datetime]$taskxml.Task.Triggers.CalendarTrigger.StartBoundary).TimeOfDay.ToString()
                $dom = $taskxml.task.Triggers.CalendarTrigger.ScheduleByMonth.DaysOfMonth.Day
                $month_text = $taskxml.Task.Triggers.CalendarTrigger.ScheduleByMonth.Months.ChildNodes.name
                $months_num = $month_text | ForEach-Object {([datetime]"01 $_ 2018").Month }

                if((Compare-Object -ReferenceObject $months_num -DifferenceObject (1..12)) -eq $null){
                    "$(Get-OrdinalNumber $dom) of month @ $tod, Every Month"
                }else{
                    "$(Get-OrdinalNumber $dom) of month @ $tod, Every $($month_text -join ", ")"
                }
                

            }
        }

        $TaskDetails = [pscustomobject]@{
            Name = $Task.TaskName
            Server = $Task.PSComputerName
            Trigger_Type  = $TriggerType
            Trigger_Frequency = $Trigger_Desc -join ", "
            Run_As = $Task.Principal.UserId
            Recipients = $null
            Script_Type = $Task.Actions[0].Execute
            Script_Path = $Task.Actions[0].Arguments
            Working_Dir = $Task.Actions[0].WorkingDirectory
        }

        $filepath = (($TaskDetails.Script_Path -split "-file")[1].Trim(" ")) -replace '"'
        $trim_enum = ($filepath -split "\\").Length - 2
        $Directory = ($filepath -split "\\")[0..$trim_enum] -join "\"
        $DirContents = Invoke-Command -ScriptBlock { Get-ChildItem $args[0] } -ComputerName $task.PSComputerName -ArgumentList $Directory

        $configFile = $null
        $configFile = $DirContents | Where-Object{
            $_.Name -like "config.json*" -or
            $_.Name -like "securitylist.txt"
        }

        $reportconfig = $null
        if($configFile){
            $filecontents = Invoke-Command -ScriptBlock{
              Get-Content $args[0]
            } -ComputerName $Task.PSComputerName -ArgumentList $configFile.FullName

            if($configFile.Name -like "config.json.txt"){
                $reportconfig = ($filecontents) -join "`n" | ConvertFrom-Json
                $TaskDetails.Recipients = if($reportconfig.To){
                    $reportconfig.To
                }else{
                    $reportconfig.EmailTo
                }
            }elseif($configFile.Name -like "securitylist.txt"){
                #if update is turned on
                if($update){
                    $CSVTask = Task-Match -Csv $updateCSV -Name $TaskDetails.Name -Server $TaskDetails.Server
                    [array]$updateRecipients = $CSVTask.Recipients.Split(",")
                    [array]$compare = Compare-Object -ReferenceObject $filecontents -DifferenceObject $updateRecipients
                    if($compare.count -ge 1){
                        Write-Host -ForegroundColor Yellow "Differences Detected - Updating Config File..."
                        $compare
                        $filecontents = $updateRecipients
                        Invoke-Command -ScriptBlock {
                            $args[0] | Out-File $args[1] -Force
                        } -ComputerName $Task.PSComputerName -ArgumentList $filecontents,$configFile.FullName
                    }else{
                        Write-Host -ForegroundColor Cyan "No Recipient Mismatches found"
                    }
                }
                
                # runs regardless of update
                $TaskDetails.Recipients = $filecontents

            }elseif($configFile.Name -like "config.json"){
                $reportconfig = ($filecontents) -join "`n" | ConvertFrom-Json
                #if Update is turned on
                if($update){
                    $CSVTask = Task-Match -Csv $updateCSV -Name $TaskDetails.Name -Server $TaskDetails.Server
                    [array]$updateRecipients = $CSVTask.Recipients.Split(",")
                    [array]$compare = Compare-Object -ReferenceObject $reportconfig.Recipients -DifferenceObject $updateRecipients
                    if($compare.count -ge 1){
                        Write-Host -ForegroundColor Yellow "Differences Detected - Updating Config File..."
                        $compare
                        $reportconfig.Recipients = $updateRecipients
                        $reportconfig_str = $reportconfig | ConvertTo-Json | Format-Json
                        Invoke-Command -ScriptBlock {
                            Write-Host -ForegroundColor Green "updating file contents: $($args[1])..."
                            $args[0] | Out-File $args[1] -Force
                        } -ComputerName $Task.PSComputerName -ArgumentList $reportconfig_str,$configFile.FullName

                    }else{
                        Write-Host -ForegroundColor Cyan "No Recipient Mismatches found"
                    }
                    
                }

                $TaskDetails.Recipients = $reportconfig.Recipients
            }else{
                $TaskDetails.Recipients = $null
            }
        }


        $TaskDetails

        $TaskInventory += $TaskDetails

    }
}

$TaskInventory | Format-Table -AutoSize

$TaskInventory | Select-Object Name,Server,Trigger_Type,Trigger_Frequency,Run_As,@{N="Recipients";E={$_.Recipients -join ","};},Script_Type,Script_Path,Working_Dir  | Export-Csv .\TaskInventory.csv -Force -NoTypeInformation

$TaskJSON = $TaskInventory | ConvertTo-Json | Set-Content .\TaskInventory.json

Invoke-Item .\TaskInventory.csv

$UniqueRecipients = $TaskInventory.Recipients | Select-Object -Unique


 