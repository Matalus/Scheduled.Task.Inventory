

#Programatically documents all scheduled tasks that meet certain filters 

#List of servers to inventory
$Servers = @(
    "IO-NJE-WEB01",
    "IO-AZR-WEB01",
    "IO-AZR-INT01"
    "IO-AZP-WEB01",
    "IO-OH-WEB01"
)

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
        "TaskName: $($Task.TaskName)"
        
        $Triggers = $Task.Triggers
        $TriggerType = $Triggers[0].CimClass.CimClassName.Replace("MSFT_Task","").Replace("Trigger","")

        $TriggerType
        $Trigger
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
                "Custom"
            }
        }

        $Recipeints = $null
        
        
        $TaskDetails = [pscustomobject]@{
            Name = $Task.TaskName
            Server = $Task.PSComputerName
            Trigger_Type  = $TriggerType
            Trigger_Frequency = $Trigger_Desc -join ", "
            Run_As = $Task.Principal.UserId
            Script_Type = $Task.Actions[0].Execute
            Script_Path = $Task.Actions[0].Arguments
            Working_Dir = $Task.Actions[0].WorkingDirectory
        }

        $TaskDetails

        $TaskInventory += $TaskDetails

    }
}

$TaskInventory | Format-Table -AutoSize
 