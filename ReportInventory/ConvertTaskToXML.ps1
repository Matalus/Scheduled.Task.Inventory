$servername = $ENV:COMPUTERNAME
$task = Get-ScheduledTask | Where-Object{$_.TaskName -like "*Migration*"}
$taskname = $Task.TaskName
$schserv = new-object -com("schedule.service");
$schserv.connect($servername);
$xml = $schserv.getfolder("\").gettask($taskname);
### This will list out the xml for the entire task.