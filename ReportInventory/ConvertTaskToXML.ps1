$servername = "yourmachinename";
$taskname = "yourtaskname";
$schserv = new-object -com("schedule.service");
$schserv.connect($servername);
$task = $schserv.getfolder("\").gettask($taskname);
### This will list out the xml for the entire task.