##################################################
##
##  Dave Wu - Feb 13, 2014
##  Code samples for PowerShell + TFS
##
##################################################

param(
[string]$filename
)

[xml]$file = gc $filename

$tfs = $file.WorkItemQuery.TeamFoundationServer
$teamproj = $file.WorkItemQuery.TeamProject
$wiql = $file.WorkItemQuery.Wiql

### Mind the version!! ###
### My VS is v12.0 (VS 2013). for VS 2012, the path is 11.0
$binpath   = "${env:ProgramFiles(x86)}\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0"
###

Add-Type -path "$binpath\Microsoft.TeamFoundation.Client.dll"
Add-Type -path "$binpath\Microsoft.TeamFoundation.TestManagement.Client.dll"
Add-Type -path "$binpath\Microsoft.TeamFoundation.WorkItemTracking.Client.dll"

$tpc = New-Object -TypeName Microsoft.TeamFoundation.Client.TfsTeamProjectCollection($tfs)
$querymgt = $tpc.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])

$results = $querymgt.Query($wiql)

## from here, you can write your code with $results ##
## below as example: get all test cases from several user stories

# for user stories
$related_testcases = $results | %{$_.Links}

# uniq test case ids
$testcase_ids = $related_testcases | %{$_.relatedworkitemid} | select -u

# write all ids into query
$str = ""
$testcase_ids | %{ $str += "$_," }
$str = $str.trimend(',')

# rendering output wiq
$content = @"
<?xml version="1.0" encoding="utf-8"?>
<WorkItemQuery Version="1">
  <TeamFoundationServer>$tfs</TeamFoundationServer>
  <TeamProject>$teamproj</TeamProject>
  <Wiql>SELECT [System.Id], [System.WorkItemType], [System.Title], [System.AssignedTo], [System.State], [Microsoft.VSTS.TCM.AutomationStatus] FROM WorkItems WHERE [System.WorkItemType] = 'Test Case'  AND  [System.State] &lt;&gt; 'Closed'  AND  [System.Id] IN ($str) ORDER BY [Microsoft.VSTS.TCM.AutomationStatus] </Wiql>
</WorkItemQuery>
"@

# if create a wiq file, to encode it as UTF8 is a must.
set-content -path ".\ttt.wiq" -value $content -Encoding UTF8