###################################################
##
## lib_infra
## Working with TFS API to get basic objects
## Aug. 2013 Dave Wu
##
###################################################

$execPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
. $execPath\config.ps1

$binpath  = $config["binpath"]
Add-Type -path "$binpath\Microsoft.TeamFoundation.Client.dll"
Add-Type -path "$binpath\Microsoft.TeamFoundation.TestManagement.Client.dll"
Add-Type -path "$binpath\Microsoft.TeamFoundation.WorkItemTracking.Client.dll"
Add-Type -path "$binpath\Microsoft.TeamFoundation.Build.Client.dll"

Function Get-TeamProjectObject{
    $tpc = New-Object -TypeName Microsoft.TeamFoundation.Client.TfsTeamProjectCollection($config["tfs"])
    $testmgt = $tpc.GetService([Microsoft.TeamFoundation.TestManagement.Client.ITestManagementService])
    $testmgt.GetTeamProject($config["teamproj"])
}

Function Get-BuildServerObject{
    $tpc = New-Object -TypeName Microsoft.TeamFoundation.Client.TfsTeamProjectCollection($config["tfs"])
    $tpc.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
}

Function Get-TestPlanObject{
    param([int]$planid)

    $teamproj = Get-TeamProjectObject
    # due to powershell bug, use reflection to get test plan helper object
    $testPlansProperty = [Microsoft.TeamFoundation.TestManagement.Client.ITestManagementTeamProject].GetProperty("TestPlans").GetGetMethod()
    $testPlans = $testPlansProperty.Invoke($teamproj, "instance,public", $null, $null, $null)   
    $testPlans.Find($planid)
}

Function Get-TestSuiteObject{
    param([int]$suiteid)

    $teamproj = Get-TeamProjectObject
    # due to powershell bug, use reflection to get test plan helper object
    $testSuitesProperty = [Microsoft.TeamFoundation.TestManagement.Client.ITestManagementTeamProject].GetProperty("TestSuites").GetGetMethod()
    $testSuites = $testSuitesProperty.Invoke($teamproj, "instance,public", $null, $null, $null)   
    $testSuites.Find($suiteid)
}

Function Get-TestRunObject{
    param([int]$runid)

    $teamproj = Get-TeamProjectObject    
    # due to powershell bug, use reflection to get test run helper object
    $testRunsProperty = [Microsoft.TeamFoundation.TestManagement.Client.ITestManagementTeamProject].GetProperty("TestRuns").GetGetMethod()
    $testRuns = $testRunsProperty.Invoke($teamproj, "instance,public", $null, $null, $null)
    $testRuns.Find($runid)   
}

Function Get-EnvironmentId{
    param(
    [string]$name
    )
    
    $env = (Get-TeamProjectObject).TestEnvironments.Query() | where{ $_.DisplayName -eq $name }
    return $env.Id
}

Function Get-EnvironmentName{
    param(
    [int]$testrunid
    )
    
    $envid = (Get-TestRunObject -runid $testrunid).TestEnvironmentId
    $env = (Get-TeamProjectObject).TestEnvironments.Query() | where{ $_.Id -eq $envid }
    return $env.DisplayName
}

Function Get-TestSettingId{
    param(
    [string]$name
    )
    
    $teamproj = Get-TeamProjectObject
    # due to powershell bug, use reflection to get test setting helper object
    $testSettingsProperty = [Microsoft.TeamFoundation.TestManagement.Client.ITestManagementTeamProject].GetProperty("TestSettings").GetGetMethod()
    $testSettings = $testSettingsProperty.Invoke($teamproj, "instance,public", $null, $null, $null)
    $res = $testSettings.Query("SELECT * FROM TestSettings")
    ($res | where{ $_.Name -eq $name }).Id
}

Function Get-TestPointObjectCollection{
    param([int]$planid,[int]$suiteid)
    
    $myplan = Get-TestPlanObject -planid $planid
    $myplan.QueryTestPoints("SELECT * FROM TestPoint WHERE SuiteId = $suiteid")
}

Function Get-ResultStatistics{
    param([int]$runid)
    (Get-TestRunObject -runid $runid).Statistics
}

Function Get-FailedCaseCollectionObject{
    param([int]$runid)
    (Get-TestRunObject -runid $runid).QueryResultsByOutcome([Microsoft.TeamFoundation.TestManagement.Client.TestOutcome]::Failed)
}

Function Get-FailedCaseIdArray{
    param([int]$runid)

    $failed_cases = Get-FailedCaseCollectionObject -runid $runid
    $array = @()
    $failed_cases | %{ $array += $_.TestCaseId }
    ,$array
}

Function Get-PassedCaseCollectionObject{
    param([int]$runid)
    (Get-TestRunObject -runid $runid).QueryResultsByOutcome([Microsoft.TeamFoundation.TestManagement.Client.TestOutcome]::Passed)
}

Function Get-PassedCaseIdArray{
    param([int]$runid)

    $passed_cases = Get-PassedCaseCollectionObject -runid $runid
    $array = @()
    $passed_cases | %{ $array += $_.TestCaseId }
    ,$array
}

Function Get-TotalCaseIdArray{
    param([int]$runid)

    $total_cases = (Get-TestRunObject -runid $runid).QueryResults()
    $array = @()
    $total_cases | %{ $array += $_.TestCaseId }
    ,$array
}