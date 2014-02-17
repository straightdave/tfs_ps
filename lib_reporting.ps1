#################################################
##
## lib_reporting
## providing functions related to reporting
## internal use (if not for debug)
## Aug. 2013 Dave Wu
## 
#################################################

$execPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
. $execPath\config.ps1
. $execPath\lib_infra.ps1

Function Send-HtmlMail{
    param(
    [string]$subject,
    [string[]]$tos,
    [string[]]$ccs,
    [string]$htmlbodyfile,
    [array]$attachments
    )

    $msg = New-Object -TypeName System.Net.Mail.MailMessage
    $msg.From = $config["mailfrom"]
    $msg.Subject = $subject
    
    $tos | %{ $msg.To.Add($_) }
    
    if($ccs -ne $null){
        $ccs | %{ $msg.CC.Add($_) }
    }
    
	if($attachments -ne $null){
        $attachments | %{ $msg.Attachments.Add($_) }
    }

    $msg.IsBodyHtml = $true
    $msg.Body = (gc $htmlbodyfile)
    
    $smtp = New-Object -TypeName System.Net.Mail.SmtpClient $config["smtphost"]
    $smtp.UseDefaultCredentials = $true # use the local credential
    $smtp.Send($msg)
}

Function Draw-PieSingle{
    param(
    [int]$runid,
    [string]$filename
    )

    $stat = Get-ResultStatistics -runid $runid
    
    $total = $stat.TotalTests
    $passed = $stat.PassedTests
    $failed = $stat.FailedTests
    $others = $total - $failed - $passed
    
    Draw-PieBase -passed $passed -others $others -failed $failed -filename $filename    
}

Function Draw-PieSeries{
    param(
    [int[]]$runids,
    [string]$filename
    )
    
    $passed_cases = @()
    $failed_cases = @()
    $total_cases = @()
    
    $runids = $runids | select -uniq    # filter off duplicated runs at first
    
    # for every instinct test run...
    $runids | %{
        $passed_cases += (Get-PassedCaseIdArray -runid $_)
        $failed_cases += (Get-FailedCaseIdArray -runid $_)
        $total_cases += (Get-TotalCaseIdArray -runid $_)
    }
    
    $passed_cases = $passed_cases | select -uniq
    $failed_cases = $failed_cases | select -uniq
    $total_cases = $total_cases | select -uniq
    
    # $failed_cases need to remove items existing in $passed_cases
    $passed_cases | %{
        if($failed_cases -contains $_){
            $failed_cases = $failed_cases -ne $_
        }    
    }

    $total_count = if($total_cases -is [System.Array]){$total_cases.count} elseif($total_cases -is [Int32]){1} else {0}
    $pass_count = if($passed_cases -is [System.Array]){$passed_cases.count} elseif($passed_cases -is [Int32]){1} else {0}
    $fail_count = if($failed_cases -is [System.Array]){$failed_cases.count} elseif($failed_cases -is [Int32]){1} else {0}
    
    $allothers =  $total_count - $pass_count - $fail_count
    
    #write-host "pass= $pass_count; fail= $fail_count; total= $total_count"
    Draw-PieBase -passed $pass_count -others $allothers -failed $fail_count -filename $filename
}

Function Draw-PieBase{
    param(
    [int]$passed,
    [int]$others,
    [int]$failed,
    [string]$filename
    )
        
    $total = $passed + $others + $failed
    #write-host "[Draw-PieBase] passed = $passed, others = $others, failed = $failed, total = $total"
        
    $rec = New-Object -TypeName System.Drawing.Rectangle(0,0,200,200)

    $blackbrush = New-Object -TypeName System.Drawing.SolidBrush([System.Drawing.Color]::black) # text
    $greenbrush = New-Object -TypeName System.Drawing.SolidBrush([System.Drawing.Color]::Green) # passed
    $redbrush   = New-Object -TypeName System.Drawing.SolidBrush([System.Drawing.Color]::Red)   # failed
    $bluebrush  = New-Object -TypeName System.Drawing.SolidBrush([System.Drawing.Color]::Blue)  # others

    $bmp = New-Object -TypeName System.Drawing.Bitmap(500,200)
    $graph = [System.Drawing.Graphics]::FromImage($bmp)
    
    $graph.FillPie($greenbrush, $rec,0,-($passed/$total)*360)
    $graph.FillPie($bluebrush,  $rec,-(($passed/$total)*360),-($others/$total)*360)
    $graph.FillPie($redbrush,   $rec,-(($passed+$others)/$total)*360,(($passed+$others)/$total)*360 - 360)
    
    $rec_pass_legend  = New-Object -TypeName System.Drawing.Rectangle(240,20,20,20)
    $rec_other_legend = New-Object -TypeName System.Drawing.Rectangle(240,60,20,20)
    $rec_fail_legend  = New-Object -TypeName System.Drawing.Rectangle(240,100,20,20)
    
    $graph.FillRectangle($greenbrush,$rec_pass_legend)
    $graph.FillRectangle($bluebrush,$rec_other_legend)
    $graph.FillRectangle($redbrush,$rec_fail_legend)
    
    $myfont = New-Object -TypeName System.Drawing.Font("Calibri",14,[system.Drawing.FontStyle]::Regular)
        
    $graph.DrawString("Passed: $passed ("+  ("{0:P}" -f ($passed/$total))  +")",$myfont,$blackbrush,270,18)
    $graph.DrawString("Others: $others ("+  ("{0:p}" -f ($others/$total))  +")",$myfont,$blackbrush,270,58)
    $graph.DrawString("Failed: $failed ("+  ("{0:p}" -f ($failed/$total))  +")",$myfont,$blackbrush,270,98)
    
    $bmp.Save($filename)
}

Function Create-HtmlMailSingle{
    param(
    [int]$runid,
    [boolean]$failedonly
    )

    $run = Get-TestRunObject -runid $runid
    $state = $run.State
    $build_number = $run.BuildNumber
    
    if(-not (Test-Path $config["localtemp"])){
        New-Item -ItemType Directory -Path $config["localtemp"] -Force
    }
    
    $uniq_name = "$runid" + (Get-DateAppendix)
    
    $tempfilename = $config["localtemp"] + "\" + $uniq_name + ".jpg"
    Draw-PieSingle -runid $runid -filename $tempfilename
    $pieaddress = $config["pieplace"] + "\" + $uniq_name + ".jpg"
    Copy-Item -Path $tempfilename -Destination $pieaddress -force
    
    if(-not $failedonly){
        $casetoshow = (Get-TestRunObject -runid $runid).QueryResults()
    }else{
        $casetoshow = Get-FailedCaseCollectionObject -runid $runid
    }

    $content = "<html>`n<head>`n"

    $content += "<style type=`"text/css`">`n"
    $content += "span.green { color:green }`n"
    $content += "span.red { color:red }`n"
    $content += "</style>`n"
    
    $content += "</head>`n<body>`n"
    $content += "<h2>Run:$runid Report - $state</h2>`n"
    
    $content += "Build Number: $build_number<br/><br/>`n"
    
    $content += "<div class=`"piechart`">`n"
    $content += "<img src=`"$pieaddress`" alt=`"pie_chart_of_this_run`" />`n"
    $content += "</div><br/>`n"
    
    if($casetoshow -ne $null){    
        $content += "<table border=`"1`">`n"
        $content += "<col /><col /><col width=`"500px`" /><col width=`"500px`" />`n"
        $content += "<tr bgcolor=`"C0C0C0`"><th></th><th>ID</th><th>Title</th><th>Error Message</th></tr>`n"
        
        $casetoshow | %{
            $tc_outcome = $_.Outcome.ToString()
            $tc_id = $_.TestCaseId
            $tc_title = $_.TestCaseTitle
            $tc_error = $_.ErrorMessage
            #$tc_owner = $_.OwnerName
            #$tc_pri = $_.Priority
        
            if($tc_outcome.Contains("Passed")){
                $tc_outcome = "<span class=`"green`">$tc_outcome</span>"
            }elseif($tc_outcome.Contains("Failed")){
                $tc_outcome = "<span class=`"red`">$tc_outcome</span>"
            }
        
            $content += "<tr><td>$tc_outcome</td><td>$tc_id</td><td>$tc_title</td><td>$tc_error</td></tr>`n"
        }        
        $content += "</table><br/>`n"
    }else{
        $content += "<p>Woops! No failed cases</p>`n"    
    }    

    $content += "Powered by Power MTM Shell (http://toolbox/pms). Do not reply.<br/>`n"
    $content += "Wicresoft MSIT SE Team, 2013`n"
    $content += "</body>`n</html>"
    
    $savehtml = $config["localtemp"] + "\$uniq_name.html"
    $content > $savehtml
    
    return $savehtml
}

Function Create-HtmlMailSeries{
    param(
    [int[]]$runids,
    [boolean]$failedonly
    )
    
    $hash_id_title = @{}
    $hash_id_errormsg = @{}
    
    # dynamically create some arrays to store all test case ids of each test run
    # array names are $a1, $a2, ...
    # and some arrays to store failed test case ids of each test run
    # these kind of arrays will named as $f1, $f2, ...
    $prefix = 1
    $build_numbers = @()
	$build_envs = @()
    $build_settings = @()
    $starts = @()
    $completes = @()
    $start = $null
    $complete = $null	
	
    $runids | %{
        # get all test case ids in this run, store into a dynamically created array 
        $ids = @()
        (Get-TestRunObject -runid $_).QueryResults() | %{
            $ids += $_.TestCaseId
            
            if( -not $hash_id_title.Contains($_.TestCaseId)){
                $hash_id_title.Add($_.TestCaseId,$_.TestCaseTitle)            
            }
        }
        New-Variable -Name "a$prefix" -Value $ids
        
        $fids = @()
        (Get-FailedCaseCollectionObject -runid $_) | %{
            $fids += $_.TestCaseId
            
            if( -not $hash_id_errormsg.Contains($_.TestCaseId)){
                $hash_id_errormsg.Add($_.TestCaseId,$_.ErrorMessage)
            }            
        }
        New-Variable -Name "f$prefix" -Value $fids
        $prefix++
        
        # for rendering use, we collect build numbers here
        $build_numbers += (Get-TestRunObject -runid $_).BuildNumber
		
		$build_envs += (Get-EnvironmentName -testrunid $_)
        $build_settings += (Get-TestRunObject -runid $_).TestSettings.Name
        $starts += (Get-TestRunObject -runid $_).DateCreated
        $completes += (Get-TestRunObject -runid $_).DateCompleted
    }
    
    # find the min start time
    $starts = @() + ($starts | sort)
    $start = $starts[0]
    
    # find the max completed time
    $completes = @() + ($completes | sort)
    $complete = $completes[$completes.Count-1]
	
    # add up all ids from each test run into one array, and select unique ids
    # so that we can get all the distinct ids of series runs
    $all_ids = @()
    For($i = 1; $i -lt $prefix; $i++){
        $all_ids += (Get-Variable -Name "a$i" -ValueOnly)
    }
    $all_ids_dist = $all_ids | select -uniq
    
    # add up all ids of Failed cases from each test run into one array
    # so that we can count how many times one id fails (appears)
    $all_failed_ids = @()
    For($i = 1; $i -lt $prefix; $i++){
        $all_failed_ids += (Get-Variable -Name "f$i" -ValueOnly)
    }
    $all_failed_ids_dist = $all_failed_ids | select -uniq
        
    
    $tar = @{} # faied ids and how many times it fails
    $occ = @{} # all ids and how many times it runs
        
    # initialize the hash table of $tar
    $all_ids_dist | %{
        $tar.Add($_,0)    
    }
    
    #count id and it runs
    $all_ids_dist | %{        
        $occ.Add($_,($all_ids -like $_).Count)
    }
    
    # start to count, use hash table to record: how many times has one failed case failed?
    if($all_failed_ids_dist -is [System.Array]){
        $all_failed_ids_dist | %{
            $tar[$_] = ($all_failed_ids -like $_).Count    
        }
    }
    elseif($all_failed_ids_dist -is [Int32]){
        # only one failed
        $tar[$all_failed_ids_dist] = ($all_failed_ids -like $all_failed_ids_dist).Count
    }
    # OK. Now the hash table $tar contains (key -> distinct id from all test runs; value -> failed-time of each id)

    if($failedonly){
        # filter $tar here if only to show failed ids
        $tar2 = @{}
        $tar.GetEnumerator() | %{
            if($_.value -ne 0){
                $tar2.Add($_.name,$_.value)
            }
        }
        $tar = $tar2        
    }
        
    #----------------------
    # draw pie chart
    #----------------------
    if(-not (Test-Path $config["localtemp"])){
        New-Item -ItemType Directory -Path $config["localtemp"] -Force
    }
    
    $uniq_name = ""
    $runids | %{
        $uniq_name += "$_" + "-"
    }
    $uniq_name = $uniq_name.TrimEnd('-')
     
    $uniq_name = "$uniq_name" + (Get-DateAppendix)
    
    $tempfilename = $config["localtemp"] + "\" + $uniq_name   + ".jpg"
    Draw-PieSeries -runids $runids -filename $tempfilename
    $pieaddress = $config["pieplace"] + "\" + $uniq_name + ".jpg"
    Copy-Item -Path $tempfilename -Destination $pieaddress -force
    
    #----------------------
    # create html content
    #----------------------
    $content = "<html>`n<head>`n"

    $content += "<style type=`"text/css`">`n"
    $content += "span.green { color:green }`n"
    $content += "span.red { color:red }`n"
    $content += "</style>`n"
    
    $content += "</head>`n<body>`n"
    
    $ids_inline = ""
    $runids | %{
        $ids_inline += "$_,"
    }
    $ids_inline = $ids_inline.TrimEnd(',')    
    $content += "<h2>Series Run: $ids_inline</h2>`n"
    
    $times = $runids.Count
    $content += "<h3>$times Run(s) combinated</h3>`n"
    
    # get start time
    $content += "Start time: $start<br/>`n"
    
    # get complete date
    $content += "Complete time: $complete<br/>`n"
    
    # get build number
	$build_numbers = $build_numbers | select -uniq
    $build_numbers_inline = ""
    $build_numbers | %{
        $build_numbers_inline += "$_,"
    }
    $build_numbers_inline = $build_numbers_inline.TrimEnd(',')
    
    $content += "Build Number: $build_numbers_inline<br/>`n"
    
	# get build env
    $build_envs = $build_envs | select -uniq
    $build_envs_inline = ""
    $build_envs | %{ $build_envs_inline += "$_," }
    $build_envs_inline = $build_envs_inline.TrimEnd(',')
    
    $content += "Build env: $build_envs_inline<br/>`n"
    
    # get build setting
    $build_settings = $build_settings | select -uniq
    $build_settings_inline = ""
    $build_settings | %{
        $build_settings_inline += "$_,"
    }
    $build_settings_inline = $build_settings_inline.TrimEnd(',')
    
    $content += "Build setting: $build_settings_inline<br/><br/>`n"	
	
    $content += "<div class=`"piechart`">`n"
    $content += "<img src=`"$pieaddress`" alt=`"pie_chart_of_this_run`" />`n"
    $content += "</div><br/>`n"
    
    $content += "<table border=`"1`">`n"
    $content += "<col /><col /><col width=`"500px`" /><col width=`"500px`" />`n"
    $content += "<tr bgcolor=`"C0C0C0`"><th>Fails/Runs</th><th>ID</th><th>Title</th><th>Error Message</th></tr>`n"
    
    # get data from the hash table
    $tar.GetEnumerator() | %{
        $failed_times = $_.value
        $run_times = $occ[$_.name]
        
        $tc_failrate = "$failed_times"+ "/"+ "$run_times"  # how many times it fails / how many times it runs
        $tc_id = $_.name
        $tc_title = $hash_id_title[$_.name]        
        $tc_error = $hash_id_errormsg[$_.name]        
               
        if($_.value -gt 0){
            $tc_failrate = "<span class=`"red`">$tc_failrate</span>"
        }else{
            $tc_failrate = "<span class=`"green`">$tc_failrate</span>"
        }
        
        $content += "<tr><td>$tc_failrate</td><td>$tc_id</td><td>$tc_title</td><td>$tc_error</td></tr>`n"
    }
    
    $content += "</table><br/>`n"
    $content += "Powered by Power MTM Shell (http://toolbox/pms). Do not reply.<br/>`n"
    $content += "Wicresoft MSIT SE Team, 2013`n"
    $content += "</body>`n</html>"
    
    $savehtml = $config["localtemp"] + "\$uniq_name.html"
    $content > $savehtml
    
    return $savehtml
}

Function Get-DateAppendix{
    (Get-Date).tostring("_yyyy-MM-dd_HH-mm-ss")
}