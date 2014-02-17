#######################################################
##
## Change bulky parameter values of test cases in MTM
## (Powered by PMS)
## dave wu, Oct. 23, 2013
##
#######################################################
param(
[string]$new_tiles_value = "new value",
[string]$new_oem_value   = "new value"
)

$execPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
. $execPath\lib_infra.ps1

# get team project
$tp = Get-TeamProjectObject

# get all the test cases of specified query (WIQL)
$cases = $tp.TestCases.Query("SELECT * FROM workitems WHERE [Work Item Type] = 'Test Case' AND [Assigned To] = 'somebodys display name'")

# for each case, if they has parameter named Tiles(or OEM), then change it
# note: I assume that parameters are all in 'defaultTable' (it seems so), and these could be 
#       multiple rows in the table containing multiple target parameters value. I change them all with same value.
#       you can DIY here to change what you really want to change 
$cases | %{    
    $_.WorkItem.Open()
    
    $_.DefaultTable.Rows | %{
        if($_.Tiles -ne $null){                
            $_.Tiles = $new_tiles_value        
        }

        if($_.OEM -ne $null){
            $_.OEM = $new_oem_value
        }
    }    
    
    $_.save()
    $_.WorkItem.Close()
    
    write-host "changed " $_.id
}