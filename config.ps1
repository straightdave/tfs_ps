########################################################
##
## config:
## Default config
## Aug. 2013 Dave Wu
##
#######################################################

$config = @{

# Your TFS url
tfs       = "http://mytfssvr:8080/tfs/tpc1"

# Your team project name
teamproj  = "MyProject"

# Path of your binaries of TFS API DLLs
# You can install Visual Studio to gain these DLLs (at least installing VS - Team Explorer)
# below path is dlls in VS2013 (12.0), change to "11.0" if you use VS2012
binpath   = "C:\Program Files\Microsoft Visual Studio 12.0\Common7\IDE\ReferenceAssemblies\v2.0"
#binpath   = "C:\Program Files\Microsoft Visual Studio 11.0\Common7\IDE\ReferenceAssemblies\v2.0"

# Your build drop folder (when creating test run, it will load test container files etc. from it)
# Please refer to function:Create-TestRun
# The format is supposed to be "\\build_drop_folder\build_definition\each_build_number" as default in my project
# if you need to change, just change the code here and there (in that function)
dropfolder = "\\MyBuildServer\drop"

# the place where pie charts store (and also other output files)
# needs a shared folder on accessible server
# used in reporting
pieplace = "\\A_Shared_Server\sharedpic"
# local temp folder, should be accessible, to carry temp files generated locally
localtemp = "E:\pms_temp"

# mail gateway / sender identity
# You need to logon the machine with your account so that scripts can use your credential to send mail
smtphost = "smtp.mycompany.com"
mailfrom = "me@mycompany.com"

}

Function Set-ConfigValue{
    param(
    [string]$key,
    [string]$value
    )

    # change existing item value or add new item
    if($config[$key] -eq $null -or $config[$key] -eq ""){
        $config.Add($key,$value)
    }
    else{
        $config[$key] = $value
    }
}