# get list of phones
$lookup = import-csv C:\git\PbxMigration\Lookup.csv
$Server = "YOUR PBX"
$Yealink = get-content C:\git\PbxMigration\ip.txt


# Invoke Selenium into script
$env:PATH += ";C:\GIT\PbxMigration" # Adds the path for ChromeDriver.exe to the environmental variable 
Add-Type -Path "C:\GIT\PbxMigration\WebDriver.dll" # Adding Selenium's .NET assembly (dll) to access it's classes in this PowerShell session


foreach ($Yealink in $Yealink)

{
$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver 

try
{


$YourURL = "http://" + $Yealink 
$adminUrl = $YourURL + "/servlet?m=mod_data&p=status&q=load"

# Login
$ChromeDriver.Navigate().GoToURL($adminUrl)
Start-Sleep -Seconds 1
$ChromeDriver.FindElementByName("username").SendKeys("YOUR USER NAME")
$ChromeDriver.FindElementByName("pwd").SendKeys("YOUR PASSWORD")
$ChromeDriver.FindElementById("idConfirm").Click() 
#Go to account 1 
Start-Sleep -Seconds 1

$AccUrl = $YourURL +  "/servlet?p=account-register&q=load&acc=0"
$ChromeDriver.Navigate().GoToURL($AccUrl)
#get the current user
$OldExtn = $ChromeDriver.FindElementByName("AccountRegisterName").GetAttribute("Value")

#check lookup for new extn
$NewExtn = $lookup.Where({$PSItem.OLD -eq $OldExtn}).NEW
$LABEL = $lookup.Where({$PSItem.OLD -eq $OldExtn}).LABEL

#Go to the second account page
$Acc2Url = $YourURL + "/servlet?m=mod_data&p=account-register&q=load&acc=1"
$ChromeDriver.Navigate().GoToURL($Acc2url)

$isEnabled = $null
$isEnabled = $ChromeDriver.FindElementByName("AccountRegisterName").GetAttribute("Value")
If ($isEnabled -match '^[a-z0-9]') { throw "Account already enabled" | out-file C:\Git\PbxMigration\error.txt -Append }


#enable the account 
$ChromeDriver.FindElementByName("AccountEnable").SendKeys([OpenQA.Selenium.Keys]::DOWN)
#Set Label
$ChromeDriver.FindElementByName("AccountLabel").Clear()
$ChromeDriver.FindElementByName("AccountLabel").Sendkeys($LABEL)
#Set Displayname 
$ChromeDriver.FindElementByName("AccountDisplayName").Clear()
$ChromeDriver.FindElementByName("AccountDisplayName").Sendkeys($LABEL)
#Set AuthName 
$ChromeDriver.FindElementByName("AccountRegisterName").Clear()
$ChromeDriver.FindElementByName("AccountRegisterName").Sendkeys($Newextn)
#set Username
$ChromeDriver.FindElementByName("AccountUserName").Clear()
$ChromeDriver.FindElementByName("AccountUserName").Sendkeys($Newextn)
#set regpass
$pass = $NewExtn + "WTCbelfast"
$ChromeDriver.FindElementByName("AccountPassword").Clear()
$ChromeDriver.FindElementByName("AccountPassword").Sendkeys($pass)
#set server
$ChromeDriver.FindElementByName("server1").Clear()
$ChromeDriver.FindElementByName("server1").Sendkeys($Server)

$ChromeDriver.FindElementByName("btnSubmit").Click() 
}

catch [System.Net.WebException],[System.Exception]
{
    Write-Host "Error", "$PSItem", "$Yealink","$NewExtn"
    $exception =  "Error $Yealink $NewExtn"  
    $exception | out-file C:\Git\PbxMigration\error.txt -Append

}

Function Stop-ChromeDriver {Get-Process -Name chromedriver -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue}
$ChromeDriver.Close()
$ChromeDriver.Quit()
Stop-ChromeDriver
}
# Cleaning up 






