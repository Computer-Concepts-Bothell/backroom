#Feb2023 -- Dakotam@conceptsnet.com
#Auto Updater Script
#remotescript is linked to the github i use for this tool, if you are using a different branch or git, update this to yours. 
$ToolLink = "https://raw.githubusercontent.com/Pixelbays/backroom/main/backroom.ps1"
#this is 'tool' name. mostly used in changing the file name in the update part
$ToolName = "1-Backroom"
#Auto Updater Script
try {
    #Current Version. Make sure to update before pushing.
    $Version = "1.4"
    $headers = @{ "Cache-Control" = "no-cache" }   
    $RemoteScript = (Invoke-WebRequest -Uri $ToolLink -Headers $headers -UseBasicParsing).Content
    $RemoteVersion = ($RemoteScript -split '\$version = "')[1].split('"')[0]
    #if the versions between local and github dont match. it will prompt for update and backup.
    #should be most self explaintory here
    if($Version -lt $RemoteVersion){
        $UpdateFound = Read-Host "Current Version $Version is out date! Would you like to update to $RemoteVersion ? y/n"
        if ($UpdateFound -eq "y") {
            #found that just having the script back itself up can be usefull. 
            $BackupRequest = Read-Host "Would you like to backup the current script? y/n"
            $UpdateRequest = "y"
        }
        if ($BackupRequest -eq "y") {
            #renames the current script file to the tool back with version number
            Rename-Item -Path .\$ToolName.ps1 -NewName "$ToolName-$Version-backup.ps1" -Force
            Write-Output "The Current Script has been renamed to '$ToolName-$Version-backup.ps1'"
        }
        if ($UpdateRequest -eq "y") {
            # download the new version if the version is different
            (Invoke-WebRequest -Uri $ToolLink -UseBasicParsing).Content | Out-File .\$toolname.ps1
            Write-Output "Please Close this script and open the updated version"
            Read-Host -Prompt "Press any key to reload the script"
            . .\$ToolName.ps1
        }
    }
    if ($Version -eq $RemoteVersion) {
        Write-Output "Current Version:$Version. is up to date!"
    }
    if ($Version -gt $RemoteVersion) {
        Write-Output "Current Version:$Version. must be a dev build, prod build is $RemoteVersion"
    }
}
catch {
    #if github was not reachable either due to internet or user error will fail. 
    Write-Output "Unable to check for update. Current Version:$Version"
}

#These Vars are editable if you need to change the subdomain or API. To change them go to the varibles.xml and change the data there. hoping to have this editable in the script using c
#First time setup. Checks if the variables.xml is real. if fails, Asks the user questions about their shopr and saves them to an external file.

try {
    $CFiles = Import-Clixml -Path .\variables.xml
} catch {
    Write-Output "Hello! Welcome to the PowerShell Repairshopr API Inventory tool!"
    Write-Output "I'll ask a couple of questions, save them Locally in a file called variables.ps1."
    Write-Output "This way I dont know them, and you don't know mine!"
    Write-Output "keep varibles.xml in the same folder as the main script otherwise it wont work, and will be talking again."
    Write-Output $Spacer
    $CFiles = [pscustomobject]@{
        CName = Read-Host "What is your Company Name?"
        SubDom = Read-Host "What is the sub domain you have at repairshopr? Example '*****.repairshopr.com'"
        APIKey = Read-Host "What is the API key you have made for this? Make sure it has ONLY The following permissions List/search, Edit, ViewCost"
    }
    $CFiles | Export-Clixml -Path .\variables.xml
    $CFiles = Import-Clixml -Path .\variables.xml
}
try {
    $LFiles = Import-Clixml -Path .\translated.xml
} catch {
    Write-Output "Hello! Welcome to the PowerShell Repairshopr API Inventory tool!"
    Write-Output "I'll ask a couple of questions, save them Locally in a file called variables.ps1."
    Write-Output "This way I dont know them, and you don't know mine!"
    Write-Output "keep translated.xml in the same folder as the main script otherwise it wont work, and will be talking again."
    Write-Output $Spacer
    Write-Output "the creation of this file must be done by you atm tool is not made to add this part... yet?"
    Write-Output "Please Create this XML file with the correct formating"
    Write-Output $Spacer
}
#These Vars setup the whole Script Please don't edit
$APIKey = $CFiles.APIKey
$SubDom = $CFiles.SubDom
$CName = $CFiles.CName 
#gets the date and changes to the formate yyMMdd so that we can use that for sort order 
$DateString = (Get-Date -UFormat %y%m%d)
#setup the vars for the API requests
$postheaders = @{Authorization = "Bearer $APIKey"
"Accept" = "application/json"}
$contenttype = "application/json"
#This Var is to use for the if statments to ignore when a cmd has been typed
$IgnoredInputs = "n", "c", "s", "o", "help", "export", "r"
$IgnoredPower = "131444", "131448", "131449", "149103", ""
$ChangeStatus = "New", "Need to order parts", "Waiting on Customer", "Waiting for Parts", "Unable to Contact"
#format spacer reused so i dont have to copy and paste the same bit or count. Lazness pays off now. 
$Spacer = "_______________"
#This var is for the product IDs to get saved when needed.
#$SavedList = @()
#creates the UPC var for later use
$TicketNum = ""
#this is a var to check if the user has selected that repair shopr has been opened and logged into. 
#$Signedin = ""
#the request user input saying so i dont have to copy and paste the same bit or count. Lazness pays off now.
$CSay = "Type/Scan a Ticket Number"
#this is the var that is setup for logs, and finally some user input!
<# 
$ELogs = [Ordered]@{
    Company = $CName
    Name = Read-Host "What is your Name?"
    Date = (Get-Date -Format "yyyy-MM-dd@HH.mm")
    NumProdScanned = 0
    NumProdSaved = 0
    UPCsNotFound = 0
    ScannedProds = @()
    SavedProdID = @()
}
#>

# Create a new SpVoice objects
Add-Type -AssemblyName System.speech

$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer

# Set the speed - positive numbers are faster, negative numbers, slower
$voice.rate = 0.75
$voice.SelectVoice("Microsoft Zira Desktop")




do {
    $Continue = Read-Host -Prompt "$CSay"
    $CFiles = Import-Clixml -Path .\variables.xml
    $LFiles = Import-Clixml -Path .\translated.xml
    if ($Continue -notin $IgnoredInputs) {
        try {
        #changes the Ticketnum var to the user input.
        $TicketNum = $Continue
        #requests the API for the Ticket using the given Ticket Number. If nothing is found with that Ticket Number, it will tell you. 
        $Request = Invoke-WebRequest -Uri "https://$SubDom.repairshopr.com/api/v1/tickets?number=$TicketNum" -ContentType $contenttype -Headers $postheaders
        #converts it to poowershell vars from jason
        $Response = $Request.Content | ConvertFrom-Json
        #if the server reponded we start to get the info we need and displays it.
        if (($Response.tickets).count -ne 0) {
            #Making vars for the things we care about from the reponse from the API.
            $Properties = $Response.tickets[0].properties 
            $TicketCustomer = $Response.tickets[0].customer_business_then_name
            $TicketID = $Response.tickets[0].id
            $TicketStatus = $Response.tickets[0].status
            #Prints those vars out to the user. So that they can double check what they see matches.
            Write-Output $Spacer
            Write-Host "Ticket ID: $TicketID"
            Write-Host "Ticket Number: $TicketNum"
            Write-Host "Ticket Customer: $TicketCustomer"
            Write-Host "Ticket Status: $TicketStatus"
            Write-Output "Ticket Properties: $Properties"
            Write-Output $Spacer
            $TranslatedLocation = $LFiles[$Properties.Location]
            Write-Output "Ticket Location: $TranslatedLocation"
            $TranslatedPower = $LFiles[$Properties."Power Supply"]
            Write-Output "Ticket Power Supply: $TranslatedPower"
            $voice.speak("Status is $TicketStatus") |Out-Null
            $Power = $Properties."Power Supply"
            if ($Power -notin $IgnoredPower) {
                $voice.speak("Make Sure $TranslatedPower is with the Ticket, Then Scan Location") |Out-Null
            }
            if ($Power -eq "131444") {
                $voice.speak("Powersupply is labeled not here yet, fix this, then Scan Location") |Out-Null
            } 
            if ($Power -in $IgnoredPower){
                $voice.speak("Scan the Location") |Out-Null
            }
            Write-Output $Spacer
            $NewLocation = Read-Host "Please Scan New location or cancel"
            Write-Output $Spacer
            $Properties.Location = $NewLocation
            $NewTranslate = $LFiles[$Properties.Location]
            if ($LFiles.ContainsKey($NewLocation)) {
                Write-Output "This Location is Real!"
                if ($TicketStatus -eq "In Progress") {
                    #converts back to json then pushes that date change to the API using the sort order field
                    $Body = @{"status" ="New"; "properties" = $Properties}
                    $jsonBody = $body | ConvertTo-Json
                    Invoke-RestMethod -Method PUT -Uri "https://$SubDom.repairshopr.com/api/v1/tickets/$TicketID" -ContentType $contenttype -Headers $postheaders -Body $jsonBody | Out-Null
                    Write-Output "Ticket Has been Moved To $NewTranslate and Has been removed from In Progress"
                    # Say something
                    $voice.speak("Location Updated to $NewTranslate, Status Changed") |Out-Null
                    Write-Output $Spacer
                }
                if ($TicketStatus -ne "In Progress") {
                    #converts back to json then pushes that date change to the API using the sort order field
                    $Body = @{"properties" = $Properties}
                    $jsonBody = $body | ConvertTo-Json
                    Invoke-RestMethod -Method PUT -Uri "https://$SubDom.repairshopr.com/api/v1/tickets/$TicketID" -ContentType $contenttype -Headers $postheaders -Body $jsonBody | Out-Null
                    Write-Output "Ticket Has been Moved To $NewTranslate"
                    Write-Output $Spacer
                    # Say something
                    $voice.speak("Location Updated to $NewTranslate") |Out-Null
                }
            }if ($NewLocation -in $ChangeStatus ) {
                #converts back to json then pushes that date change to the API using the sort order field
                $Body = @{"status" = "$NewLocation"}
                $jsonBody = $body | ConvertTo-Json
                Invoke-RestMethod -Method PUT -Uri "https://$SubDom.repairshopr.com/api/v1/tickets/$TicketID" -ContentType $contenttype -Headers $postheaders -Body $jsonBody | Out-Null
                Write-Output "Ticket Has been Moved To Status $NewLocation"
                Write-Output $Spacer
                # Say something
                $voice.speak("Status Updated to $NewLocation") |Out-Null
            }
            
            else {
                Write-Output "Location not found!"
                Write-Output $Spacer
            }

            #this is were the ticket number fails at
        }elseif ($Response.tickets.Count -eq 0 -and $Continue -notin $IgnoredInputs) {
            Write-Output $Spacer
            Write-Host "Ticket not found"
            $voice.speak("Ticket not found") |Out-Null
            Write-Output $Spacer
        }
        #had issues with wifi sometimes on our "location" device. in order to resolve that issue and it using the last scanned ticket not the new scanned ticket. 
        $Properties = ""
        $TicketCustomer = ""
        $TicketID = ""
        $TicketStatus = ""
        $TicketNum = ""
        
    }
    catch {
        Write-Output $Spacer
        Write-Output "Something went wrong, check internet connection"
        $voice.speak("Something went wrong, check internet connection") |Out-Null
        Write-Output $Spacer
    }
    }
    
        

<#
 if ($Continue -eq "c"){
    #This is the area to Change the saved vars in varibles 
    $SettingsSay = "Type your requested change, APIKey, CompanyName, Subdomain, MStock. Type N to Cancel"
    Write-Output $Spacer
    Write-Output "what do you want to change?"
    Write-Output $Spacer
    do  {
        $ChangedYes = 0 
        if ($ChangingVar -eq "APIKey") {
            Write-Output $Spacer
            $CFiles.APIKey = Read-Host "What is the API key you have made for this? Make sure it has ONLY The following permissions List/search, Edit, ViewCost"
            $ChangedYes =+ 1
            Write-Output $Spacer
        }
                           
        $CFiles | Export-Clixml -Path .\variables.xml    
        $ChangingVar = Read-Host $SettingsSay
        if ($ChangingVar -eq "n" -and $ChangedYes -ne "0"){
            #Write-Output "Please Close this script and open the updated version"
            Write-Output $Spacer
            Read-Host -Prompt "Press any key to reload the script"
            Start-Process powershell -ArgumentList "-File `".\1-inventory.ps1`"" -NoNewWindow
            Write-Output $Spacer
            Exit
        }
        if ($ChangingVar -eq "n"){
            $ChangingVar = "f"
            Write-Output $Spacer
        }
    }while($ChangingVar -ne "f")
}
if ($Continue -eq "help"){
    #This area is to inform what this tool can do.
    Write-Output $Spacer
    Write-Output "Hello and welcome to $CName PowerShell Repairshopr Inventory Tool! "
    Write-Output $Spacer
    Write-Output "Using this tool we are able to quickly get zones of inventory checked out without using the slow website."
    Write-Output "Anything saved will be lost once the script is closed. If you have saved Products to look at later using 's' make sure to open them using 'o' before closing the script."
    Write-Output "While using this tool, if the qty in front of you does not match up with the qty shown, please find out why."
    Write-Output "Need to save it? you can export the data you've taken"
    Write-Output "The Date format is yyMMdd. This is so it plays nicer with the Sort Order on the website. Since that that is the column we are high jacking for the scanned date."
    Write-Output "Type 'cmds' to get the list of valid commands"
    Write-Output "'credits' - Typing credits will bring up credits of this script"
    Write-Output $Spacer 
}
if ($Continue -eq "cmds"){
    #this is the area for what cmds are ready for the user to use
    Write-Output $Spacer
    Write-Output "Here are valid commands that you can do with this tool."
    Write-Output "Scanning/Typing the UPC on a Product using a scanner will bring up the Product ID, Name, Retailed Price, Expected Qty, and last scanned date."
    Write-Output "'s' - Typing S will save the current product ID that you last scanned to a list that you can open at a later time to evaluate it on the website."
    Write-Output "'o' - Typing O will open the saved products to their product page in repairshopr, Make sure you are already signed in or it will just open a lot of sign in pages."
    Write-Output "'n' - Typing N will close the script out."
    Write-Output "'export'- Typing export will export a json file with usefull info. like who, date, links what products were scanned, saved, total amount of scanned/saved."
    Write-Output "'r' - Typing R will reload the script."
    Write-Output "'c' - Typing C will let you make changes to the Company Varibles saved during first time setup."
    Write-Output "'e' - Typing E will expand the current product scanned. Opening the link for just that product instead of all the saved ones."
    Write-Output $Spacer
}
#>
    if ($Continue -eq "r"){
        Write-Output $Spacer
        $RConfirm = Read-Host "Are you sure you would like to reload? Y/N"
        Write-Output $Spacer
        if ($Rconfirm -eq "y") {
            #Start-Process powershell -ArgumentList "-File `".\backroom.ps1`"" -NoNewWindow
            Write-Output $Spacer
            Write-Output $Spacer
            Write-Output $Spacer
            . .\backroom.ps1
            #Exit
        }
    }   
    if ($Continue -eq "credits"){
        Write-Output "This Tool is a open source Powershell tool created by Dakota @ pixelbays on github. If you have any problems or issues make and issue on github."
    }   


}while ($Continue -ne "n")        #Exit
