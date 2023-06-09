#Feb2023 -- Dakotam@conceptsnet.com
#Auto Updater Script
#remotescript is linked to the github i use for this tool, if you are using a different branch or git, update this to yours. 
$ToolLink = "https://raw.githubusercontent.com/Pixelbays/backroom/main/backroom.ps1"
#this is 'tool' name. mostly used in changing the file name in the update part
$ToolName = "1-Backroom"
#Auto Updater Script
try {
    #Current Version. Make sure to update before pushing.
    $Version = "1.7.0"
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
            RestartScript
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
        TeamsWebHook = Read-Host "What is the Teams WebHook key you have made for this?"
    }
    $CFiles | Export-Clixml -Path .\variables.xml
    $CFiles = Import-Clixml -Path .\variables.xml
}
try {
    $LFiles = Import-Clixml -Path .\translated.xml
} catch {
    Write-Output "Hello! Welcome to the PowerShell Repairshopr API Inventory tool!"
    Write-Output "I'll ask a couple of questions, save them Locally in a file called translated.xml."
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
$TeamsWebhook = $CFiles.TeamsWebhook 
#setup the vars for the API requests
$postheaders = @{Authorization = "Bearer $APIKey"
"Accept" = "application/json"}
$contenttype = "application/json"
$BaseURL = "https://$SubDom.repairshopr.com"
#This Var is to use for the if statments to ignore when a cmd has been typed
$IgnoredInputs = "n", "settings", "s", "cmds", "help", "export", "r", "teams push"
$IgnoredPower = "131444", "131448", "131449", "149103", ""
$ChangeStatusWord = "New", "Need to order parts", "Waiting on Customer", "Waiting for Parts", "Unable to Contact" 
$ChangeStatusTrans = @{
"101" = "New"
"206" = "Need to order parts"
"418" = "Waiting on Customer"
"404" = "Waiting for Parts"
"444" = "Unable to Contact"
}
#format spacer reused so i dont have to copy and paste the same bit or count. Lazness pays off now. 
$Spacer = "_______________"
#creates the UPC var for later use
$TicketNum = ""
#the request user input saying so i dont have to copy and paste the same bit or count. Lazness pays off now.
$CSay = "Scan Location or Status"

# Create a new SpVoice objects
Add-Type -AssemblyName System.speech

$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer

# Set the speed and model - positive numbers are faster, negative numbers, slower

$voice.rate = 0.75
$voice.SelectVoice("Microsoft Zira Desktop")

#Function API call for tickets, (Should be able to make this more modular)
function TAPICall {
    param(
        [string]$TicketNum
    )
    #requests the API for the Ticket using the given Ticket Number. If nothing is found with that Ticket Number, it will tell you. 
    $Request = Invoke-WebRequest -Uri "$BaseURL/api/v1/tickets?number=$TicketNum" -ContentType $contenttype -Headers $postheaders
    #converts it to poowershell vars from jason
    return $Response = $Request.Content | ConvertFrom-Json

}
#Function API Push for tickets, (Should be able to make this more modular)
function TAPIPush {
    param(
        [string]$TicketID,
        [array]$Body,
        [string]$Change
    )
    #converts var back to json to push to API
    $jsonBody = $Body | ConvertTo-Json
    #Write-Output $jsonBody
    Invoke-RestMethod -Method PUT -Uri "$BaseURL/api/v1/tickets/$TicketID" -ContentType $contenttype -Headers $postheaders -Body $jsonBody | Out-Null
    if ($Change -eq "1"){
        Write-Output "Ticket Has been Moved To $NewTranslate and Has been removed from In Progress"
        # Say something
        $voice.speak("Location Updated to $NewTranslate, Status was changed") |Out-Null
    }
    if ($Change -eq "2"){
        Write-Output "Ticket Status updated to $Statustranslated"
        # Say something
        $voice.speak("Ticket Status updated to $Statustranslated") |Out-Null
    }else {   
        Write-Output "Ticket updated to $NewTranslate"
        # Say something
        $voice.speak("Location Updated to $NewTranslate") |Out-Null
    }
    Write-Output $Spacer
}

function TeamsPush {
    param(
        [string]$Header,
        [array]$Message
    )
    # Create the message payload
    $message = @{
        "@type" = "MessageCard"
        "@context" = "http://schema.org/extensions"
        "themeColor" = "0072C6"
        "summary" = "PowerShell Webhook Test"
        "sections" = @(
            @{
                "activityTitle" = "$Header"
                "activitySubtitle" = "$Message"
            }
        )
    }

    # Convert the message to JSON
    $jsonMessage = $message | ConvertTo-Json

    # Send the message via webhook
    Invoke-RestMethod -Method Post -Uri $TeamsWebhook -Body $jsonMessage -ContentType 'application/json'
}
function RestartScript {
    Write-Output $Spacer
    $RConfirm = Read-Host "Are you sure you would like to reload? Y/N"
    Write-Output $Spacer
    if ($Rconfirm -eq "y") {
        Write-Output $Spacer
        Write-Output $Spacer
        Write-Output $Spacer
        . .\$ToolName.ps1
        #Exit
    }
}
do {
    $Continue = Read-Host -Prompt "Scan or Type Ticket Number"
    $CFiles = Import-Clixml -Path .\variables.xml
    $LFiles = Import-Clixml -Path .\translated.xml
    if ($Continue -notin $IgnoredInputs) {
        try {
        #changes the Ticketnum var to the user input.
        $TicketNum = $Continue
        $Response = TAPICall -TicketNum $TicketNum
        if (($Response.tickets).count -ne 0) {
            #Making vars for the things we care about from the reponse from the API.
            $Properties = $Response.tickets[0].properties 
            $TicketCustomer = $Response.tickets[0].customer_business_then_name
            $TicketID = $Response.tickets[0].id
            $TicketStatus = $Response.tickets[0].status
            $TranslatedLocation = $LFiles[$Properties.Location]
            $TranslatedPower = $LFiles[$Properties."Power Supply"]
            #Prints those vars out to the user. So that they can double check what they see matches.
            Write-Output $Spacer
            Write-Host "Ticket Number: $TicketNum"
            Write-Host "Ticket Customer: $TicketCustomer"
            Write-Host "Ticket Status: $TicketStatus"
            Write-Output "Ticket Properties: $Properties"
            Write-Output $Spacer
            Write-Output "Ticket Location: $TranslatedLocation"
            Write-Output "Ticket Power Supply: $TranslatedPower"
            $voice.speak("Status is $TicketStatus") |Out-Null
            #Checks what the Powersupply status is and reads the correct response
            $Power = $Properties."Power Supply"
            if ($Power -notin $IgnoredPower) {
                $voice.speak("Make Sure $TranslatedPower is with the Ticket, Then $CSay") |Out-Null
            }
            if ($Power -eq "131444") {
                $voice.speak("Powersupply is labeled not here yet, fix this, Then $CSay") |Out-Null
            } 
            if ($Power -in $IgnoredPower){
                $voice.speak("$CSay") |Out-Null
            }
            #gets user input for the new status or location
            Write-Output $Spacer
            $NewLocation = Read-Host "Please $CSay"
            Write-Output $Spacer
            #updates the ticket var to the new location
            $Properties.Location = $NewLocation
            $NewTranslate = $LFiles[$Properties.Location]
            #Checks if the user input is location status or not found
            if ($LFiles.ContainsKey($NewLocation)) {
                #Write-Output "This Location is Real!"
                #checks if the ticket is left in progres or not. 
                if ($TicketStatus -eq "In Progress") {
                    #This block will push to API changing ticket to new and update the location
                    #converts var back to json to push to API
                    $Body = @{"status" ="New"; "properties" = $Properties}
                    TAPIPush -TicketID $TicketID -Body $Body -Change "1"
                }
                if ($TicketStatus -ne "In Progress") {
                    #This block will push to API changing ticket updating the location
                    #converts var back to json to push to API
                    $Body = @{"properties" = $Properties}
                    TAPIPush -TicketID $TicketID -Body $Body -Change "0"
                }
            }else {
                $NotLocation = 1
            }
            #checks if user input is a status code start the block to change the status
            #first block if for barcode
            if ($ChangeStatusTrans.ContainsKey($NewLocation)) {
                #This block will push to API changing ticket updating the status
                #converts var back to json to push to API
                Write-output "Changing Status Via Numbers"
                $Statustranslated = $ChangeStatusTrans.$NewLocation
                $Body = @{"status" = "$Statustranslated"}
                TAPIPush -TicketID $TicketID -Body $Body -Change "2"
                $NotLocation = 0
            #second block is for typing
            }if ($NewLocation -in $ChangeStatusWord) {
                #This block will push to API changing ticket updating the location
                #converts var back to json to push to API
                $Statustranslated = $Newlocation
                $Body = @{"status" = "$NewLocation"}
                Write-output "Changing Status Via Words"
                TAPIPush -TicketID $TicketID -Body $Body -Change "2"
                $NotLocation = 0
            }
            if ($Newlocation -eq "n") {
                $Continue = "n"
            }
            if ($Newlocation -eq "r") {
                RestartScript
            }
            if ($Newlocation -eq "teams push"){
                $Header = "Backroom Script Sent this"
                $Message = "https://$SubDom.repairshopr.com/tickets/$TicketID"
                TeamsPush -Header $Header -Message $Message |Out-Null
                $voice.speak("Message Pushed to teams") |Out-Null
                Write-Output $Spacer
                Write-Output "Message Pushed to teams"

            }
            #if the user input is not found in locations or status, let user know
            if ($NewLocation -notin $ChangeStatusWord -and $NotLocation -eq 1 -and $NewLocation -notin $IgnoredInputs) {
                $voice.speak("Input not found") |Out-Null
                Write-Output $Spacer
                Write-Output "Location or status not found!"
                Write-Output $Spacer
                $NotLocation = 0
            }
            #If API response no tickets let user know
        }elseif ($Response.tickets.Count -eq 0 -and $Continue -notin $IgnoredInputs) {
            Write-Output $Spacer
            Write-Host "Ticket not found"
            $voice.speak("Ticket not found") |Out-Null
            Write-Output $Spacer
           
        }
        #had issues with wifi sometimes on our "location" device. in order to resolve that issue for wifi and it using the last scanned ticket not the new scanned ticket. Also Just use ethernet will fix this.
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
 if ($Continue -eq "Settings"){
    #This is the area to Change the saved vars in varibles 
    $SettingsSay = "Type your requested change, 'APIKey', 'CompanyName', 'Subdomain', 'Teams Webhook' Type N to Cancel"
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
        if ($ChangingVar -eq "CompanyName") {
            Write-Output $Spacer
            $CFiles.Cname = Read-Host "What is your Company Name?"
            $ChangedYes =+ 1
            Write-Output $Spacer
        }
        if ($ChangingVar -eq "Subdomain") {
            Write-Output $Spacer
            $CFiles.Subdom = Read-Host "What is the sub domain you have at repairshopr? Example '*****.repairshopr.com'"
            $ChangedYes =+ 1
            Write-Output $Spacer
        }
        if ($ChangingVar -eq "Teams Webhook") {
            Write-Output $Spacer
            $CFiles.TeamsWebHook = Read-Host "What is the Teams WebHook key you have made for this?"
            $ChangedYes =+ 1
            Write-Output $Spacer
        }              
        $CFiles | Export-Clixml -Path .\variables.xml    
        $ChangingVar = Read-Host $SettingsSay
        if ($ChangingVar -eq "n" -and $ChangedYes -ne "0"){
            #Write-Output "Please Close this script and open the updated version"
            Write-Output $Spacer
            Read-Host -Prompt "Press any key to reload the script"
            RestartScript
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
    Write-Output "Using this tool we are able to quickly update a ticket location, status, or push a link to teams by scanning the ticket barcode then the barcode for your actionr"
    Write-Output "Type 'cmds' to get the list of valid commands"
    Write-Output "'credits' - Typing credits will bring up credits of this script"
    Write-Output $Spacer 
}
if ($Continue -eq "cmds"){
    #this is the area for what cmds are ready for the user to use
    Write-Output $Spacer
    Write-Output "Here are valid commands that you can do with this tool."
    Write-Output "Scanning/Typing the Ticket Number will allow you to push a new ticket location, status, or push to teams"
    Write-Output "'n' - Typing N will close the script out."
    Write-Output "'r' - Typing R will reload the script."
    Write-Output "'settings' - Typing C will let you make changes to the Company Varibles saved during first time setup."
    Write-Output $Spacer
}
#>
    if ($Continue -eq "r"){
        RestartScript
    }   
    if ($Continue -eq "credits"){
        Write-Output "This Tool is a open source Powershell tool created by Dakota @ pixelbays on github. If you have any problems or issues make and issue on github."
    }   
}while ($Continue -ne "n")        #Exit