<#
.NOTES
    NAME: Send-Woof.ps1
    Type: PowerShell

        AUTHOR:  David Schulte
        DATE:    2022-06-20
        EMAIL:   celerium@celerium.org
        Updated:
        Date:

    VERSION HISTORY:
    0.1 - 2022-06-20 - Initial Release

    TODO:
    N\A

.SYNOPSIS
    Sends a dog image & fact to a Teams channel.

.DESCRIPTION
    The Send-Woof script sends a dog image & fact to a Teams channel using a Teams webhook connector URI.

    Various filters are in place to try and prevent any inappropriate images from being sent.

    An image is randomly selected from a random dog subreddit. Dog facts are pulled from the dog-facts-api.herokuapp.com API

    Unless the -Verbose parameter is used, no output is displayed.

.PARAMETER TeamsURI
    A string that defines where the Microsoft Teams connector URI sends information to.

.EXAMPLE
    .\Send-Woof.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....'

    Using the defined webhooks connector URI a random dog image & fact are sent to the webhooks Teams channel.

    No output is displayed to the console.

.EXAMPLE
    .\Send-Woof.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....' -Verbose

    Using the defined webhooks connector URI a random dog image & fact are sent to the webhooks Teams channel.

    Output is displayed to the console.

.INPUTS
    TeamsURI

.OUTPUTS
    Console, TXT

.LINK
    Celerium - https://www.celerium.org/
    Dog Facts - https://dog-facts-api.herokuapp.com

#>

<############################################################################################
                                        Code
############################################################################################>
#Requires -Version 5.0

#Region  [ Parameters ]

[CmdletBinding()]
param(
        [Parameter(ValueFromPipeline = $true, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$TeamsURI
    )

#EndRegion  [ Parameters ]

Write-Verbose ''
Write-Verbose "START - $(Get-Date -Format yyyy-MM-dd-HH:mm)"
Write-Verbose ''
Write-Verbose " - (1/3) - $(Get-Date -Format MM-dd-HH:mm) - Gathering Woof Data"

#Region     [ Prerequisites ]

    $Log = "C:\Celerium\Logs\Send-Woof-Report"
    $TXTReport = "$Log\Send-WoofLog.txt"

#EndRegion  [ Prerequisites ]

#Region  [ Main Code ]

try {

    $WoofSources = @(   'https://www.reddit.com/r/rarepuppers/.json?sort=top&t=week&limit=25' ,
                        'https://www.reddit.com/r/dogswithjobs/.json?sort=top&t=week&limit=25' ,
                        'https://www.reddit.com/r/WhatsWrongWithYourDog/.json?sort=top&t=week&limit=25' ,
                        'https://www.reddit.com/r/DOG/.json?sort=top&t=week&limit=25' ,
                        'https://www.reddit.com/r/dogpictures/.json?sort=top&t=week&limit=25',
                        'https://www.reddit.com/r/puppies/.json?sort=top&t=week&limit=25'
                )
        $WoofURI = Get-Random -InputObject $WoofSources

    $WoofData = Invoke-RestMethod -Uri $WoofURI -ErrorAction Stop
    $WoofImage = Get-Random $( $WoofData.data.children.data | Where-Object {$_.author -ne 'TrendingBot' -and $_.over_18 -eq $false -and $_.is_video -eq $false -and $_.url -notlike "*gallery*" -and $_.url -notlike "*v.redd*" -and $_.url -notlike "*gif*" -and $_.is_self -eq $false} )

    $WoofFact = (Invoke-RestMethod -Uri 'https://dog-facts-api.herokuapp.com/api/v1/resources/dogs?number=1' -ErrorAction Stop).fact

}
catch {
    Write-Error $_

    if ( (Test-Path -Path $Log -PathType Container) -eq $false ){
        New-Item -Path $Log -ItemType Directory > $null
    }

    (Get-Date -Format yyyy-MM-dd-HH:mm) + " - " + "[ Step (1/3) ]" + " - " + $_.Exception.Message | Out-File $TXTReport -Append -Encoding utf8

    exit
}

#EndRegion  [ Main Code ]

Write-Verbose " - (2/3) - $(Get-Date -Format MM-dd-HH:mm) - Formatting Woof Data"

#Region     [ Adjust for Puppies ]

$DeployPuppy = if ( $WoofURI -like "*puppies*" ){$true}else{$false}

    switch ($DeployPuppy){
        $true  {
            $TitleText = "The Weekly Woof! - !!! INCOMING PUPPY !!!"
            $TitleColor = "warning"
            $SubTitleText = "- I'll be the bestest boy ever! \r"
            $FactText = "Did you know: _$($WoofFact)_"
            $ImageHeight = 'auto'
        }
        $false {
            $TitleText = "The Weekly Woof!"
            $TitleColor = "accent"
            $SubTitleText = "- Borking at the evil vacuum monster since 1907 \r"
            $FactText = "Did you know: _$($WoofFact)_"
            $ImageHeight = '350px'
        }

    }

#EndRegion  [ Adjust for Puppies ]

Write-Verbose " - (3/3) - $(Get-Date -Format MM-dd-HH:mm) - Sending Woof Data"


#Region     [ Teams Code ]

$JSONBody = @"
{
    "type":"message",
    "attachments":[
    {
        "contentType":"application/vnd.microsoft.card.adaptive",
        "contentUrl":null,
        "content":{
            "$('$schema')":"http://adaptivecards.io/schemas/adaptive-card.json",
            "type":"AdaptiveCard",
            "version":"1.4",
            "body":[
                    {
                        "type": "TextBlock",
                        "size": "Large",
                        "weight": "Bolder",
                        "color": "$TitleColor",
                        "text": "$TitleText"
                    },
                    {
                        "type": "TextBlock",
                        "size": "Small",
                        "text": "$SubTitleText",
                        "isSubtle" : true
                    },
                    {
                        "type": "TextBlock",
                        "text": "$FactText",
                        "wrap": true
                    },
                    {
                        "type": "Image",
                        "url": "$($WoofImage.url)",
                        "altText": "$($WoofImage.title)",
                        "height": "$ImageHeight",
                        "width": "auto",
                        "msTeams": {
                            "allowExpand": true
                        }
                    }
                ],
                "actions": [
                    {
                        "type": "Action.OpenUrl",
                        "title": "Adopt a pupper",
                        "url": "https://www.hsbh.org/adopt-a-dog/#adopt-a-dog"
                    },
                    {
                        "type": "Action.OpenUrl",
                        "title": "Source",
                        "url": "https://reddit.com$($WoofImage.permalink)"
                    }
                ],
                "msTeams": {
                    "width": "Full"
                }
            }
        }
    ]
}
"@

try {

    Invoke-RestMethod -Uri $TeamsURI -Method Post -ContentType 'application/json' -Body $JsonBody -ErrorAction Stop > $null

}
catch {
    Write-Error $_

    if ( (Test-Path -Path $Log -PathType Container) -eq $false ){
        New-Item -Path $Log -ItemType Directory > $null
    }

    (Get-Date -Format yyyy-MM-dd-HH:mm) + " - " + "[ Step (3/3) ]" + " - " + $_.Exception.Message | Out-File $TXTReport -Append -Encoding utf8

    exit
}

#EndRegion  [ Teams Code ]

Write-Verbose ''
Write-Verbose "End - $(Get-Date -Format yyyy-MM-dd-HH:mm)"
Write-Verbose ''