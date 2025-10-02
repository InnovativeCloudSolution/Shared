<#

Mangano IT - DeskDirector - Update Form Question Options
Created by: Gabriel Nugent
Version: 1.0.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][int]$EntityId,
    [Parameter(Mandatory)][string]$FormQuestionName,
    [Parameter(Mandatory)][array]$DesiredChoices,
    [switch]$DropDown
)

## SCRIPT VARIABLES ##

$Result = $false

# Bearer token
$UserKey = .\DeskDirector-GetUserKey.ps1

# Converts supplied array to an ArrayList, allowing for values to be removed
[System.Collections.ArrayList]$DesiredChoicesAL = $DesiredChoices

# To be built out with the final list of choices
$NewChoices = @()

## GET FORM ##

$GetFormArguments = @{
    Uri = "https://portal.manganoit.com.au/api/v2/ddform/forms/$EntityId"
    Method = 'GET'
    ContentType = "application/json"
    Headers = @{
        Authorization = "DdAccessToken $UserKey"
    }
    UseBasicParsing = $true
}

try {
    $Form = Invoke-RestMethod @GetFormArguments
    Write-Warning "SUCCESS: Fetched form $EntityId."
    $Result = $true
} catch {
    Write-Error "Unable to fetch form $EntityId : $_"
    $Form = $null
}

## UPDATE CURRENT OPTIONS IN FORM ##

if ($null -ne $Form) {
    # Build new object from form
    $NewForm = @{
        form = @{
            entityId = $Form.form.entityId
            name = $Form.form.name
            description = $Form.form.description
            titleFormat = $Form.form.titleFormat
            fields = @()
            readOnly = $Form.form.readOnly
        }
    }

    # Look for question in form that matches the requested question
    foreach ($Field in $Form.form.fields) {
        if ($Field.name -eq $FormQuestionName) {
            # Sort through choices that are currently on the form
            foreach ($Choice in $Field.choices) {
                $ChoiceName = $Choice.name
                $ChoiceIdentifier = $Choice.identifier
                if ($DesiredChoicesAL.Contains($ChoiceName)) {
                    Write-Warning "$ChoiceName found in form question with identifier $ChoiceIdentifier."
                    $NewChoices += @{       # Add existing choice to new list
                        name = $ChoiceName
                        identifier = $ChoiceIdentifier
                    }  
                    $DesiredChoicesAL.Remove($Choice.name)      # Remove located choice from list
                }
            }

            # Add whatever new choices are left
            foreach ($Name in $DesiredChoicesAL) {
                Write-Warning "Adding $Name as a choice to the form question."
                $Identifier = -join ((48..57) + (97..122) | Get-Random -Count 6 | ForEach-Object {[char]$_})        # Generate random identifier
                $NewChoices += @{
                    name = $Name
                    identifier = $Identifier
                }
            }

            # Build new field object
            $NewField = @{      
                type = $Field.type
                name = $Field.name
                identifier = $Field.identifier
                required = $Field.required
                choices = $NewChoices | Sort-Object {$_.name}
            }

            # Add description if not null
            if ($null -ne $Field.description) {
                $NewField += @{
                    description = $Field.description
                }
            }

            # Make field dropdown if requested
            if ($DropDown) {
                $NewField += @{
                    meta = @{
                        render = "dropdown"
                    }
                }
            }

            # Add new field object to new form
            $NewForm.form.fields += $NewField
        }
        
        # Add existing field to list if it doesn't match the question
        else { $NewForm.form.fields += $Field }
    }
}

## PUSH NEW FORM TO DESKDIRECTOR ##

if ($null -ne $NewForm) {
    $UpdateFormArguments = @{
        Uri = "https://portal.manganoit.com.au/api/v2/ddform/forms/$EntityId"
        Method = 'PUT'
        ContentType = "application/json"
        Headers = @{
            Authorization = "DdAccessToken $UserKey"
        }
        Body = $NewForm | ConvertTo-Json -Depth 100
        UseBasicParsing = $true
    }

    try {
        Invoke-RestMethod @UpdateFormArguments | Out-Null
        Write-Warning "SUCCESS: Updated form with entity ID $EntityId."
        $Result = $true
    } catch {
        Write-Error "Unable to update form $EntityId : $_"
    }
}

## SEND OUTPUT TO FLOW ##

$Output = @{
    Result = $Result
}

Write-Output $Output | ConvertTo-Json