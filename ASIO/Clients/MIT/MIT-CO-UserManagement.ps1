function OnboardingUser {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$userDetails
    )

    Write-MessageLog "Starting onboarding process for user $($userDetails.userPrincipalName)."

    try {
        $DinopassUrl = "https://www.dinopass.com/password/strong"
        $Password = Invoke-RestMethod -Uri $DinopassUrl

        $params = @{
            accountEnabled    = $true
            userPrincipalName = $userDetails.userPrincipalName
            displayName       = $userDetails.displayName
            mail              = $userDetails.userPrincipalName
            mailNickname      = $userDetails.mailNickname
            passwordProfile   = @{
                forceChangePasswordNextSignIn        = $true
                forceChangePasswordNextSignInWithMfa = $true
                password                             = $Password
            }
        }

        if ($userDetails.givenName) { $params.givenName = $userDetails.givenName }
        if ($userDetails.surname) { $params.surname = $userDetails.surname }
        if ($userDetails.jobTitle) { $params.jobTitle = $userDetails.jobTitle }
        if ($userDetails.department) { $params.department = $userDetails.department }
        if ($userDetails.mobilePhone) { $params.mobilePhone = $userDetails.mobilePhone }
        if ($userDetails.officeLocation) { $params.officeLocation = $userDetails.officeLocation }
        if ($userDetails.streetAddress) { $params.streetAddress = $userDetails.streetAddress }
        if ($userDetails.city) { $params.city = $userDetails.city }
        if ($userDetails.state) { $params.state = $userDetails.state }
        if ($userDetails.postalCode) { $params.postalCode = $userDetails.postalCode }
        if ($userDetails.country) { $params.country = $userDetails.country }

        $ExistingUser = Get-MgUser -Filter "UserPrincipalName eq '$($userDetails.userPrincipalName)'"
        if ($ExistingUser) {
            Write-ErrorLog "A user with UserPrincipalName $($userDetails.userPrincipalName) already exists."
            return
        }

        $NewUser = New-MgUser @params
        Write-MessageLog "User $($userDetails.userPrincipalName) created successfully."
    }
    catch {
        Write-ErrorLog "Failed to create user $($userDetails.userPrincipalName): $_"
        return
    }

    if ($userDetails.licenseGroupNames) {
        foreach ($groupName in $userDetails.licenseGroupNames) {
            try {
                $groupId = (Get-MgGroup -Filter "displayName eq '$groupName'").Id
                if ($groupId) {
                    New-MgGroupMember -GroupId $groupId -DirectoryObjectId $NewUser.Id
                    Write-MessageLog "User $($userDetails.userPrincipalName) added to license group '$groupName' with ID $groupId."
                }
                else {
                    Write-ErrorLog "Failed to find license group with name '$groupName'."
                }
            }
            catch {
                Write-ErrorLog "Failed to add user $($userDetails.userPrincipalName) to license group '$groupName': $($_.Exception.Message)"
            }
        }
    }

    if ($userDetails.groupsNames) {
        foreach ($groupName in $userDetails.groupsNames) {
            try {
                $groupId = (Get-MgGroup -Filter "displayName eq '$groupName'").Id
                if ($groupId) {
                    New-MgGroupMember -GroupId $groupId -DirectoryObjectId $NewUser.Id
                    Write-MessageLog "User $($userDetails.userPrincipalName) added to Microsoft 365 group '$groupName' with ID $groupId."
                }
                else {
                    Write-ErrorLog "Failed to find Microsoft 365 group with name '$groupName'."
                }
            }
            catch {
                Write-ErrorLog "Failed to add user $($userDetails.userPrincipalName) to Microsoft 365 group '$groupName': $_"
            }
        }
    }

    # Broken
    # if ($userDetails.sharedMailbox) {
    #     foreach ($sharedMailboxEmail in $userDetails.sharedMailbox) {
    #         try {
    #             $sharedMailboxId = (Get-MgUser -Filter "mail eq '$sharedMailboxEmail'").Id
    #             if ($sharedMailboxId) {
    #                 New-MgUserTransitiveMemberOf -UserId $NewUser.Id -DirectoryObjectId $sharedMailboxId
    #                 Write-MessageLog "Access granted to shared mailbox $sharedMailboxEmail with ID $sharedMailboxId."
    #             }
    #             else {
    #                 Write-ErrorLog "Failed to find shared mailbox with email '$sharedMailboxEmail'."
    #             }
    #         }
    #         catch {
    #             Write-ErrorLog "Failed to grant access to shared mailbox '$sharedMailboxEmail': $_"
    #         }
    #     }
    # }

    # Broken
    # if ($userDetails.distributionLists) {
    #     foreach ($dl in $userDetails.distributionLists) {
    #         try {
    #             $groupId = (Get-MgGroup -Filter "displayName eq '$dl'").Id
    #             if ($groupId) {
    #                 New-MgGroupMember -GroupId $groupId -DirectoryObjectId $NewUser.Id
    #                 Write-MessageLog "User $($userDetails.userPrincipalName) added to Distribution List '$groupName' with ID $groupId."
    #             }
    #             else {
    #                 Write-ErrorLog "Failed to find Distribution List with name '$groupName'."
    #             }
    #         }
    #         catch {
    #             Write-ErrorLog "Failed to add user $($userDetails.userPrincipalName) to Microsoft 365 group '$groupName': $_"
    #         }
    #     }
    # }

    # if ($userDetails.roles) {
    #     foreach ($role in $userDetails.roles) {
    #         try {
    #             New-MgRoleAssignment -RoleDefinitionId $role -PrincipalId $userDetails.Id -DirectoryScopeId "/"
    #             Write-MessageLog "Role $role assigned to user."
    #         }
    #         catch {
    #             Write-ErrorLog "Failed to assign role $($role): $_"
    #         }
    #     }
    # }
    # Write-MessageLog "Onboarding process for user $($userDetails.userPrincipalName) completed."
}

function OffboardingUser {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$userDetails
    )

    Write-MessageLog "Starting offboarding process for user $($userDetails.userPrincipalName)."

    try {
        # Reset the user's password
        Write-MessageLog "Resetting password for user $($userDetails.userPrincipalName)."
        $DinopassUrl = "https://www.dinopass.com/password/strong"
        $Password = Invoke-RestMethod -Uri $DinopassUrl

        $passwordProfile = @{
            Password                      = $Password
            ForceChangePasswordNextSignIn = $true
        }

        Update-MgUser -UserId $userDetails.Id -PasswordProfile $passwordProfile
        Write-MessageLog "Password reset for user $($userDetails.userPrincipalName)."

        # Block the user's sign-in
        Write-MessageLog "Blocking sign-in for user $($userDetails.userPrincipalName)."
        Update-MgUser -UserId $userDetails.Id -AccountEnabled $false
        Write-MessageLog "Sign-in blocked for user $($UserDetails.userPrincipalName)."

        # Assign mailbox delegation to the manager
        Write-MessageLog "Assigning mailbox delegation to $($UserDetails.ManagerUserPrincipalName) for user $($userDetails.userPrincipalName)."
        $ManagerId = (Get-MgUser -Filter "UserPrincipalName eq '$($UserDetails.ManagerUserPrincipalName)'").Id

        $DelegateParameters = @{
            principalId = $ManagerId
            role        = "Delegate"
        }

        New-MgUserMailFolderPermission -UserId $userDetails.Id -MailFolderId "inbox" -BodyParameter $DelegateParameters
        Write-MessageLog "Mailbox delegation granted to $ManagerUserPrincipalName for user $($userDetails.userPrincipalName)."

        # Hide the user from the Global Address List
        $GALParameters = @{
            ExtensionAttribute10 = "HiddenFromGAL"
        }
        Write-MessageLog "Hiding user $($userDetails.userPrincipalName) from the GAL."
        Update-MgUser -UserId $userDetails.Id -OnPremisesExtensionAttributes $GALParameters
        Write-MessageLog "User $($userDetails.userPrincipalName) hidden from GAL."

    }
    catch {
        Write-ErrorLog "Failed during initial offboarding steps for user $($userDetails.userPrincipalName): $_"
    }

    # Remove the user from all groups
    try {
        Write-MessageLog "Fetching all groups for user $($userDetails.userPrincipalName)."
        $Groups = Get-MgUserMemberOf -UserId $userDetails.Id | Where-Object { $_.ODataType -eq "#microsoft.graph.group" }
        foreach ($group in $Groups) {
            Remove-MgGroupMember -GroupId $group.Id -DirectoryObjectId $userDetails.Id
            Write-MessageLog "User $($userDetails.userPrincipalName) removed from group $($group.DisplayName)."
        }
    }
    catch {
        Write-ErrorLog "Failed to fetch or remove groups for user $($userDetails.userPrincipalName): $_"
    }

    # Remove all assigned roles for the user
    try {
        Write-MessageLog "Fetching all roles assigned to user $($userDetails.userPrincipalName)."
        $RoleAssignments = Get-MgUserAppRoleAssignment -UserId $userDetails.Id
        foreach ($role in $RoleAssignments) {
            Remove-MgUserAppRoleAssignment -UserId $userDetails.Id -AppRoleAssignmentId $role.Id
            Write-MessageLog "Role assignment removed: $($role.PrincipalDisplayName) for user $($userDetails.userPrincipalName)."
        }
    }
    catch {
        Write-ErrorLog "Failed to fetch or remove roles for user $($userDetails.userPrincipalName): $_"
    }

    # Remove the user from all distribution lists
    try {
        Write-MessageLog "Fetching all distribution lists for user $($userDetails.userPrincipalName)."
        $DistributionLists = Get-MgUserMemberOf -UserId $userDetails.Id | Where-Object { $_.ODataType -eq "#microsoft.graph.group" }
        foreach ($dl in $DistributionLists) {
            Remove-MgGroupMember -GroupId $dl.Id -DirectoryObjectId $userDetails.Id
            Write-MessageLog "User $($userDetails.userPrincipalName) removed from distribution list $($dl.DisplayName)."
        }
    }
    catch {
        Write-ErrorLog "Failed to fetch or remove distribution lists for user $($userDetails.userPrincipalName): $_"
    }

    Write-MessageLog "Offboarding process for user $($userDetails.userPrincipalName) completed."
}
