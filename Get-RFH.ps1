﻿Function Get-RFH {
    [Cmdletbinding()]
    Param (

        [Parameter(ValueFromPipeline=$true)]
        [string[]]$ComputerName,

        [Parameter(Mandatory)]
        [ValidateSet("D","O","W","M","P","V","F","A","S","C","L","H","G")]
        [string[]]$Library,

        [string[]]$ExcludeAccount,

        [string]$LogAll,
        [string]$LogError,

        [switch]$ShowHost = $false
    )
    
    BEGIN {
        Write-Verbose "         Script Started        "
        Write-Verbose "   >> BEGIN BLOCK STARTED <<   "
        
        # getting current date/time and the number of libraries being checked
        $DateStart = Get-Date
        $LCount = $Library.Count

        # collecting the full names of libraries being checked for logging and output purposes based on the input provided to the -Library parameter
        $LFullCollection = @()
        $i = 0
        Do {
            switch ($Library[$i]) {
                "D" {$LFull = "Desktop"}
                "O" {$LFull = "Documents"}
                "W" {$LFull = "Downloads"}
                "M" {$LFull = "Music"}
                "P" {$LFull = "Pictures"}
                "V" {$LFull = "Video"}
                "F" {$LFull = "Favorites"}
                "A" {$LFull = "AppData"}
                "S" {$LFull = "Start Menu"}
                "C" {$LFull = "Contacts"}
                "L" {$LFull = "Links"}
                "H" {$LFull = "Searches"}
                "G" {$LFull = "Saved Games"}
            }
            $LFullCollection += "$LFull"
            $i += 1
            }
            While ($i -lt $LCount)

        # used as a property/column selector for the CSV output generated by the -LogError parameter
        $PropertyOutput = @("ComputerName","User")
        $PropertyOutput += $LFullCollection

        # writing script beginning to the application event log
        eventcreate /ID 13 /L APPLICATION /T INFORMATION /SO RedirectedFolderHealth /D "RFH script started on $DateStart for Library(ies) $LFullCollection" > $null

        # prep work by gathering enabled AD user objects with their name and SID as well as creating the empty collection arrays
        $ADUser = Get-ADUser -Filter 'enabled -eq $true'
                Write-Verbose "Stored enabled Active Directory user objects"
        $UserCollection = @()
                Write-Verbose "Created empty array for custom user object information"
        $ResultCollection = @()
                Write-Verbose "Created empty array for custom result object information"

        ForEach ($a in $ADUser) {
            $ObjUser = New-Object -TypeName psobject
            $ObjUser | Add-Member -MemberType NoteProperty -Name "Name" -Value $a.SamAccountName
            $ObjUser | Add-Member -MemberType NoteProperty -Name "SID" -Value ($a.SID).Value

            If ($ExcludeAccount -notcontains $a.SamAccountName) {
                $UserCollection += $ObjUser
            }
        }
            Write-Verbose "Custom user object information gathered and completed"

        # gets the full name of the libraries selected to show script progress and information on the console when specifying the -ShowHost switch parameter
        If ($ShowHost) {
            Write-Host -ForegroundColor Yellow "RFH script started on $DateStart." `n "Library(ies) being checked:"
            $j = 0
            Do {
                switch ($Library[$j]) {
                    "D" {$LFull = "Desktop"}
                    "O" {$LFull = "Documents"}
                    "W" {$LFull = "Downloads"}
                    "M" {$LFull = "Music"}
                    "P" {$LFull = "Pictures"}
                    "V" {$LFull = "Videos"}
                    "F" {$LFull = "Favorites"}
                    "A" {$LFull = "AppData"}
                    "S" {$LFull = "Start Menu"}
                    "C" {$LFull = "Contacts"}
                    "L" {$LFull = "Links"}
                    "H" {$LFull = "Searches"}
                    "G" {$LFull = "Saved Games"}
                }
                Write-Host -ForegroundColor Yellow "   "$LFull
                $j += 1
            }
            While ($j -lt $LCount)
        }

        Write-Verbose "   >> BEGIN BLOCK FINISHED <<   "
    }
    PROCESS {
        ForEach ($c in $ComputerName) {
            Write-Verbose "   >> PROCESS BLOCK STARTED FOR $c <<   "

            Try {
                # creates the session to the computer
                $Session = New-PSSession -ComputerName $c -Name "PSS_$c" -ErrorAction Stop
                If ($ShowHost) {Write-Host -ForegroundColor Yellow "Checking for redirections loaded on $c..."}
                    Write-Verbose "$c -   [SESSION ESTABLISHED]"

                # gets the user accounts to check on the computer
                $LocalUserFolder = Invoke-Command -Session $Session {Get-ChildItem C:\users}
                $LocalUserName = $LocalUserFolder.Name
                    Write-Verbose "$c -   Stored users"
                
                # stores the SID of each user that was found on the machine
                $LocalUserFull = $UserCollection | Where-Object {$LocalUserName -match $_.Name} | Select-Object Name,SID
                    Write-Verbose "$c -   Obtained SID information for stored users"

                # operates the data gathering for each user that was found on the computer
                ForEach ($l in $LocalUserFull) {
                    $CurrentUserSID = $l.SID
                    $CurrentUserName = $l.Name
                        Write-Verbose "$c - Checking $CurrentUserName"
                    
                    # store the "D" desktop path value for the user
                    If ($Library -eq "D") {
                        $DesktopPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name Desktop -ErrorAction SilentlyContinue).Desktop
                        }
                        If ($DesktopPath) {Write-Verbose "$c -   Desktop value stored as $DesktopPath"}

                        # conditions for logging and reporting
                        If ($DesktopPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Desktop path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($DesktopPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Desktop path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($DesktopPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Desktop path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Desktop path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "O" documents path value for the user
                    If ($Library -eq "O") {
                        $DocumentsPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name Personal -ErrorAction SilentlyContinue).Personal
                        }
                        If ($DocumentsPath) {Write-Verbose "$c -   Documents value stored as $DocumentsPath"}

                        # conditions for logging and reporting
                        If ($DocumentsPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Documents path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($DocumentsPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Documents path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($DocumentsPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Documents path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Documents path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "W" downloads path value for the user
                    If ($Library -eq "W") {
                        $DownloadsPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{374DE290-123F-4565-9164-39C4925E467B}" -ErrorAction SilentlyContinue)."{374DE290-123F-4565-9164-39C4925E467B}"
                        }
                        If ($DownloadsPath) {Write-Verbose "$c -   Downloads value stored as $DownloadsPath"}

                        # conditions for logging and reporting
                        If ($DownloadsPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Downloads path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($DownloadsPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Downloads path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($DownloadsPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Downloads path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Downloads path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "M" music path value for the user
                    If ($Library -eq "M") {
                        $MusicPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "My Music" -ErrorAction SilentlyContinue)."My Music"
                        }
                        If ($MusicPath) {Write-Verbose "$c -   Music value stored as $MusicPath"}

                        # conditions for logging and reporting
                        If ($MusicPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Music path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($MusicPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Music path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($MusicPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Music path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Music path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "P" pictures path value for the user
                    If ($Library -eq "P") {
                        $PicturesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "My Pictures" -ErrorAction SilentlyContinue)."My Pictures"
                        }
                        If ($PicturesPath) {Write-Verbose "$c -   Pictures value stored as $PicturesPath"}

                        # conditions for logging and reporting
                        If ($PicturesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Pictures path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($PicturesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Pictures path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($PicturesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Pictures path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Pictures path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "V" video path value for the user
                    If ($Library -eq "V") {
                        $VideoPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "My Video" -ErrorAction SilentlyContinue)."My Video"
                        }
                        If ($VideoPath) {Write-Verbose "$c -   Video value stored as $VideoPath"}

                        # conditions for logging and reporting
                        If ($VideoPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Video path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($VideoPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Video path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($VideoPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Video path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Video path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "F" favorites path value for the user
                    If ($Library -eq "F") {
                        $FavoritesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name Favorites -ErrorAction SilentlyContinue).Favorites
                        }
                        If ($FavoritesPath) {Write-Verbose "$c -   Favorites value stored as $FavoritesPath"}

                        # conditions for logging and reporting
                        If ($FavoritesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Favorites path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($FavoritesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Favorites path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($FavoritesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Favorites path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Favorites path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "A" roaming appdata path value for the user
                    If ($Library -eq "A") {
                        $AppDataPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name AppData -ErrorAction SilentlyContinue).AppData
                        }
                        If ($AppDataPath) {Write-Verbose "$c -   AppData value stored as $AppDataPath"}

                        # conditions for logging and reporting
                        If ($AppDataPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   AppData path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($AppDataPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   AppData path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($AppDataPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   AppData path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "AppData path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "S" start menu path value for the user
                    If ($Library -eq "S") {
                        $StartMenuPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "Start Menu" -ErrorAction SilentlyContinue)."Start Menu"
                        }
                        If ($StartMenuPath) {Write-Verbose "$c -   StartMenu value stored as $StartMenuPath"}

                        # conditions for logging and reporting
                        If ($StartMenuPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   StartMenu path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($StartMenuPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   StartMenu path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($StartMenuPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   StartMenu path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "StartMenu path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "C" contacts path value for the user
                    If ($Library -eq "C") {
                        $ContactsPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{56784854-C6CB-462B-8169-88E350ACB882}" -ErrorAction SilentlyContinue)."{56784854-C6CB-462B-8169-88E350ACB882}"
                        }
                        If ($ContactsPath) {Write-Verbose "$c -   Contacts value stored as $ContactsPath"}

                        # conditions for logging and reporting
                        If ($ContactsPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Contacts path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($ContactsPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Contacts path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($ContactsPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Contacts path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Contacts path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "L" links path value for the user
                    If ($Library -eq "L") {
                        $LinksPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}" -ErrorAction SilentlyContinue)."{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}"
                        }
                        If ($LinksPath) {Write-Verbose "$c -   Links value stored as $LinksPath"}

                        # conditions for logging and reporting
                        If ($LinksPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Links path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($LinksPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Links path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($LinksPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Links path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Links path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "H" searches path value for the user
                    If ($Library -eq "H") {
                        $SearchesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}" -ErrorAction SilentlyContinue)."{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}"
                        }
                        If ($SearchesPath) {Write-Verbose "$c -   Searches value stored as $SearchesPath"}

                        # conditions for logging and reporting
                        If ($SearchesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Searches path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($SearchesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   Searches path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($SearchesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   Searches path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "Searches path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # store the "G" saved games path value for the user
                    If ($Library -eq "G") {
                        $SavedGamesPath = Invoke-Command -Session $Session {
                            (Get-ItemProperty -Path "Registry::HKEY_USERS\$($using:CurrentUserSID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -Name "{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}" -ErrorAction SilentlyContinue)."{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}"
                        }
                        If ($SavedGamesPath) {Write-Verbose "$c -   SavedGames value stored as $SavedGamesPath"}

                        # conditions for logging and reporting
                        If ($SavedGamesPath -like "\\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   SavedGames path for user $CurrentUserName on machine $c is redirected."}
                        }
                        ElseIf ($SavedGamesPath -like "*\OneDrive\*") {
                            If ($ShowHost) {Write-Host -ForegroundColor Green "   SavedGames path for user $CurrentUserName on machine $c is in OneDrive."}
                        }
                        ElseIf ($SavedGamesPath) {
                            If ($ShowHost) {Write-Host -ForegroundColor Red "   SavedGames path for user $CurrentUserName on machine $c is not redirected!"}
                            eventcreate /ID 13 /L APPLICATION /T WARNING /SO RedirectedFolderHealth /D "SavedGames path for user $CurrentUserName on machine $c is not redirected!" > $null
                        }
                    }

                    # creates and writes members to the object for the result of user on the computer based on selected libraries
                    $ObjResult = New-Object -TypeName psobject
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$c"
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "User" -Value "$CurrentUserName"

                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Desktop" -Value $DesktopPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Documents" -Value $DocumentsPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Downloads" -Value $DownloadsPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Music" -Value $MusicPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Pictures" -Value $PicturesPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Video" -Value $VideoPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Favorites" -Value $FavoritesPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "AppData" -Value $AppDataPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "StartMenu" -Value $StartMenuPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Contacts" -Value $ContactsPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Links" -Value $LinksPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "Searches" -Value $SearchesPath
                    $ObjResult | Add-Member -MemberType NoteProperty -Name "SavedGames" -Value $SavedGamesPath

                    # appends the resulting object of the computer's redirection check to the collection of results only if any of the paths contain values
                    # because an object is made for every user account found on the target machine regardless of being logged in, this prevents objects with blank path values from being returned
                    If($ObjResult.Desktop -or 
                        $ObjResult.Documents -or
                        $ObjResult.Downloads -or
                        $ObjResult.Music -or
                        $ObjResult.Pictures -or
                        $ObjResult.Video -or
                        $ObjResult.Favorites -or
                        $ObjResult.AppData -or
                        $ObjResult.StartMenu -or
                        $ObjResult.Contacts -or
                        $ObjResult.Links -or
                        $ObjResult.Searches -or
                        $ObjResult.SavedGames) 
                    {
                        $ResultCollection += $ObjResult
                        Write-Verbose "$c -   Object added to the collection array"
                    }
                }

                # removes the session to the computer
                Remove-PSSession -Session $Session
                    Write-Verbose "$c -   [SESSION REMOVED]"
            }
            Catch {
                Write-Output "Warning: The computer $c could not be contacted!"
            }

            Write-Verbose "   >> PROCESS BLOCK FINISHED FOR $c <<   "
        }
    }
    END {
        Write-Verbose "   >> END BLOCK STARTED <<   "

        # outputs the collection of results as specified, containing all computers and all users found on each computer
        If ($LogAll) {
            # writes all findings to csv
            Write-Output $ResultCollection | Export-Csv -Path $LogAll -NoTypeInformation
        }
        If ($LogError) {
            # writes only problems to csv
            ForEach ($r in $ResultCollection) {
                If (
                    $r.Desktop -and $r.Desktop -notlike "\\*" -and $r.Desktop -notlike "*OneDrive*" -or
                    $r.Documents -and $r.Documents -notlike "\\*" -and $r.Documents -notlike "*OneDrive*" -or
                    $r.Downloads -and $r.Downloads -notlike "\\*" -and $r.Downloads -notlike "*OneDrive*" -or
                    $r.Music -and $r.Music -notlike "\\*" -and $r.Music -notlike "*OneDrive*" -or
                    $r.Pictures -and $r.Pictures -notlike "\\*" -and $r.Pictures -notlike "*OneDrive*" -or
                    $r.Video -and $r.Video -notlike "\\*" -and $r.Video -notlike "*OneDrive*" -or
                    $r.Favorites -and $r.Favorites -notlike "\\*" -and $r.Favorites -notlike "*OneDrive*" -or
                    $r.AppData -and $r.AppData -notlike "\\*" -and $r.AppData -notlike "*OneDrive*" -or
                    $r.StartMenu -and $r.StartMenu -notlike "\\*" -and $r.StartMenu -notlike "*OneDrive*" -or
                    $r.Contacts -and $r.Contacts -notlike "\\*" -and $r.Contacts -notlike "*OneDrive*" -or
                    $r.Links -and $r.Links -notlike "\\*" -and $r.Links -notlike "*OneDrive*" -or
                    $r.Searches -and $r.Searches -notlike "\\*" -and $r.Searches -notlike "*OneDrive*" -or
                    $r.SavedGames -and $r.SavedGames -notlike "\\*" -and $r.SavedGames -notlike "*OneDrive*"
                )
                {
                    Write-Output $r | Select-Object $PropertyOutput | Export-Csv -Path $LogError -NoTypeInformation -Append
                }
            }
        }
        If (!$LogAll -and !$LogError) {
            # writes all findings and full object info to the pipeline
            Write-Output $ResultCollection
        }

        # obtaining info to report on elapsed time taken for the script to complete
        $DateEnd = Get-Date
        $DateDiff = $DateEnd - $DateStart
        $Hour = $DateDiff.Hours
        $Minute = $DateDiff.Minutes
        $Second = $DateDiff.Seconds

        # writing completion to the application event log and out to host
        eventcreate /ID 13 /L APPLICATION /T INFORMATION /SO RedirectedFolderHealth /D "RFH script completed on $DateEnd after $Hour hour(s), $Minute minute(s), and $Second second(s) for Library(ies) $LFullCollection." > $null
        If ($ShowHost) {Write-Host -ForegroundColor Yellow "RFH script completed on $DateEnd after $Hour hour(s), $Minute minute(s), and $Second second(s) for Library(ies) $LFullCollection"}

        Write-Verbose "   >> END BLOCK FINISHED <<   "
        Write-Verbose "       Script Completed       "
    }
}

<#
    - add in account exclusion options (needs to take multiple values) (continue testing this)
    - address issue with not being able to supply multiple values to the ComputerName parameter
        - currently the function works fine while feeding multiple machines in from the pipeline, but the ability to use other methods to supply that input for multiple values does not work
        - get the tool to work for something like
            Get-RFH -ComputerName sl-computer-001,sl-computer-002 -Library D
            Get-RFH -ComputerName (Get-Content C:\Test\Computers.txt) -Library D
    - write in email sending functionality (change param names to reflect the name of the respective param in the Send-MailMessage cmdlet)
        - param [switch]SendEmail
        - param [string]ToAddress
        - param [string]FromAddress
        - param [string]SmtpServer
        - param [int]Port
        - param [switch]UseSSL
    - parameter set will need to be made so that the logging options are required if the email option is selected
#>

