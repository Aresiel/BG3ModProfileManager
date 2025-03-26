$ErrorActionPreference = "Stop"
$linkType = "Junction"

$appdataFolder = Join-Path $env:LocalAppdata "\Larian Studios\Baldur's Gate 3\" -Resolve
$profileManagerFolder = Join-Path $appdataFolder "\AresielProfileManager"

function Is-Profile-Linked {
  return Test-Path (Join-Path $profileManagerFolder "\current")
}

function Unlink-Profile {
    if((Get-Item -Path (Join-Path $appdataFolder "\Mods") -Force).LinkType -eq "Junction")
    {
       (Get-Item (Join-Path $appdataFolder "\Mods")).Delete()
    }
    if((Get-Item -Path (Join-Path $appdataFolder "\PlayerProfiles") -Force).LinkType -eq "Junction")
    {
        (Get-Item (Join-Path $appdataFolder "\PlayerProfiles")).Delete()
    }
    if((Get-Item -Path (Join-Path $appdataFolder "\Script Extender") -Force).LinkType -eq "Junction")
    {
        (Get-Item (Join-Path $appdataFolder "\Script Extender")).Delete()
    }
    Remove-Item (Join-Path $profileManagerFolder "\current")
}

function Link-Profile {
    param (
        [string] $ProfileName
    )

    $profilePath = Join-Path $profileManagerFolder $ProfileName -Resolve

    if(Test-Path (Join-Path $profileManagerFolder "\current"))
    {
        throw "Cannot link new profile. A profile is already linked."
    }
    
    New-Item -ItemType $linkType -Path (Join-Path $appdataFolder "\Mods") -Value (Join-Path $profilePath "\Mods") | Out-Null
    New-Item -ItemType $linkType -Path (Join-Path $appdataFolder "\PlayerProfiles") -Value (Join-Path $profilePath "\PlayerProfiles") | Out-Null
    New-Item -ItemType $linkType -Path (Join-Path $appdataFolder "\Script Extender") -Value (Join-Path $profilePath "\Script Extender") | Out-Null

    New-Item -ItemType "File" -Path (Join-Path $profileManagerFolder "\current") -Value $ProfileName | Out-Null
}

# Initialisation
if (!(Test-Path (Join-Path $profileManagerFolder "\initalized")))
{
    Write-Output "Profile Manager not initalized"
    $answer = Read-Host "Do you want to initalize the profile manager? Make backups before proceeding. (Y/N)"
    if($answer -ne "Y")
    {
        Write-Output "Exiting profile manager."
        Exit
    }
    Write-Output "Initializing profile manager."

    New-Item -ItemType "Directory" -Force -Path $profileManagerFolder | Out-Null
    New-Item -ItemType "File" -Path (Join-Path $profileManagerFolder "\initalized") | Out-Null

    if(Test-Path (Join-Path $profileManagerFolder "\Default"))
    {
        Write-Host "Default profile already exists. Profile manager is in an invalid state since the default profile exists yet the profile manager isn't initialized."
        Write-Host "If you are okay with losing everything inside the Default profile, you may delete it and run the profile manager again."
        Write-Host "Aborting."
        Exit
    }

    New-Item -ItemType "Directory" -Path (Join-Path $profileManagerFolder "\Default") | Out-Null
    $defaultProfilePath = Join-Path $profileManagerFolder "\Default" -Resolve
    Write-Output "Created Default profile folder."


    if(Test-Path (Join-Path $appdataFolder "\Mods"))
    {
        Move-Item (Join-Path $appdataFolder "\Mods") -Destination $defaultProfilePath
        Write-Output "Imported existing Mods folder."
    } else {
        New-Item -ItemType "Directory" -Path (Join-Path $appdataFolder "\Mods")
        Write-Output "Existing Mods folder not found, creating."
    }
    
    if(Test-Path (Join-Path $appdataFolder "\PlayerProfiles"))
    {
        Move-Item (Join-Path $appdataFolder "\PlayerProfiles") -Destination $defaultProfilePath
        Write-Output "Imported existing PlayerProfiles folder."
    } else {
        New-Item -ItemType "Directory" -Path (Join-Path $appdataFolder "\PlayerProfiles")
        Write-Output "Existing PlayerProfiles folder not found, creating."
    }


    if(Test-Path (Join-Path $appdataFolder "\Script Extender"))
    {
        Move-Item (Join-Path $appdataFolder "\Script Extender") -Destination $defaultProfilePath
        Write-Output "Imported existing Script Extender folder."
    } else {
        New-Item -ItemType "Directory" -Path (Join-Path $defaultProfilePath "\Script Extender")
        Write-Output "Existing Script Extender folder not found, creating."
    }
    
    Link-Profile "Default"

    Write-Output "Initialized profile manager. Please restart the profile manager to use it."
    Exit
}

# Profile Manager
function Write-Help {
    Write-Output "Command Help"
    Write-Output "- help - View this message"
    Write-Output "- list - List existing profiles"
    Write-Output "- link <profile> - Link (Activate) a profile."
    Write-Output "- unlink - Unlink (Deactive) the existing profile. This will result in no profile being active, use with caution."
    Write-Output "- copy <profile> <new profile> - Create a copy of a profile"
    Write-Output "Commands are case sensitive."
    Write-Output ""
    Write-Output "General Help"
    Write-Output "All profile names need to be valid directory names without spaces."
}

Write-Output "Welcome to Aresiel's Profile Manager"
Write-Help
While($true)
{
    Write-Host ""
    $input = Read-Host "Enter command"
    $command, $args = -split $input

    if($command -eq "help")
    {
        Write-Help
    } elseif($command -eq "list") 
    {
        if($args.Count -ne 0) {
          Write-Output "Invalid arguments."
          continue
        }

        $profiles = Get-ChildItem -Directory $profileManagerFolder
        $current = if(Test-Path (Join-Path $profileManagerFolder "\current")) {
            Get-Content (Join-Path $profileManagerFolder "\current")
        } else { "" }
        foreach($profile in $profiles)
        {
          $active_str = if ($profile.Name -eq $current) {" (active)"} else {""}
          Write-Output "- $($profile.Name)$($active_str)"
        }

    } elseif($command -eq "link")
    {
        if($args.Count -ne 1) {
          Write-Output "Invalid arguments."
          continue
        }
        
        $profile = $args
        $confirm = Read-Host "Link profile $($profile)? (Y/N)"
        if($confirm -eq "Y")
        {
            if(Is-Profile-Linked)
            {
                $existing = Get-Content (Join-Path $profileManagerFolder "\current")
                Write-Host "Unlinking existing profile $($existing)..."
                Unlink-Profile
                Write-Host "Unlinked profile $($existing)!"
            }
            Write-Host "Linking profile $($profile)..."
            Link-Profile $profile
            Write-Host "Linked profile $($profile)!"
        }
    } elseif($command -eq "unlink")
    {
        if($args.Count -ne 0) {
          Write-Output "Invalid arguments."
          continue
        }

        if(Is-Profile-Linked) {
            $profile = Get-Content (Join-Path $profileManagerFolder "\current")
            $confirm = Read-Host "Unlink profile $($profile)? (Y/N)"
            if($confirm -eq "Y")
            {
                Write-Host "Unlinking profile $($profile)..."
                Unlink-Profile
                Write-Help "Unlinked profile $($profile)!"
            }
        } else {
          Write-Output "No profile is currently linked."
        }
    } elseif($command -eq "copy") 
    {
        if($args.Count -ne 2) {
          Write-Output "Invalid arguments."
          continue
        }

        $src = $args[0]
        $dest = $args[1]

        if(($dest -eq "initialized") -or ($dest -eq "current"))
        {
            Write-Output "Profiles may not be named `"initialized`" or `"current`"."
            continue
        }

        $confirm = Read-Host "Copy profile $($src) to $($dest)? (Y/N)"
        if($confirm -eq "Y") 
        {
            if(!(Test-Path (Join-Path $profileManagerFolder $src)))
            {
                Write-Output "Profile $($src) does not exist."
                continue
            }

            if(Test-Path (Join-Path $profileManagerFolder $dest))
            {
                Write-Output "Profile $($dest) already exists!"
                continue
            }

            Write-Output "Copying profile $($src) to $($dest)..."
            Copy-Item (Join-Path $profileManagerFolder $src) -Destination (Join-Path $profileManagerFolder $dest) -Recurse
            Write-Output "Copied profile $($src) to $($dest)!"
        }
    } else
    {
        Write-Output "Invalid command."
    }
}