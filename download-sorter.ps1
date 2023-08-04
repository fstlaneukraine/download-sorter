# Organizes torrents based on category to a series of folders
# Written by FstLaneUkraine
# January 2018
# Last Updated March 2023
# qBittorrent run external program on torrent completion command: powershell -ExecutionPolicy ByPass -File ".\download-sorter.ps1" "%F" "%L" "%N" "%R"

<#

#Arguments being passed from qBittorrent

%~1 = %F = Content path (same as root path for multiple torrent)
%~2 = %L = Category
%~3 = %N = Torrent Name
%~4 = %R = Root path (first torrent subdirectory path)

#>

Param(
  [string]$filePath,
  [string]$Category,
  [string]$TorrName,
  [string]$fileDirectory
)

# Setting variables

$WinRarPath = "C:\Software\WinRAR\RAR.exe" # Path to RAR.exe
$TempLoc = "H:\Temp" # Temporary location to be used for unRARing, etc.
$Ext1 = "*.avi"
$Ext2 = "*.mkv"
$Ext3 = "*.mp4"
$TVDBBaseURI = "https://api4.thetvdb.com/v4"
$TVDBAPIKey = "somekey"
$TVDBPin = "somepin"
$MotoGPTVDBSeriesID = "81083"
$F1TVDBSeriesID = "387219"
$MotoGPSeason = Get-Date -Format yyyy
$Formula1Season = Get-Date -Format yyyy
$DestPath = "H:\$Category" # Final destination for the files
$TorrPartName = $TorrName.SubString(0,3) # Getting the first 3 characters of the torrent name
$Location = "H:\Completed" # Source location of the downloaded torrent
$FileDirs = @(Get-ChildItem $Location -Recurse -Exclude Proof*, Subs*, Sample* | Where-Object { $_.PSIsContainer } | Select-Object Name) # Getting the main files of a torrent and skipping the sample, subs, proof, etc.
$TorrFiles = Get-ChildItem $Location -Recurse -Include  *.torrent | Where-Object {! $_.PSIsContainer } # Getting the list of torrent files in location
$F1Emails = @('someemail@gmail.com')
$MotoGPEmails = @('someemail@gmail.com')
$TopGearEmails = @('someemail@gmail.com')
$GrandTourEmails = @('someemail@gmail.com')
$SMTPServer = "smtp.gmail.com"
$SMTPUsername = "someemail@gmail.com";
$SMTPPassword = "somepassword";
 
# Creating reusable functions
 
function SendEmail ($EmailFrom, $EmailTo, $SMTPServer, $SMTPUsername, $SMTPPassword)
{
    $Subject = "[FstLaneUkraine's Server] - $NewFileName Available"
    $Body = "<p>Hey!</p>
 
    <p>This is just an automated email notifying you that <b><i>$NewFileName</b></i> has been downloaded and is available on my <i>PleX</i> and <i>NAS</i> Server! If you have any issues finding it, just let me know!</p>
 
    <p>Thanks, FstLaneUkraine</p>"
    $Message = New-Object System.Net.Mail.Mailmessage $EmailFrom, $EmailTo, $Subject, $Body
    $Message.IsBodyHTML=$True
    $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, 587)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPUsername", "$SMTPPassword");
    $SMTPClient.Send($Message)
 
    Write-Host ""
    Write-Host "      Notification Email Sent!" -ForegroundColor Green
}
 
function CheckFolder ($Folder)
{
    Write-Host ""
    Write-Host "Checking if " -ForegroundColor Yellow -NoNewLine;
    Write-Host "$Folder " -ForegroundColor Cyan -NoNewLine;
    Write-Host "exists..." -ForegroundColor Yellow
 
    if (!(Test-Path $Folder))
    {
        New-Item $Folder -Type Directory
        if (Test-Path $Folder)
        {
            Write-Host ""
            Write-Host "    $Folder" -ForegroundColor Cyan -NoNewLine;
            Write-Host " directory created..." -ForegroundColor Green
            Write-Host ""
        }
        elseif (!(Test-Path $Folder))
        {
            Write-Host ""
            Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
            Write-Host "$Folder!" -ForegroundColor Cyan
            Write-Host ""
        }
    }
    else
    {
        Write-Host "    $Folder" -ForegroundColor Cyan -NoNewLine;
        Write-Host " already exists..." -ForegroundColor Green
    }
}
 
function UnRarSortToCategory ($Category)
{
    # Updating the fileDirectory if qBittorent didn't pass it along because of how the torrent was packaged
    
    if (!($fileDirectory))
    {
        $fileDirectory = Split-Path -Path $filePath
    }
    
    Write-Host ""
    Write-Host "$Category" -ForegroundColor Magenta -NoNewLine;
    Write-Host " torrent detected..." -ForegroundColor Yellow
    Write-Host ""
    
    if ($Category -eq "Games")
    {
        # UnRARing files to main category repository
 
        Write-Host "UnRARing files to proper category repository..." -ForegroundColor Yellow
        Write-Host ""
 
        &$WinRarPath x -ibck -inul $fileDirectory/*.r* $DestPath
 
        Get-Process WinRar | Wait-Process
    }
    else
    {
        # Copying all files from download to temporary location
 
        Write-Host "Copying files from " -ForegroundColor Yellow -NoNewLine;
        Write-Host "$fileDirectory" -ForegroundColor Cyan -NoNewLine;
        Write-Host " to " -ForegroundColor Yellow -NoNewLine;
        Write-Host "$TempLoc..." -ForegroundColor Cyan
 
        # If Torrent Name is a MotoGP or Formula 1 event, extending the TorrName substring to 40 characters to avoid working with both qualifying and grand prix events at the same time

        if ($TorrName -match 'Sprint' -or $TorrName -match 'Qualifying' -or $TorrName -match 'Race' -or $TorrName -match 'Prix' -or $TorrName -match 'Grand.Prix' -or $TorrName -match 'Grand Prix' -or $TorrName -match 'Shootout')
        {
            # Renaming MotoGP and Formula 1 files to the torrent name for better renaming later

            $MotoGPFiles = Get-ChildItem -Path $fileDirectory -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo,*Pre-*,*Post-*,*Analysis*,*Sample* -Recurse | Where-Object { $_.Name -match "MotoGP" -OR $_.Name -match "Qualifying.MotoGP" -OR $_.Name -match "Race.MotoGP" -OR $_.Name -match "04.MotoGP.Sprint.Race" }
            $Formula1Files = Get-ChildItem -Path $fileDirectory -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo,*Pre-*,*Post-*,*Analysis*,*Sample* -Recurse | Where-Object { $_.Name -match "Formula.1" -OR $_.Name -match "02.Qualifying.Session" -OR $_.Name -match "02.Race.Session" -OR $_.Name -match "prix" -OR $_.Name -match "02.Sprint.Session" -OR $_.Name -match "Shootout" -OR $_.Name -match "testing" }

            if ($MotoGPFiles)
            {
                foreach ($MotoGPFile in $MotoGPFiles)
                {
                    $FileExtension = $MotoGPFile.Extension
                    $FilePath = Split-Path -Path $MotoGPFile
                    $NewFileName = $TorrName+$FileExtension
                    Robocopy $FilePath $TempLoc $($MotoGPFile.Name) /z /j /NOOFFLOAD /MT
                    Rename-Item -Path $TempLoc\$($MotoGPFile.Name) $($NewFileName)
                } 
            }

            if ($Formula1Files)
            {
                foreach ($Formula1File in $Formula1Files)
                {
                    $FileExtension = $Formula1File.Extension
                    $FilePath = Split-Path -Path $Formula1File
                    $NewFileName = $TorrName+$FileExtension
                    Robocopy $FilePath $TempLoc $($Formula1File.Name) /z /j /NOOFFLOAD /MT
                    Rename-Item -Path $TempLoc\$($Formula1File.Name) $($NewFileName)
                }
            }
        }
        else
        {
            # Ensuring right file is being copied based on $TorrName and length of the file name to avoid grabbing all episodes of a particular show

            $File = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo -Recurse | Where-Object { $_.Name -match $TorrName }

            if ($File)
            {
                $FileLength = [int]$File.Name.Length
                if ($FileLength -le "20")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 5) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
                if ($FileLength -ge "21" -AND $FileLength -le "30")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 10) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
                if ($FileLength -ge "31" -AND $FileLength -le "40")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 15) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
                if ($FileLength -ge "41" -AND $FileLength -le "50")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 25) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
                if ($FileLength -ge "51" -AND $FileLength -le "60")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 35) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
                if ($FileLength -ge "61" -AND $FileLength -le "70")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 45) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
                if ($FileLength -ge "71" -AND $FileLength -le "80")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 55) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
                if ($FileLength -ge "81")
                {
                    $FileToCopy = Get-ChildItem -Path $fileDirectory\* -Include *.r*,$Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo | Where-Object { $_.Name -match $TorrName.SubString(0, 65) }
                    $FilePath = Split-Path -Path $FileToCopy
                    Robocopy $fileDirectory $TempLoc $FileToCopy.Name /z /j /NOOFFLOAD /MT
                }
            }            
        }
 
        # Checking if files actually copied...proceeding if so and exiting if not

        if (!((Test-Path $TempLoc\* -Include $TorrPartName*.*) -OR (Test-Path $TempLoc\* -Include $($FileToCopy.Name))))
        {
            Write-Host "    No files copied!" -ForegroundColor Red
            Write-Host ""
        }
        elseif ((Test-Path $TempLoc\* -Include $TorrPartName*.*) -OR (Test-Path $TempLoc\* -Include $($FileToCopy.Name)))
        {
            Write-Host "    Files copied..." -ForegroundColor Green
            Write-Host ""
 
            $TempFiles = Get-ChildItem $TempLoc -Filter $TorrPartName*
 
            # Renaming files within the archives in the temp location to the torrent name as some groups don't name the files the same as they do the torrent
 
            if (!(Get-ChildItem -Path $TempLoc -Filter *.rar))
            {
                Write-Host "$TempLoc" -ForegroundColor White -NoNewLine;
                Write-Host " contains no RAR'ed files..." -ForegroundColor Yellow
            }
            elseif (Get-ChildItem -Path $TempLoc -Filter *.rar)
            {
                Write-Host "RAR'ed files detected in " -ForegroundColor Yellow -NoNewLine;
                Write-Host "$TempLoc" -ForegroundColor Cyan -NoNewLine;
                Write-host "..." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Renaming files within " -ForegroundColor Yellow -NoNewLine;
                Write-Host "$Torrname" -ForegroundColor White -NoNewLine;
                Write-Host " RAR archives found in " -ForegroundColor Yellow -NoNewLine;
                Write-Host "$TempLoc..." -ForegroundColor Cyan
                Write-Host ""
 
                foreach ($file in $TempFiles)
                {
                    $ToRenameOrNot = &$WinRarPath lb "$($file.FullName)"
                    if (!($ToRenameOrNot -match $TorrName))
                    {
                        Write-Host "   Currently renaming file from " -ForegroundColor White -NoNewLine;
                        Write-Host "$file" -ForegroundColor Red -NoNewLine;
                        Write-Host " to" -ForegroundColor White -NoNewLine;
                        Write-Host " $TorrName" -ForegroundColor Green -NoNewLine;
                        Write-Host "..."
                        &$WinRarPath -ibck -inul RN $TempLoc\$file "*.avi" "$TorrName.avi"
                        Start-Sleep -m 350
                        &$WinRarPath -ibck -inul RN $TempLoc\$file "*.mkv" "$TorrName.mkv"
                        Start-Sleep -m 350
                        &$WinRarPath -ibck -inul RN $TempLoc\$file "*.mp4" "$TorrName.mp4"
                        Start-Sleep -m 350
                    }
                    else
                    {
                        Write-Host "   The file " -ForegroundColor White -NoNewLine;
                        Write-Host "$file" -ForegroundColor Red -NoNewLine;
                        Write-Host " does not need to be renamed..." -ForegroundColor White
                    }
                }
 
                # UnRARing renamed files to main category repository
 
                Write-Host ""
                Write-Host "UnRARing/moving renamed version of files to proper category repository..." -ForegroundColor Yellow
 
                &$WinRarPath x -ibck -inul $TempLoc/*.r* $DestPath
 
                #Get-Process WinRar | Wait-Process # No longer needed at of 6/4/23 due to switch from using WinRar.exe to Rar.exe
            }
 
            Start-Sleep -s 5
 
            # Copying all .avi, .mkv and .mp4 files to their destination

            $FilesToCopy = Get-ChildItem -Path $TempLoc\* -Include $Ext1,$Ext2,$Ext3 -Exclude *.torrent,*.nfo,*sample*.*
            foreach ($FileToCopy in $FilesToCopy)
            {
                $FilePath = Split-Path -Path $FileToCopy
                #$NewFileName = $TorrName+$FileExtension
                Robocopy $TempLoc $DestPath $($FileToCopy.Name) /z /j /NOOFFLOAD /MT /xf *sample*.*
                #Rename-Item -Path $TempLoc\$FileToCopy.Name -NewName $NewFileName
            }

            Start-Sleep -s 5
 
            # Checking to see if the destination has files from the torrent
 
            if (!((Test-Path $DestPath\* -Include $TorrPartName*.*) -OR (Test-Path $DestPath\* -Include $($FileToCopy.Name))))
            {
                Write-Host "    $DestPath" -ForegroundColor Cyan -NoNewLine;
                Write-Host " is missing " -ForegroundColor Red -NoNewLine;
                Write-Host "$TorrName" -ForegroundColor White -NoNewLine;
                Write-Host " files!" -ForegroundColor Red
                Write-Host "    Exiting script..." -ForegroundColor Red
                EXIT
            }
            elseif ((Test-Path $DestPath\* -Include $TorrPartName*.*) -OR (Test-Path $DestPath\* -Include $($FileToCopy.Name)))
            {
                # Deleting files in the temp location
 
                Write-Host ""
                Write-Host "  $DestPath" -ForegroundColor Cyan -NoNewLine;
                Write-Host " contains " -ForegroundColor Green -NoNewLine;
                Write-Host "$TorrName" -ForegroundColor White -NoNewLine;
                Write-Host " files...proceeding with cleanup..." -ForegroundColor Green
                Write-Host ""
                Write-Host "    Deleting temporary files from " -ForegroundColor Yellow -NoNewLine;
                Write-Host "$TempLoc" -ForegroundColor Cyan -NoNewLine;
                Write-Host "..." -ForegroundColor Yellow
                Write-Host ""
 
                $CleanupFiles = Get-ChildItem $TempLoc | Where-Object { $_.Name -contains $TorrName -OR $_.Name -match $($FileToCopy.Name)}
 
                foreach ($CleanupFile in $CleanupFiles)
                {
                    Write-Host "      Removing item: " -ForegroundColor White -NoNewLine;
                    Write-Host "$CleanupFile" -ForegroundColor Magenta
                    Remove-Item -Path $TempLoc\* -Recurse -Force | Where-Object { $_.Name -contains $TorrName -OR $_.Name -match $($FileToCopy.Name)}
                }
 
                if (Test-Path $TempLoc\* -Include *.*)
                {
                    Write-Host "        Cleanup not completed, files still exist!" -ForegroundColor Red
                    Write-Host ""
                }
                elseif (!(Test-Path $TempLoc\* -Include *.*))
                {
                    Write-Host ""
                    Write-Host "        Cleanup completed, folder is empty..." -ForegroundColor Green
                    Write-Host ""
                }
            }
        }
    }
}
 
function SortShows ($TorrName)
{
    Write-Host "Now sorting shows from " -ForegroundColor Yellow -NoNewLine
    Write-Host "$DestPath" -ForegroundColor Cyan -NoNewLine
    Write-Host "..." -ForegroundColor Yellow
 
    if (!($TorrName -match 'MotoGP' -or $TorrName -match 'Formula.1' -or $TorrName -match 'Formula1' -or $TorrName -match 'Formula.One' -or $TorrName -match 'Formula 1'))
    {
        $TorrName -match '([S]\d\d)' | Out-Null
        $Season = $matches[1].SubString(1,2) # Getting the first 3 characters of the torrent name and determining what season the torrent is part of
    }
    
    $Shows = Get-ChildItem -Path "H:\TV Shows" -Recurse | Select-Object Name,DirectoryName,LastWriteTime
 
    foreach ($Show in $Shows)
    {
        # Sorting The Grand Tour Episodes
 
        if ($($Show.name) -like "the.grand.tour.s$Season*")
        {
            if ($Season -lt 10)
            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "The Grand Tour" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            $TGTPath = "I:\Movies and Shows\Car Shows\The Grand Tour\Season $($Season)"
            if (!(Test-Path $TGTPath))
            {
                New-Item $TGTPath -Type Directory | Out-Null
                if (Test-Path $TGTPath)
                {
                    Write-Host "    $TGTPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $TGTPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$TGTPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            $NewFileName = Get-ChildItem -Path "H:\$Category" | Where-Object {$_.Name -match $TorrPartName}
            $NewFileName = $NewFileName.basename
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $TGTPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov /is /it /im
 
            foreach ($GrandTourEmail in $GrandTourEmails)
            {
                $EmailFrom = "$SMTPUsername"
                $EmailTo = "$GrandTourEmail"
 
                SendEmail $EmailFrom $EmailTo $SMTPServer $SMTPUsername $SMTPPassword
            }
        }
 
        # Sorting Top Gear US Episodes
 
        if ($($Show.name) -like "top.gear.us.s$Season*")
        {
            if ($Season -lt 10)
            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Top Gear US" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            
            $TGUSPath = "I:\Movies and Shows\Car Shows\Top Gear US\Season $($Season)"
            if (!(Test-Path $TGUSPath))
            {
                New-Item $TGUSPath -Type Directory | Out-Null
                if (Test-Path $TGUSPath)
                {
                    Write-Host "    $TGUSPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $TGUSPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$TGUSPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
            
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            $NewFileName = Get-ChildItem -Path "H:\$Category" | Where-Object {$_.Name -match $TorrPartName}
            $NewFileName = $NewFileName.basename
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $TGUSPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov /is /it /im
 
            foreach ($TopGearEmail in $TopGearEmails)
            {
                $EmailFrom = "$SMTPUsername"
                $EmailTo = "$TopGearEmail"
 
                SendEmail $EmailFrom $EmailTo $SMTPServer $SMTPUsername $SMTPPassword
            }
        }
 
        # Sorting Top Gear UK Episodes
 
        if ($($Show.name) -like "top.gear.s$Season*")
        {
            if ($Season -lt 10)
            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Top Gear UK" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            $TGUKPath = "I:\Movies and Shows\Car Shows\Top Gear UK\Season $($Season)"
            if (!(Test-Path $TGUKPath))
            {
                New-Item $TGUKPath -Type Directory | Out-Null
                if (Test-Path $TGUKPath)
                {
                    Write-Host "    $TGUKPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $TGUKPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$TGUKPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            $NewFileName = Get-ChildItem -Path "H:\$Category" | Where-Object {$_.Name -match $TorrPartName}
            $NewFileName = $NewFileName.basename
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $TGUKPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov /is /it /im

            foreach ($TopGearEmail in $TopGearEmails)
            {
                $EmailFrom = "$SMTPUsername"
                $EmailTo = "$TopGearEmail"
 
                SendEmail $EmailFrom $EmailTo $SMTPServer $SMTPUsername $SMTPPassword
            }
        }
 
        # Sorting Top Gear America Episodes
 
        if ($($Show.name) -like "top.gear.america.s$Season*")
        {
            if ($Season -lt 10)
            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Top Gear America" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            $TGAmericaPath = "I:\Movies and Shows\Car Shows\Top Gear America\Season $($Season)"
            if (!(Test-Path $TGAmericaPath))
            {
                New-Item $TGAmericaPath -Type Directory | Out-Null
                if (Test-Path $TGAmericaPath)
                {
                    Write-Host "    $TGAmericaPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $TGAmericaPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$TGAmericaPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            $NewFileName = Get-ChildItem -Path "H:\$Category" | Where-Object {$_.Name -match $TorrPartName}
            $NewFileName = $NewFileName.basename
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $TGAmericaPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov
 
            foreach ($TopGearEmail in $TopGearEmails)
            {
                $EmailFrom = "$SMTPUsername"
                $EmailTo = "$TopGearEmail"
 
                SendEmail $EmailFrom $EmailTo $SMTPServer $SMTPUsername $SMTPPassword
            }
        }
 
        # Sorting Clarksons Farm Episodes
 
        if ($($Show.name) -like "clarksons.farm.s$Season*")
        {
            if ($Season -lt 10)
            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Clarkson's Farm" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            $CFPath = "I:\Movies and Shows\TV Shows\Clarkson's Farm\Season $($Season)"
            if (!(Test-Path $CFPath))
            {
                New-Item $CFPath -Type Directory | Out-Null
                if (Test-Path $CFPath)
                {
                    Write-Host "    $CFPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $CFPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$CFPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $CFPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov
        }
 
        # Sorting Richard Hammonds Workshop Episodes
 
        if ($($Show.name) -like "richard.hammonds.workshop.s$Season*")
        {
            if ($Season -lt 10)
            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Richard Hammondss Workshop" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            $RHWPath = "I:\Movies and Shows\Car Shows\Richard Hammonds Workshop\Season $($Season)"
            if (!(Test-Path $RHWPath))
            {
                New-Item $RHWPath -Type Directory | Out-Null
                if (Test-Path $RHWPath)
                {
                    Write-Host "    $RHWPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $RHWPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$RHWPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $RHWPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov
        }
 
        # Sorting Fast'N'Loud Episodes
 
        if ($($Show.name) -like "fast.n.loud.s$Season*")
        {
            if ($Season -lt 10)
            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Fast'N'Loud" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            $FastNLoudPath = "I:\Movies and Shows\Car Shows\Fast N Loud\Season $($Season)"
            if (!(Test-Path $FastNLoudPath))
            {
                New-Item $FastNLoudPath -Type Directory | Out-Null
                if (Test-Path $FastNLoudPath)
                {
                    Write-Host "    $FastNLoudPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $FastNLoudPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$FastNLoudPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $FastNLoudPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov
        }

        # Sorting Formula 1 Drive to Survive Episodes
 
        if ($($Show.name) -match "drive.to.survive")
        {
            $TorrName -match '([S]\d\d)' | Out-Null
            $Season = $matches[1].SubString(1,2) # Getting the first 3 characters of the torrent name and determining what season the torrent is part of
            if ($Season -lt 10)

            {
                $Season = $Season.Trim("0")
            }
            
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Formula 1 Drive to Survive" -ForegroundColor Cyan -NoNewLine
            Write-Host " episode from season " -ForegroundColor Yellow -NoNewLine
            Write-Host "$($Season)" -ForegroundColor Cyan -NoNewLine
            Write-Host " found...proceeding to sort..." -ForegroundColor Yellow
 
            $DriveToSurvivePath = "I:\Movies and Shows\Sports Shows\Formula 1 - Drive to Survive\Season $($Season)"
            if (!(Test-Path $DriveToSurvivePath))
            {
                New-Item $DriveToSurvivePath -Type Directory | Out-Null
                if (Test-Path $DriveToSurvivePath)
                {
                    Write-Host "    $DriveToSurvivePath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $DriveToSurvivePath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$DriveToSurvivePath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category $DriveToSurvivePath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov
            $NewFileName = "$($FileToCopy.BaseName)"
 
            foreach ($F1Email in $F1Emails)
            {
                $EmailFrom = "$SMTPUsername"
                $EmailTo = "$F1Email"
 
                SendEmail $EmailFrom $EmailTo $SMTPServer $SMTPUsername $SMTPPassword
            }
        }
        
        # Sorting MotoGP Races
 
        if ($($Show.name) -like "MotoGP*" -or $($Show.name) -like "Moto.GP.*" -OR $($Show.name) -match "MotoGP" -OR $($Show.name) -match "Qualifying.MotoGP" -OR $($Show.name) -match "Race.MotoGP")
        {
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "MotoGP" -ForegroundColor Cyan -NoNewLine
            Write-Host " event has been found...proceeding to sort... " -ForegroundColor Yellow
 
            $MotoGPPath = "I:\Movies and Shows\Sporting Events\MotoGP"
            if (!(Test-Path $MotoGPPath))
            {
                New-Item $MotoGPPath -Type Directory | Out-Null
                if (Test-Path $MotoGPPath)
                {
                    Write-Host "    $MotoGPPath" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $MotoGPPath))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$MotoGPPath!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }

            # Creating Header for Login

            $MotoGPLoginHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $MotoGPLoginHeaders.Add('Accept', 'application/json')
            $MotoGPLoginHeaders.Add('Content-Type', 'application/json')

            # Creating Body for Login

            $MotoGPJSONLogin = ConvertTo-Json(@{
                apikey = $TVDBAPIKey;
                pin = $TVDBPin;
            })

            # Logging into TheTVDB to get the API Token

            $Login = Invoke-RestMethod -Method Post -Uri "$TVDBBaseURI/login"-Header $MotoGPLoginHeaders -Body $MotoGPJSONLogin 
            $LoginToken = $Login.data.token

            # Using API Token to now search for Event

            # Creating Header for Event search

            $MotoGPEventHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $MotoGPEventHeaders.Add('Accept', 'application/json')
            $MotoGPEventHeaders.Add('Authorization', "Bearer $($LoginToken)")

            # Listing Episodes for the Series listed in the Base variables

            $MotoGPEventList = Invoke-RestMethod -Method Get -Uri "$TVDBBaseURI/series/$MotoGPTVDBSeriesID/extended?meta=episodes&short=true" -Header $MotoGPEventHeaders
            $MotoGPEvents = $MotoGPEventList.data.episodes | Where-Object {$_.aired -match $MotoGPSeason -AND $_.name -match "MotoGP"}

            # Getting the properties of each episode that fits our filter

            # Getting date range for events for hashtable filtering purposes later

            $Date = Get-Date -Format "yyyy-MM-dd"

            # Building MotoGP Hashtable

            $MGPEvents = @{}

            foreach ($MotoGPEvent in $MotoGPEvents)
            {
                #Write-Host "Event Name: $($MotoGPEvent.name)"
                #Write-Host "Event Date: $($MotoGPEvent.aired)"
                #Write-Host "Event Episode: S$($MotoGPEvent.seasonNumber)E$($MotoGPEvent.number)"
                #Write-Host "Event Full Name: $($MotoGPEvent.Value)"
                #Write-Host ""
                if ($($MotoGPEvent.number) -lt 10)
                {
                    $MotoGPEventNumber = '{0:d2}' -f $($MotoGPEvent.number) # Adding a leading zero to episodes that are single digit for sorting purposes
                    $MGPEvents."$($MotoGPEvent.name)" = "MotoGP - S$($MotoGPEvent.seasonNumber)E$MotoGPEventNumber - $($MotoGPEvent.name) ($($MotoGPEvent.aired))"
                }
                else
                {
                    $MotoGPEventNumber = '{0:d2}' -f $($MotoGPEvent.number) # Adding a leading zero for sorting purposes
                    $MGPEvents."$($MotoGPEvent.name)" = "MotoGP - S$($MotoGPEvent.seasonNumber)E$MotoGPEventNumber - $($MotoGPEvent.name) ($($MotoGPEvent.aired))"
                }
            }

            #$MGPEvents.GetEnumerator() | Sort-Object Value

            foreach ($MGPEvent in $MGPEvents.GetEnumerator())
            {
                if ($($MGPEvent.Value) -match $Date)
                {
                    #Write-Host "Event Name: $($MotoGPEvent.name)"
                    #Write-Host "Event Date: $($MotoGPEvent.aired)"
                    #Write-Host "Event Episode: S$($MotoGPEvent.seasonNumber)E$($MotoGPEvent.number)"
                    #Write-Host "Event Full Name: $($MotoGPEvent.Value)"
                    Write-Host ""
                    Write-Host "Matching MotoGP event found for this torrent file:" -ForegroundColor Yellow -NoNewline
                    Write-Host " $($MGPEvent.Value)" -ForegroundColor Cyan

                    $MotoGPFiles = Get-ChildItem -Path H:\$Category

                    Foreach ($MotoGPFile in $MotoGPFiles)
                    {
                        $SprintEvent = $MGPEvents.GetEnumerator() | Where-Object {$_.Value -match "Sprint" -AND $_.Value -match $Date} | Sort-Object Value
                        $QualifyingEvent = $MGPEvents.GetEnumerator() | Where-Object {$_.Value -match "Qualifying" -AND $_.Value -match $Date} | Sort-Object Value
                        $RaceEvent = $MGPEvents.GetEnumerator() | Where-Object {$_.Value -match "Race" -AND $_.Value -match $Date -AND $_.Value -notmatch "Sprint"} | Sort-Object Value

                        #Write-Host "Sprint Event: $($SprintEvent.Value)"
                        #Write-Host "Qualifying Event: $($QualifyingEvent.Value)"
                        #Write-Host "Race Event: $($RaceEvent.Value)"

                        if ($SprintEvent -AND $($MotoGPFile.Name) -match "Sprint")
                        {
                            Write-Host ""
                            Write-Host "  Original File Name: $($MotoGPFile.Name)"
                            $NewFileName = $($SprintEvent.Value) -replace ".{13}$"
                            Write-Host "  New File Name     : $NewFileName"
                        
                            $Extension = $MotoGPFile.Extension
                            Rename-Item -Path H:\$Category\$($MotoGPFile.Name) $($NewFileName+$Extension)
                        }
                        elseif ($QualifyingEvent -AND $($MotoGPFile.Name) -match "Qualifying")
                        {
                            Write-Host ""
                            Write-Host "  Original File Name: $($MotoGPFile.Name)"
                            $NewFileName = $($QualifyingEvent.Value) -replace ".{13}$"
                            Write-Host "  New File Name     : $NewFileName"
                        
                            $Extension = $MotoGPFile.Extension
                            Rename-Item -Path H:\$Category\$($MotoGPFile.Name) $($NewFileName+$Extension)
                        }
                        elseif ($RaceEvent)
                        {
                            Write-Host ""
                            Write-Host "  Original File Name: $($MotoGPFile.Name)"
                            $NewFileName = $($RaceEvent.Value) -replace ".{13}$"
                            Write-Host "  New File Name     : $NewFileName"
                        
                            $Extension = $MotoGPFile.Extension
                            Rename-Item -Path H:\$Category\$($MotoGPFile.Name) $($NewFileName+$Extension)
                        }

                        foreach ($MotoGPEmail in $MotoGPEmails)
                        {
                            $EmailFrom = "$SMTPUsername"
                            $EmailTo = "$MotoGPEmail"
            
                            SendEmail $EmailFrom $EmailTo $SMTPServer $SMTPUsername $SMTPPassword
                        }
                    }

                    $FilesToCopy = Get-ChildItem -Path H:\$Category
                    foreach ($FileToCopy in $FilesToCopy)
                    {
                        Robocopy H:\$Category $MotoGPPath $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov /is /it /im
                    }
                }
            }
        }
        
        # Sorting F1 Races
        
        if ($($Show.Name) -match "Survive*")
        {
            Break
        }
        elseif ($($Show.Name) -like "Formula.1.*" -or $($Show.Name) -like "Formula1.*" -or $($Show.Name) -like "Formula.One.*" -or $($Show.Name) -like "Formula1*" -or $($Show.Name) -like "Formula 1*")
        {
            Write-Host ""
            Write-Host "A " -ForegroundColor Yellow -NoNewLine
            Write-Host "Formula 1" -ForegroundColor Cyan -NoNewLine
            Write-Host " event has been found...proceeding to sort... " -ForegroundColor Yellow
 
            $F1Path = "I:\Movies and Shows\Sporting Events\Formula 1"
            if (!(Test-Path $F1Path))
            {
                New-Item $F1Path -Type Directory | Out-Null
                if (Test-Path $F1Path)
                {
                    Write-Host "    $F1Path" -ForegroundColor Cyan -NoNewLine;
                    Write-Host " directory created..." -ForegroundColor Green
                    Write-Host ""
                }
                elseif (!(Test-Path $F1Path))
                {
                    Write-Host "    Unable to create " -ForegroundColor Red -NoNewLine;
                    Write-Host "$F1Path!" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
 
            # Creating Header for Login

            $F1LoginHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $F1LoginHeaders.Add('Accept', 'application/json')
            $F1LoginHeaders.Add('Content-Type', 'application/json')

            # Creating Body for Login

            $F1JSONLogin = ConvertTo-Json(@{
                apikey = $TVDBAPIKey;
                pin = $TVDBPin;
            })

            # Logging into TheTVDB to get the API Token

            $Login = Invoke-RestMethod -Method Post -Uri "$TVDBBaseURI/login"-Header $F1LoginHeaders -Body $F1JSONLogin 
            $LoginToken = $Login.data.token

            # Using API Token to now search for Event

            # Creating Header for Event search

            $F1EventHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $F1EventHeaders.Add('Accept', 'application/json')
            $F1EventHeaders.Add('Authorization', "Bearer $($LoginToken)")

            # Listing Episodes for the Series listed in the Base variables

            $F1EventList = Invoke-RestMethod -Method Get -Uri "$TVDBBaseURI/series/$F1TVDBSeriesID/extended?meta=episodes&short=true" -Header $F1EventHeaders
            $F1Events = $F1EventList.data.episodes | Where-Object {$_.aired -match $Formula1Season}

            # Getting the properties of each episode that fits our filter

            # Getting date range for events for hashtable filtering purposes later

            $Date = Get-Date -Format "yyyy-MM-dd"

            # Building Formula 1 Hashtable

            $Formula1Events = @{}

            foreach ($F1Event in $F1Events)
            {
                #Write-Host "Event Name: $($F1Event.name)"
                #Write-Host "Event Date: $($F1Event.aired)"
                #Write-Host "Event Episode: S$($F1Event.seasonNumber)E$($F1Event.number)"
                #Write-Host "Event Full Name: $($F1Event.Value)"
                #Write-Host ""
                if ($($F1Event.number) -lt 10)
                {
                    $F1EventNumber = '{0:d3}' -f $($F1Event.number) # Adding a leading zero to episodes that are single digit for sorting purposes
                    $Formula1Events."$($F1Event.name)" = "Formula 1 - S$($F1Event.seasonNumber)E$F1EventNumber - $($F1Event.name) ($($F1Event.aired))"
                }
                else
                {
                    $F1EventNumber = '{0:d3}' -f $($F1Event.number) # Adding a leading zero for sorting purposes
                    $Formula1Events."$($F1Event.name)" = "Formula 1 - S$($F1Event.seasonNumber)E$F1EventNumber - $($F1Event.name) ($($F1Event.aired))"
                }
            }

            #$Formula1Events.GetEnumerator() | Sort-Object Value

            foreach ($Formula1Event in $Formula1Events.GetEnumerator())
            {
                if ($($Formula1Event.Value) -match $Date)
                {
                    #Write-Host "Event Name: $($Formula1Event.name)"
                    #Write-Host "Event Date: $($Formula1Event.aired)"
                    #Write-Host "Event Episode: S$($Formula1Event.seasonNumber)E$($Formula1Event.number)"
                    #Write-Host "Event Full Name: $($Formula1Event.Value)"
                    Write-Host ""
                    Write-Host "Matching Formula 1 event found for this torrent file:" -ForegroundColor Yellow -NoNewline
                    Write-Host " $($Formula1Event.Value)" -ForegroundColor Cyan

                    $F1Files = Get-ChildItem -Path H:\$Category

                    Foreach ($F1File in $F1Files)
                    {
                        $SprintEvent = $Formula1Events.GetEnumerator() | Where-Object {$_.Value -match "Sprint" -AND $_.Value -match $Date -AND $_.Value -notmatch "Shootout"} | Sort-Object Value
                        $ShootoutEvent = $Formula1Events.GetEnumerator() | Where-Object {$_.Value -match "Shootout" -AND $_.Value -match $Date} | Sort-Object Value
                        $QualifyingEvent = $Formula1Events.GetEnumerator() | Where-Object {$_.Value -match "Qualifying" -AND $_.Value -match $Date} | Sort-Object Value
                        $RaceEvent = $Formula1Events.GetEnumerator() | Where-Object {$_.Value -match "Race" -AND $_.Value -match $Date -AND $_.Value -notmatch "Sprint"} | Sort-Object Value

                        #Write-Host "Sprint Event: $($SprintEvent.Value)"
                        #Write-Host "Shootout Event: $($ShootoutEvent.Value)"
                        #Write-Host "Qualifying Event: $($QualifyingEvent.Value)"
                        #Write-Host "Race Event: $($RaceEvent.Value)"

                        if ($SprintEvent -AND $($F1File.Name) -match "Sprint" -AND $($F1File.Name) -notmatch "Shootout")
                        {
                            Write-Host ""
                            Write-Host "  Original File Name: $($F1File.Name)"
                            $NewFileName = $($SprintEvent.Value) -replace ".{13}$"
                            Write-Host "  New File Name     : $NewFileName"
                        
                            $Extension = $F1File.Extension
                            Rename-Item -Path H:\$Category\$($F1File.Name) $($NewFileName+$Extension)
                        }
                        elseif ($ShootoutEvent -AND $($F1File.Name) -match "Shootout")
                        {
                            Write-Host ""
                            Write-Host "  Original File Name: $($F1File.Name)"
                            $NewFileName = $($ShootoutEvent.Value) -replace ".{13}$"
                            Write-Host "  New File Name     : $NewFileName"
                        
                            $Extension = $F1File.Extension
                            Rename-Item -Path H:\$Category\$($F1File.Name) $($NewFileName+$Extension)
                        }
                        elseif ($QualifyingEvent -AND $($F1File.Name) -match "Qualifying")
                        {
                            Write-Host ""
                            Write-Host "  Original File Name: $($F1File.Name)"
                            $NewFileName = $($QualifyingEvent.Value) -replace ".{13}$"
                            Write-Host "  New File Name     : $NewFileName"
                        
                            $Extension = $F1File.Extension
                            Rename-Item -Path H:\$Category\$($F1File.Name) $($NewFileName+$Extension)
                        }
                        elseif ($RaceEvent)
                        {
                            Write-Host ""
                            Write-Host "  Original File Name: $($F1File.Name)"
                            $NewFileName = $($RaceEvent.Value) -replace ".{13}$"
                            Write-Host "  New File Name     : $NewFileName"
                        
                            $Extension = $F1File.Extension
                            Rename-Item -Path H:\$Category\$($F1File.Name) $($NewFileName+$Extension)
                        }

                        foreach ($F1Email in $F1Emails)
                        {
                            $EmailFrom = "$SMTPUsername"
                            $EmailTo = "$F1Email"
            
                            SendEmail $EmailFrom $EmailTo $SMTPServer $SMTPUsername $SMTPPassword
                        }
                    }

                    $FilesToCopy = Get-ChildItem -Path H:\$Category
                    foreach ($FileToCopy in $FilesToCopy)
                    {
                        Robocopy H:\$Category $F1Path $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov /is /it /im
                    }
                }
            }
        }
 
        # Moving all other new episodes
 
        $TVShowsFolderSize = Get-ChildItem "H:\$Category" | Measure-Object
        if (!($($TVShowsFolderSize).count -eq 0))
        {
            Write-Host ""
            Write-Host "Other episodes found...proceeding to sort..." -ForegroundColor Yellow
        
            & filebot -rename -r "H:\$Category" --db TheTVDB --format "{n} ({y}) - {s00e00} - {t}" -non-strict
            Start-Sleep -m 350
            $FileToCopy = Get-ChildItem -Path H:\$Category
            Robocopy H:\$Category "I:\Movies and Shows\TV Shows\New Episodes" $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov
        }
    }
}
 
function SortMovies ($TorrName)
{
    Write-Host "Now sorting movies from " -ForegroundColor Yellow -NoNewLine
    Write-Host "$DestPath" -ForegroundColor Cyan -NoNewLine
    Write-Host "..." -ForegroundColor Yellow
 
    $Movies = Get-ChildItem -Path "H:\$Category" -Recurse | Select-Object Name,DirectoryName,LastWriteTime
 
    foreach ($Movie in $Movies)
    {
        Write-Host ""
        Write-Host "Moving movies to correct directory..." -ForegroundColor Yellow
 
        & filebot -rename -r "H:\$Category" --db TheMovieDB --format "{n.colon(' - ')} {[y, certification, rating]}" -non-strict
        Start-Sleep -m 350
        $FileToCopy = Get-ChildItem -Path H:\$Category
        Robocopy H:\$Category "I:\Movies and Shows\Movies\New" $FileToCopy.Name $Ext1 $Ext2 $Ext3 /z /j /NOOFFLOAD /MT /mov
    }
}
 
function CleanupFiles ($TorrFiles)
{
    Write-Host ""
    Write-Host "Cleaning up Torrent Files..." -ForegroundColor Yellow
    Write-Host ""
 
    foreach ($TorrFile in $TorrFiles)
    {
        Write-Host "Current torrent file being checked:" -ForegroundColor Cyan
        Write-Host "  $($TorrFile.Name)"
        if ($FileDirs.Name -contains $TorrFile.BaseName)
        {
            Write-Host "      Matching directory found! Leaving .torrent file alone!" -ForegroundColor Red
        }
        elseif (Get-ChildItem -Path $fileDirectory\* -Include $Ext1,$Ext2,$Ext3 | Where-Object { $_.Name -contains $TorrFile.BaseName })
        {
            Write-Host "      Matching file found in root of Completed directory! Leaving .torrent file alone!" -ForegroundColor Red
        }
        else
        {
            Write-Host "      No matching directory found...removing torrent file!" -ForegroundColor Green
            Get-ChildItem $Location\* -Include  *.torrent -Recurse | Where-Object { $_.Name -contains $($TorrFile.Name) } | Remove-Item -Force
            Write-Host "        $TorrFile removed!" -ForegroundColor Green
        }
        Write-Host ""
    }
}
 
# Beginning processing of file(s) - moving files to temp directory, renaming contents of files to match Torrent name, unRARing properly named RAR to correct location and then cleaning up Temp folder
 
# Checking if temporary location exists and creating if it does not
 
CheckFolder $TempLoc
 
# Checking if destination location exists and creating if it does not
 
CheckFolder $DestPath
 
# Organizing based on category
 
UnRarSortToCategory $Category
 
# Sorting shows & movies and finishing the script
 
if ($Category -eq "TV Shows")
{
    SortShows $TorrName
    # Cleaning up remnant .torrent files
    CleanupFiles $TorrFiles
    Exit
}
if ($Category -eq "Movies")
{
    SortMovies $TorrName
    # Cleaning up remnant .torrent files
    CleanupFiles $TorrFiles
    Exit
}
else
{
    # Cleaning up remnant .torrent files
    CleanupFiles $TorrFiles
    Exit
}