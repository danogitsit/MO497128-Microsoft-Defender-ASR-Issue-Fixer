<#
    MIT License

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>


$appPathsRoot = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"
$appPaths = get-childitem -path $appPathsRoot

function createShortcut($path, $target){
    
    $wshshell = New-Object -ComObject WScript.Shell

    $shortcut = $wshshell.CreateShortcut($path)
    $shortcut.TargetPath = $target
    $shortcut.IconLocation = "$($target),0"
        
    try{
        $shortcut.Save()
        write-host "`tShortcut creation attempt complete. If you did not apply the ASR audit this creation likely failed."
        }
    catch{

        Write-Error -Message "Shortcut creation failed due to exception. ASR may still be blocking this shortcut"
        }

}

$results = foreach($aPath in $appPaths){
    
    $aFile = $null
    $aFilePath = $null

    try{
       $aFile = (Get-ItemProperty -Path $($appPathsRoot+$aPath.PSChildName) -Name "(default)" -ErrorAction Stop).'(default)'.Replace("""","")
    }catch{
    }

    try{
       $aFilePath = (Get-ItemProperty -Path $($appPathsRoot+$aPath.PSChildName) -Name "Path" -ErrorAction Stop).Path.Replace("""","")
    }catch{
    }

    if($aFile -and $aFilePath){

    $aFileInfo = Get-Item $aFile

    $file = New-Object -TypeName PSObject -Property @{
        Name=$aFileInfo.VersionInfo.FileDescription
        Product=$aFileInfo.VersionInfo.ProductName
        Path=$aFile
        Directory=$aFilePath
    }

    $file}

}

$resultsGroup = $results | Group-Object Product

foreach($resultGroup in $resultsGroup){

    do{

        write-host "`n"$resultGroup.Name -ForegroundColor Yellow
        write-host "--------------------------------------------------------" -ForegroundColor Yellow
        write-host "This software has $($resultGroup.Count) applications" -ForegroundColor Gray
        write-host "`tChoose shortcut types:"
        write-host "`n`t`t 0. Skip"
        write-host "`t`t 1. Start menu only"
        write-host "`t`t 2. Desktop only"
        write-host "`t`t 3. Start menu & Desktop"

        [int]$resp = read-host -Prompt "`nEnter your choice (default 0 skip)"


    }while($resp -gt 5 -or $resp -lt 0 -or $resp -eq $null)

    switch($resp){
        
        0 {
            write-host "`nResult: $($resultGroup.Name) will be skipped."
        }
        1{

            foreach($shortcut in $resultGroup.Group){

                  $shortcutDir = $env:APPDATA+"\Microsoft\Windows\Start Menu\Programs\$($resultGroup.Name)\"
                  if (!(test-path $shortcutDir)){new-item -itemtype Directory -path $($env:APPDATA+"\Microsoft\Windows\Start Menu\Programs\") -Name $resultGroup.Name | out-null}

                  $shortcutPath = $env:APPDATA+"\Microsoft\Windows\Start Menu\Programs\$($resultGroup.Name)\$($shortcut.Name.Trim()).lnk"
                  write-host "Result: Creating shortcut in Start Menu ($shortcutPath) to $($shortcut.Path)"
                  createShortcut $shortcutPath $shortcut.Path
            }    
        }
        2{
            foreach($shortcut in $resultGroup.Group){

                  $shortcutPath = [Environment]::GetFolderPath("Desktop")+"\$($shortcut.Name.Trim()).lnk"
                  write-host "Result: Creating shortcut in Desktop ($shortcutPath) to $($shortcut.Path)"
                  createShortcut $shortcutPath $shortcut.Path
                }
        }
        3{
            foreach($shortcut in $resultGroup.Group){

                  $shortcutDir = $env:APPDATA+"\Microsoft\Windows\Start Menu\Programs\$($resultGroup.Name)\"
                  if (!(test-path $shortcutDir)){new-item -itemtype Directory -path $($env:APPDATA+"\Microsoft\Windows\Start Menu\Programs\") -Name $resultGroup.Name | out-null}

                  $shortcutPathStart = $env:APPDATA+"\Microsoft\Windows\Start Menu\Programs\$($resultGroup.Name)\$($shortcut.Name.Trim()).lnk"
                  $shortcutPathDesk = [Environment]::GetFolderPath("Desktop")+"\$($shortcut.Name.Trim()).lnk"
                  write-host "Result: Creating shortcut in Start Menu & Desktop ($shortcutPathStart), ($shortcutPathDesk ) to $($shortcut.Path)"
                  
                  createShortcut $shortcutPathStart $shortcut.Path
                  createShortcut $shortcutPathDesk $shortcut.Path

                }
        }

    }

    start-sleep -Seconds 2

    cls

}
