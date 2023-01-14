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

write-host "Fetching Defender events..."

$events = @()

try{
    $events = Get-WinEvent -LogName "Microsoft-Windows-Windows Defender/Operational" -MaxEvents 10000 -ErrorAction Stop | where {$_.Id -eq 1121 -and $_.Message -like "*.lnk*"}
}catch{
    write-error -Message "Unable to access the event viewer events to access required logs."
    exit
}

if($events.Count -eq 0){
    write-warning -Message "No events matching this issue were identified, script will now close"
    exit
}else{
    write-host "$($events.Count) shortcut issue related events identified"
}

$eventsFormatted = foreach($event in $events){
    
    $evSplit = $event.Message.Split([Environment]::NewLine)

    New-Object -TypeName PSObject -Property @{

    eventUser = $evSplit[8].Replace("User: ","").Trim()
    eventPath = $evSplit[10].Replace("Path: ","").Trim()
    eventProc = $evSplit[12].Replace("Process Name: ","").Trim()
    eventParentCmd = $evSplit[16].Replace("Parent Commandline: ","").Trim()

    }

}

$user = $env:USERDOMAIN+"\"+$env:USERNAME

foreach($formattedevent in $eventsFormatted){

write-host "Working on shortcut for:"$formattedevent.eventPath -ForegroundColor Yellow


if($formattedevent.eventUser -eq $user){

    write-host "`tParent Command:"$formattedevent.eventParentcmd -ForegroundColor Gray
    write-host "`tArguments:"$formattedevent.eventParentcmd.Replace("""$($formattedevent.eventProc)""","") -ForegroundColor Gray

    if((Test-Path $formattedevent.eventPath) -eq $false){

        $wshshell = New-Object -ComObject WScript.Shell

        $shortcut = $wshshell.CreateShortcut($formattedevent.eventPath)
        $shortcut.TargetPath = $formattedevent.eventProc
        $shortcut.IconLocation = "$($formattedevent.eventParentcmd),0"
        $shortcut.Arguments = $formattedevent.eventParentcmd.Replace("""$($formattedevent.eventProc)""","")
        
        try{
            $shortcut.Save()
            write-host "`tShortcut creation attempt complete. If you did not apply the ASR audit this creation likely failed."
            }
        catch{

            Write-Error -Message "Shortcut creation failed due to exception. ASR may still be blocking this shortcut or the shortcut included additional parameters not saved in the event log"
            }

    }else{
        write-host "`tShortcut already existed so was not recreated"
    }

    }else{
        Write-Host "`tUsername in event $($formattedevent.eventUser) did not match current logged on user $user, shortcut skipped" -ForegroundColor Magenta
    }

}
 
