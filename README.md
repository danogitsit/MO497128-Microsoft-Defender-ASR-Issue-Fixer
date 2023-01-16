# MO497128-Microsoft-Defender-ASR-Issue-Fixer
PowerShell scripts to help recreate shortcuts as a result of the Microsoft Defender ASR issue.

## recover-shortcuts.ps1
This script uses Microsoft Defender events stored in Windows (Event Viewer) in an attempt to recover the shortcuts. This script does not completely recreate the shortcuts as they were as the events generated do not store information such as shortcut icon, parent path. The script must be run in the users context, the user must have read access to event viewer (usually do). This script should be used as an attempt #1 but you may be better off going straight to the next script.

Copy this script to a file on your computer called "recover-shortcuts.ps1" then, open PowerShell, use cd "C:\somedirectory\someplace\recover-shortcuts.ps1", then;

**Note. The change to Microsoft Defender rule to audit should be in place or your tenant should have recieved the Microsoft fix before you run this or Defender will just delete shortcuts again. PowerShell Constrained Language Mode not supported due to dependency on COM objects for shortcut generation.**

> .\recover-shortcuts.ps1

## generate-shortcutsinteractive.ps1
This script uses information in the registry to identify programs in which to recreate shortcuts. This is similar to a script already provided by Microsoft but different in that it can handle creating sub folders in the start menu using EXE metadata and provides a way to create either Start Menu and/or Desktop shortcuts in one go. The script must be run in the users context and is not suitable for automation as it will prompt for action. Additionally, the registry key in which installed applications are gathered can include background components you would not want shortcuts for and also, some applications do not register their location in here so it likely will not cover everything.

### Including custom shortcuts
You can include shortcuts not typically listed in the App Paths registry key by specifying an exe location in the $customAppShortcuts variable at the top of the script. This will be included should the exe be located on the machine in which the script is running. This list may be updated to include common programs as required. As an example, to include the command prompt that is then stored in a "Windows System" folder in the start menu, the $customAppShortcuts would be as follows;

> $customAppShortcuts = @{Name="Command Prompt"; Group="Windows System"; Path="C:\Windows\System32\cmd.exe"}

Copy this script to a file on your computer called "generate-shortcutsinteractive.ps1" then, open PowerShell, use cd "C:\somedirectory\someplace\recover-shortcuts.ps1", then;

**Note. The change to Microsoft Defender rule to audit should be in place or your tenant should have recieved the Microsoft fix  before you run this or Defender will just delete shortcuts again. PowerShell Constrained Language Mode not supported due to dependency on COM objects for shortcut generation.**

> .\generate-shortcutsinteractive.ps1
