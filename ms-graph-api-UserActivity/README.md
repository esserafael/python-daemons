## Important Considerations

To run the PowerShell script `ConvertTo-ExcelCustomReportHTML.ps1` (great name, by the way) in a Windows Server as a scheduled task, some unusual maneuvers will be needed (because of the COM Object representing the Excel app):

1. Create the directory: `C:\Windows\System32\config\systemprofile\Desktop` (if Office installed is 64 bits, else create in SySWOW64).
- The user running the task must have Modify permissions in the directory created.
2. In DCOMCNFG, right click on the My Computer and select Properties.
- Choose the COM Securities tab. 
- In Access Permissions, click "Edit Defaults" and add Network Service to it and give it "Allow local access" permission. Do the same for <Machine_name>\Users.
- In Launch and Activation Permissions, click "Edit Defaults" and add Network Service to it and give it "Local launch" and "Local Activation" permission. Do the same for <Machine_name>\Users

And that should do the trick.
