## Important Considerations

To run the PowerShell script `ConvertTo-ExcelCustomReportHTML.ps1` (great name, by the way) in a Windows Server as a scheduled task, some unusual maneuvers will be needed (because of the COM Object representing the Excel app):

1. Create the directory: `C:\Windows\System32\config\systemprofile\Desktop` (if Office installed is 64 bits, else create in SySWOW64).
- The user running the task must have Modify permissions in the directory created.
2. In DCOMCNFG, right click on the My Computer and select Properties.
- Choose the COM Securities tab. 
- In Access Permissions, click "Edit Defaults" and add Network Service to it and give it "Allow local access" permission. Do the same for <Machine_name>\Users.
- In Launch and Activation Permissions, click "Edit Defaults" and add Network Service to it and give it "Local launch" and "Local Activation" permission. Do the same for <Machine_name>\Users

And that should do the trick.


### Thanks to these guys:

[[Stack Overflow] Microsoft Office Excel cannot access the file](https://stackoverflow.com/questions/7106381/microsoft-office-excel-cannot-access-the-file-c-inetpub-wwwroot-timesheet-app)

[[Microsoft Community] Issue with Powershell running through Task scheduler](https://answers.microsoft.com/en-us/windows/forum/all/issue-with-powershell-running-through-task/8d77b7fe-93a3-4363-8e8c-a6cd5deb0284)
