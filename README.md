# Office365-send-as

A PowerShell script to set Send As permissions for Office365 users and groups

## Usage

### Prerequisites

- The script uses the `ExchangeOnlineManagement` PowerShell module, which requires a minimum PowerShell version of 7.0.3 to be installed. Powershell 7 is installed by default on any Windows 10 or newer PC released after March 2020. The `ExchangeOnlineManagement` module will install automatically as part of the script if not already present.
- The PowerShell execution policy must be set to run unsigned scripts. While running the script may automatically prompt to change the execution policy, it can be changed in PowerShell by running the command `Set-ExecutionPolicy Unrestricted`.
  - Execution policy can be reset after running the script using `Set-ExecutionPolicy RemoteSigned`

### Running the Script

- The script can be run by right clicking the `setSendAs.ps1` file and choose “Run with PowerShell,” but I would advise running the script within a PowerShell window to monitor the output.
  - Steps to run it in an open PowerShell window:
    - Hit the Windows key and the R key together to open a `Run` dialog box. Type “powershell” in the box and hit the Enter key.
    - If execution policy hasn't been set to `Unrestricted` yet, type `Set-ExecutionPolicy Unrestricted` and hit the Enter key.
    - Drag and drop the script file into the window to paste the path to the script in the shell. Hit the Enter key.
  - It will ask for the admin user email to authenticate to Exchange Online, the user email getting SendAs Trustee permissions, whether setting an individual user `U` or all users in a Office365 group `G`, and the user email or Office 365 group to assign the permission to. Fill in the prompts as desired to set the parameters.
  - A Microsoft sign-in window will appear to authenticate with Exchange Online once the parameters are provided.

## License

This project is licensed under the terms of the MIT license.
