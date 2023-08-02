# EXOMigrationInfoExport V2

We have had request from Microsoft Support to deliver logfiles en status of the migration of large (archive) mailboxes (more than 200Gb). The migration was done by staff outside our company but the results where not consistent 
due to various reasons. Also commands send by Microsoft did not work in the environment. 

This very simple script mitigates al lot of those issues and removes some requirements to have Powershell and procedural skills.

## How to use

This requires that the latest version of the EXO Module is present and the Active Directory Module. As most admin have the Active Directory Module installed, there is no active check. For the EXO module it checks if the latest Module is installed, and if not, it will install the Module.

The script expects a c:\temp folder to be present and writeable.

Before running the script you can add multiple usernames in the variable $UserMailBoxAddress. 

The script will loop them through and generate a minimum of 3 text files per mailbox and 3 XML files per mailbox.
