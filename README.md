# EXOMigrationInfoExport

We have had request from Microsoft Support to deliver logfiles en status of the migration of large (archive) mailboxes (more than 200Gb). The migration was done by staff outside our company but the results where not consistent 
due to various reasons. Also commands send by Microsoft did not work in the environment. 

This very simple script mitigates al lot of those issues and removes some requirements to have Powershell and procedural skills.

## How to use

This version has no loops. So three email accounts are designated in variables, where the commands are executed. Add up to three email accounts in the variables and execute the script. The script expects a c:\temp folder to be 
present and writeable.

## Whats Next (whislist)

Working on creating a loop with the mailboxes from the list so no mailboxes need to be manually added.
