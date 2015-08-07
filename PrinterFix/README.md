# PrinterCheckAndClear

A known issue with printers delivered via a 2008 r2 print server is sometimes a regkey that lists the dependent files for a printer driver ends up blank when a shared printer is delivered to a computer.

At the time there was no fix from Microsoft but a workaround was to delete the entire regkey for the affected driver, restart the print spooler, and do a gpupdate to re-download the driver. This was somewhat laborious so a script was made to automate the process, along with friendly prompts so that the pc operations team could use it.

There's a fix for this issue from Microsoft now, but the pc operations team found it didn't work so they still use this script to resolve the issue.