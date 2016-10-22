So some users take it upon themselves to create desktop shortcuts to Citrix applications (usually from Receiver created shortcuts dropped in the start menu or desktop), which is fine by itself, except if Receiver removes or refreshes its shortcut, the user created shortcut disappears. This can happen for a multitude of reasons, so if anything is being worked on or changed it happens quite frequently.

Some users will deal with it and just re-create the shortcuts. 

Others will call their IT department.

This set of commandlets was created for those other users.

Uses some methods on a shell object to read shortcuts and then either save their info to a CSV, or re-create them depending on the commandlet run. Won't delete any, but it will overwrite shortcuts with the same name. Because of the way the pin to taskbar/start menu methods work, it has to be run as the user you want to save the shortcuts as. The defaults are as generic as possible, so really it's pretty easy to use as a logon script or a batch file the users can run, so long as Powershell has been set up to run scripts. Doesn't need admin rights.
