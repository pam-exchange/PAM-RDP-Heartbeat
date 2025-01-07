# PAM RDP Heartbeat
Have you experienced a sometime lengthy process starting an RDP session through
your favorite PAM tool, only to find that your server will kick in with a screen
saver? To unlock the session again, you will need a password, but you do not
have it?

Nothing to it, you have to close the session and restart the process through
your PAM tool to open a new session again.

This is where **PAM RDP Heartbeat** comes into play.

The program will send heartbeat signals to the open RDP sessions, thus preventing
them to go into screen saver/lock mode.

When starting the program, it will check that the user's desktop has a screen
saver/lock active. If no active screen saver/lock is found, it will not start.

If started, a small window is opened. It sill show the open RDP sessions. For
each RDP window opened, it will show the duration of the session. In the window
you can double click on a session and the window is brought to forground. In the
settings you can configure how frequent the heartbeat signal is used and other
settings too. To show/hide the PAM-RDP-Heartbeat window to the foreground press
`Ctrl-Alt-ScrollLock`.

**PAM RDP Heartbeat** will send heartbeat signals to RDP sessions opened through
mstsc.exe (regular RDP client) and to RDP sessions opened through Symantec PAM
RDP applets.

![PAM RDP Heartbear](/docs/pam-rdp-heartbeat.png)

## Quick quide
+ `Ctrl-Alt-ScrollLock` will show/hide the PAM RDP Heartbeat window. This works regardless of focus is the PAM RDP Heartbeat window or not.
+ `Ctrl-Alt-Down` will minimize all RDP sessions (focus is PAM RDP Heartbeat window)
+ `Ctrl-Alt-Up` will show all RDP sessions (focus is PAM RDP Heartbeat window)
+ `F5` will refresh the list (focus is PAM RDP Heartbeat window)

## Installation
There is an installer program available. Run the installer as administrator and the 
programs are installed and registry is updated.

You can also merge the registry file `pam-rdp-heartbeat.reg` and run the program without
installing it first.

## Compilation
You want to compile the AutoHotKey sources yourself?

Not a problem.

You need to have AutoHotKey version 1.37.02 or newer version 1 series installed. 
It will not compile with AHK v2.

If you want to build the installer you will need Inno Setup version 6.3.3.

Run the script `build.cmd` and it will package the source as executable and create 
the installer file and package everything into a zip-file.

If you really want to control the compilation and packaging of the AutoHotKey source 
file, run the command below.

To create an executable from the sources use the Ahk2Exe utility GUI or you can
use this command (one line):

```
"c:\Opt\AutoHotKey-v1\Compiler\ahk2exe.exe"
    /base "c:\Opt\AutoHotKey-v1\Compiler\Unicode 64-bit.bin"
    /compress 2
    /icon ".\src\pam-rdp-heartbeat.ico"
    /in ".\src\pam-rdp-heartbeat.ahk"
    /out ".\bin\pam-rdp-heartbeat.exe"
```

Here AutoHotKey is installed in c:\opt\AutohotKey-v1 directory. The command
uses UPX for compressions, but you can easily do without. Just remove the /compress parameter.

## Log files
`pam-rdp-heartbeat` program will create a log file in the users %TEMP% folder. If the size
of the log file is larger than 5 MB, it will be rolled to
`pam-rdp-heartbeat.log.1`. An existing pam-rdp-heartbeat.log.1 file will be rolled
to `pam-rdp-heartbeat.log.2`, etc. Last log file kept is `pam-rdp-heartbeat.log.5`.

## Registry
Parameters to the program is stored in the Windows Registry at the key
`HKLM\PAM-Exchange\PAM-RDP-Heartbeat`.

The file `pam-rdp-heartbeat.reg` contains the registry keys created by the installer. 
If you do not use the installer, you can merge the file to your registry.

## Configuration
In the `pam-rdp-heartbeat` program you can define various settings. These settings user individual 
and are stored in the users roaming profile in the file 
`%AppData%\PAM-Exchange\PAM-RDP-Heartbeat\pam-rdp-heartbeat.properties`. 

## Security considerations
**PAM RDP Heartbeat** can be seen as a violation to the server settings having
a screen saver/lock active and you should consult your IT administrators before
using it.

## Known issues and limitations
+ When using Symantec PAM applets for RDP sessions the start time shown is the
time when the Web session to PAM was started and not when the RDP session was started,
thus the duration shown is not for the RDP session but for the PAM session.

+ RDP sessions opened through Citrix is not handled.
