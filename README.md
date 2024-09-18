# PAM-RDP-Heartbeat
Have you experienced a sometime lengthy process starting an RDP session through
your favorite PAM tool, only to find that your server will kick in with a screen
saver? To unlock the session again, you will need a password, but you do not
have it?

Nothing to it, you have to close the session and restart the process through
your PAM tool to open a new session again.

This is where PAM-RDP-Heartbeat comes into play.

The program will send heartbeat signals to the open sessions, thus preventing
them to go into screen lock/saver mode.

When starting the program, it will check that the user's desktop has a screen
lock/saver active. If no active screen lock/saver is found, it will not start.

When started, a small window is opened. It sill show the open RDP sessions. For
each RDP window opened, it will show the duration of the session. In the window
you can double click on a session and the window is brought to forground. In the
settings you can configure how frequent the heartbeat signal is used and other
settings too. To show/hide the PAM-RDP-Heartbeat window to the foreground press
`Ctrl-Alt-ScrollLock`.

The PAM-RDP-Heartbeat will send heartbeat signals to RDP sessions opened through
mstsc.exe (regular RDP client) and to RDP sessions opened through Symantec PAM
RDP applets.

## Quick quide

+ `Ctrl-Alt-ScrollLock` will show/hide the PAM RDP Heartbeat window. This works regardless of focus is the PAM RDP Heartbeat window or not.
+ `Ctrl-Alt-Down` will minimize all RDP sessions (focus is PAM RDP Heartbeat window)
+ `Ctrl-Alt-Up` will show all RDP sessions (focus is PAM RDP Heartbeat window)
+ `F5` will refresh the list (focus is PAM RDP Heartbeat window)

## Installation

Just copy the executable to you favorite location.

## Compilation
You want to compile the AutoHotKey sources yourself?

Not a problem.

You need to install AutoHotKey version 1.37.02 or newer version 1 series. It will not compile with AHK v2.

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

Here the AutoHotKey is installed in c:\opt\AutohotKey-v1 directory. The command
uses UPX for compressions. If this is not available, remove the /compress
parameter.

## Log files

pam-rdp-heartbeat will create a log file in the users %TEMP% folder. If the size
of the log file is larger than 5 MB, it will be rolled to
pam-rdp-heartbeat.log.1. An existing pam-rdp-heartbeat.log.1 file will be rolled
to pam-rdp-heartbeat.log.2. This continues until pam-rdp-heartbeat.log.5, which
will be deleted when pam-rdp-heartbeat.log.4 is rolled.

## Registry

Parameters to the program is stored in the Windows Registry at the key
HKCU\PAM-Exchange\PAM-RDP-Heartbeat.

## Security considerations

The pam-rdp-heartbeat can be seen as a violation to the server settings having
a screen lock/saver active and you should consult your IT administrators before
using it.

## Known issues and limitations

+ When using Symantec PAM applets for RDP sessions the start time shown is the
time when the PAM Client was started and not when the RDP session was started,
thus the shown duration is not for the RDP session.

+ RDP sessions opened through Citrix is not handled.
