# printerAdder

C# program to add/remove printers shared over a network.

# usage
```
printerAdder <options>
options:
-s <ip>, --server <ip>       sets the IP of the server, where the printer lives
-p <name>, --printer <name>  sets the name of the printer
-a, --add                    just adds the printer, no removal
-r, --remove                 just removes the printer, no addition
-u, --user                   no user input in the end
-h, --help                   shows this text
```
e.g.:
```
printerAdder -s 192.188.244.7 -p "HP LaserJet Pro MFP M521 PCL 6"
```
First removes **HP LaserJet Pro MFP M521 PCL 6** from the system (if exists), then tries to add **\\\192.188.244.7\HP LaserJet Pro MFP M521 PCL 6**.
Lastly it waits for an input from the user before exit.
