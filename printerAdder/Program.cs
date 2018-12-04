using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Security.Principal;

namespace printerAdder
{
    class Program
    {
        static private string _printer = "HP LaserJet Pro MFP M521 PCL 6";
        static private string _server = "192.188.244.7";
        static bool _add = true, _remove = true, _user = true;

        static void Main(string[] args)
        {
            Console.WriteLine("--==[ printerAdder ]==-- -[ 2018 ]- -[ gmb ]-\n");

            for(int a = 0; a < args.Length; a++)
            {
                if(args[a] == "-s" || args[a] == "--server")
                {
                    _server = args[++a];
                } else if (args[a] == "-p" || args[a] == "--printer")
                {
                    _printer = args[++a];
                } else if (args[a] == "-a" || args[a] == "--add")
                {
                    _remove = false;
                } else if (args[a] == "-r" || args[a] == "--remove")
                {
                    _add = false;
                } else if (args[a] == "-u" || args[a] == "--user")
                {
                    _user = false;
                } else if(args[a] == "-h" || args[a] == "--help")
                {
                    _showHelp();
                    Environment.Exit(0);
                }
            }

            if (_isAdmin())
            {
                if (_remove)
                {
                    _removePrinter(_printer);
                }

                if (_add)
                {
                    _addPrinter(_server, _printer);
                }
            }
           

            if (_user)
            {
                Console.WriteLine("Press any key to exit!");
                Console.ReadKey();
            }
        }

        private static void _showHelp()
        {
            Console.WriteLine("usage: printerAdder <options>\n\n" +
                              "options:\n" +
                              "-s <ip>, --server <ip>       sets the IP of the server, where the printer lives\n" +
                              "-p <name>, --printer <name>  sets the name of the printer\n" +
                              "-a, --add                    just adds the printer, no removal\n" +
                              "-r, --remove                 just removes the printer, no addition\n" +
                              "-u, --user                   no user input in the end\n" +
                              "-h, --help                   shows this text\n");
        }

        private static bool _isAdmin()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            bool isElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);

            if (isElevated)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Admin rights checked.");
            } else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("User has NO admin rights, please run the program 'as admin'.");
            }
            Console.ResetColor();
            return isElevated;
        }

        private static void _removePrinter(string name)
        {
            Console.WriteLine("Removing " + name + " ...");
            bool removed = false;
            ConnectionOptions options = new ConnectionOptions();
            options.EnablePrivileges = true;
            ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath, options);
            scope.Connect();
            ManagementClass win32Printer = new ManagementClass("Win32_Printer");
            ManagementObjectCollection printers = win32Printer.GetInstances();
            Console.WriteLine("List of printers:");

            foreach (ManagementObject printer in printers)
            {
                if (printer["DeviceID"].ToString().Contains(name))
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    printer.Delete();
                    removed = true;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                }
                Console.WriteLine("  " + printer["DeviceID"]);
            }
           
            if (removed)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Removed: " + name);
            } else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("No printer was removed.");
            }

            Console.ResetColor();
        }

        private static void _addPrinter(string server, string name)
        {
            Console.WriteLine("Adding " + name + "@" + server + " ...");
            ManagementClass win32Printer = new ManagementClass("Win32_Printer");
            ManagementBaseObject inputParam = win32Printer.GetMethodParameters("AddPrinterConnection");
            inputParam.SetPropertyValue("Name", "\\\\" + server + "\\" + name);
            ManagementBaseObject result = (ManagementBaseObject)win32Printer.InvokeMethod("AddPrinterConnection", inputParam, null);
            uint errorCode = (uint)result.Properties["returnValue"].Value;

            if(errorCode == 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
            } else
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }

            switch (errorCode)
            {
                case 0:
                    Console.Out.WriteLine("Successfully connected: " + name + "@" + server);
                    break;
                case 5:
                    Console.Out.WriteLine("Access Denied.");
                    break;
                case 123:
                    Console.Out.WriteLine("The filename, directory name, or volume label syntax is incorrect.");
                    break;
                case 1801:
                    Console.Out.WriteLine("Invalid Printer Name.");
                    break;
                case 1930:
                    Console.Out.WriteLine("Incompatible Printer Driver.");
                    break;
                case 3019:
                    Console.Out.WriteLine("The specified printer driver was not found on the system and needs to be downloaded.");
                    break;
                default:
                    Console.Out.WriteLine("Unknown error. No printer was added.");
                    break;
            }
            Console.ResetColor();
        }
    }
}
