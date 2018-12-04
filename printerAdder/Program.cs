using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;

namespace printerAdder
{
    class Program
    {
        static void Main(string[] args)
        {
            removePrinter("HP LaserJet Pro MFP M521 PCL 6");
            addPrinter("192.188.244.7","HP LaserJet Pro MFP M521 PCL 6");

            Console.ReadKey();
        }

        private static void removePrinter(string name)
        {
            bool removed = false;
            ConnectionOptions options = new ConnectionOptions();
            options.EnablePrivileges = true;
            ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath, options);
            scope.Connect();
            ManagementClass win32Printer = new ManagementClass("Win32_Printer");
            ManagementObjectCollection printers = win32Printer.GetInstances();
            Console.WriteLine("Listing printers:");

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
            Console.ResetColor();

            if (removed)
            {
                Console.WriteLine("Removed: " + name);
            }
        }

        private static void addPrinter(string server, string name)
        {
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
                    Console.Out.WriteLine("Successfully connected: " + name);
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
            }
            Console.ResetColor();
        }
    }
}
