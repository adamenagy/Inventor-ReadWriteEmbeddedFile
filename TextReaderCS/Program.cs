using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using System.Configuration;
using System.Reflection; 

namespace TextReaderCS
{
  class Program
  {
    const string kOriginalCommandString = "OriginalCommandString";
    const string kSuppressDialog = "SuppressDialog";
    const string kCommandKey = @"Software\Classes\txtfile\shell\open\command";

    static void ModifyRegistry(bool register)
    {
      // HKEY_CLASSES_ROOT\textfile\shell\open\command [default]
      // controls which program is used to edit txt files
      // But the CURRENT_USER version overrides it and writing that does not require 
      // administrator rights or "Run as administrator":
      // HKEY_CURRENT_USER\Software\Classes\txtfile\shell\open\command
      // This key might not exist so we have to handle that
      RegistryKey commandKey = Registry.CurrentUser.OpenSubKey(
        kCommandKey, true);

      /*
      // HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\AttachmentExecute\{F20DA720-C02F-11CE-927B-0800095AE340}
      // lists the extensions for which the "Do you want to open this file?" dialog is skipped/suppressed
      RegistryKey showDialogKey = Registry.CurrentUser.OpenSubKey(
        @"Software\Microsoft\Windows\Shell\AttachmentExecute\{F20DA720-C02F-11CE-927B-0800095AE340}", true);
       */

      if (register)
      {
        // Let's store the current value so that we can restore it 
        // when we unregister our app
        string commandString;
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        try
        {
          commandString = commandKey.GetValue(null) as string;
          config.AppSettings.Settings.Remove(kOriginalCommandString);
          config.AppSettings.Settings.Add(kOriginalCommandString, commandString);
        }
        catch 
        { 
          // Otherwise set it to empty string
          config.AppSettings.Settings.Remove(kOriginalCommandString);
          config.AppSettings.Settings.Add(kOriginalCommandString, "");
        }
        config.Save(ConfigurationSaveMode.Modified);
        
        // Get the exe path
        string codeBase = Assembly.GetExecutingAssembly().CodeBase;
        UriBuilder uri = new UriBuilder(codeBase);
        string exePath = Uri.UnescapeDataString(uri.Path);
        exePath = exePath.Replace('/', '\\');

        // Set the registry entry to use our app
        commandString = string.Format("\"{0}\" \"%1\"", exePath);
        if (commandKey == null)
        {
          // Create key
          commandKey = Registry.CurrentUser.CreateSubKey(kCommandKey); 
        }

        commandKey.SetValue(null, commandString);

        // Now suppress the "Do you want to open this file?" dialog 
        // Did not seem to be needed in the end
      }
      else
      {
        // We have to restore the original value
        string originalCommandString = ConfigurationManager.AppSettings.Get(kOriginalCommandString);

        // If we have an empty string then the key did not previously exist
        if (originalCommandString == "")
        {
          Registry.CurrentUser.DeleteSubKey(kCommandKey); 
        }
        else
        {
          commandKey.SetValue(null, originalCommandString);
        }
      }
    }

    static void EditTextFile(string filePath)
    {
      // Write to file using "append = true"
      using (StreamWriter outputFile = new StreamWriter(filePath, true))
      {
        outputFile.WriteLine(System.DateTime.Now.ToString());
        outputFile.Flush();
      }
    }

    static void Main(string[] args)
    {
      // If you want to debug into the code when the exe is used from
      // outside e.g. from VBA, then just call Break() somewhere
      // Debugger.Break();

      // Nothing to do
      if (args.Length != 1)
        return;

      switch (args[0])
      {
        case "/r": 
          // Register
          ModifyRegistry(true);
          break;

        case "/u": 
          // Unregister
          ModifyRegistry(false);
          break;

        default:   
          // Edit text file
          EditTextFile(args[0]);
          break;
      }
    }
  }
}
