using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;

namespace BotWatcher
{
    public enum ServiceState
    {
        SERVICE_STOPPED = 0x00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ServiceStatus
    {
        public int dwServiceType;
        public ServiceState dwCurrentState;
        public int dwControlsAccepted;
        public int dwWin32ExitCode;
        public int dwServiceSpecificExitCode;
        public int dwCheckPoint;
        public int dwWaitHint;
    };
    public partial class WatcherService : ServiceBase
    {
        private bool started, pending;
        private EventLog el;
        private static readonly object _sync = new object();

        private delegate void launchDelegate(string filePath, Guid identifier);

        public WatcherService()
        {
            InitializeComponent();
            started = false; //This tracks whether the service has started
            pending = false; //This tracks whether a bot is in progress
            el = new EventLog();
            if (!EventLog.SourceExists("BotWatcher"))
            {
                EventLog.CreateEventSource(
                    "BotWatcher", "BotWatcherLog");
            }
            el.Source = "BotWatcher";
            el.Log = "BotWatcherLog";
        }

        protected override void OnStart(string[] args)
        {
            el.WriteEntry("BotWatcher Service Begins");
            //Keep ChxbotPath in memory
            try
            {
                // Update the service state to Start Pending.
                el.WriteEntry("BotWatcher service successfully started");
                Run(); //Run watches a directory for an input file until the service stops
            }
            catch (Exception x)
            {
                started = false;
                string fail = $"FATAL ERROR.  The monitoring process failed with the following error:  {x.Message}";
                el.WriteEntry(fail, EventLogEntryType.Error);
                return;
            }
        }

        protected override void OnStop()
        {
            started = false; //Signal directory watcher to stop

            // Update the service state to Stop Pending.
            el.WriteEntry("BotWatcher Service was stopped");
        }

        //Microsoft Flow will create a new file in the Activators directory when a relevant email arrives
        //Watch for this file, and act on it when it arrives
        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        private void Run()
        {
            el.WriteEntry("Monitoring");
            try
            {
                //Arrange the Monitor (watcher)
                FileSystemWatcher watcher = new FileSystemWatcher();
                watcher.Path = @"C:\Flow\Input"; //Something I read from the registry

                // Watch for changes in LastAccess and LastWrite times, and
                // the renaming of files or directories.
                watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;

                // Only watch *.input files.  Files will be of the form {guid}.input
                watcher.Filter = @"*.input"; // $"*{cfg.RPAInputExtension}"; //Something I read from the registryrf

                // Add event handlers.
                watcher.Created += OnNewFileReceived;

                // Begin watching.
                watcher.EnableRaisingEvents = true;

                //Stay Resident
                while (started)
                {
                    System.Threading.Thread.Sleep(100);
                }
            }
            catch (Exception x)
            {
                el.WriteEntry($"Directory watcher failed with the following error:  {x.Message}", EventLogEntryType.Error);
                throw x;  //This is fatal, because the Run() method failed.
            }
        }


        // When a file arrives, launch a single bot asynchronously, so that monitoring may continue.
        private void OnNewFileReceived(object source, FileSystemEventArgs e)
        {
            try
            {
                string triggerPath = e.FullPath; //This is the path to the {guid}.input file
                el.WriteEntry($"Received new file {triggerPath}.");
                FileInfo fi = new FileInfo(triggerPath);
                string[] fileParts = fi.Name.Split('.');
                Guid candidate = Guid.Empty;
                try
                {
                    candidate = Guid.Parse(fileParts[0]);
                }
                catch
                {
                    el.WriteEntry($"Skipping Invalid file \"{triggerPath}\".", EventLogEntryType.Warning);
                    return;
                }
                //So now you have a file names {guid}.input.  Read it and make sure it's valid
                //The following is how to asynchronously make a C# method call. 
                AsyncCallback cb = new AsyncCallback(onBotComplete);               //The callback fires when the bot is complete.

                //With everything set up, the following calls createBot and returns immediately to monitoring
                //1.  botPath is the full path of the {guid}.input file
                //2.  out fileName is just the filename portion.  This is done so that the callback can correlate the input with the output
                //3.  cb is the callback method (onBotsComplete)
                //4.  ld is the launch delegate object again. 
                //IAsyncResult ar = ld.BeginInvoke(botPath, out fileName, cb, ld);
                launchDelegate ld = new launchDelegate(launchIntermediate);                 //The delegate, declared above, can just take a method name
                IAsyncResult ar = ld.BeginInvoke(triggerPath, candidate, cb, ld);
            }
            catch (Exception x) //Log and continue so that no particular input file can gum up the works
            {
                el.WriteEntry($"Unexpected error in method \"OnNewFileReceived\":  {x.Message}", EventLogEntryType.Error);
            }
        }

        private void launchIntermediate(string filePath, Guid identifier)
        {
            lock (_sync)
            {
                el.WriteEntry("launchIntermediate:  Entry gained");
                Process p = new Process();
                string botName, argString;
                try
                {
                    el.WriteEntry("launchIntermediate try block:  Entry gained", EventLogEntryType.Information, 22333);
                    parseInputFile(filePath, out botName, out argString);
                    el.WriteEntry($"The following input arguments have been parsed for bot {botName}:  {argString}");
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    FileInfo fi = new FileInfo(@"C:\Flow\Launch\wfLauncher.exe");
                    if (!fi.Exists)
                        throw new Exception("Dude, cannot find the launcher");

                    SecureString ss = new SecureString();
                    foreach (char c in "1234Quick")
                        ss.AppendChar(c);
                    startInfo.WorkingDirectory = fi.DirectoryName;
                    startInfo.FileName = fi.FullName;
                    startInfo.UseShellExecute = true;
                    startInfo.Arguments = $"{botName} {identifier} {argString}";
                    //startInfo.UserName = @"clyde\thurber";
                    //startInfo.Password = ss;
                    //el.WriteEntry($"Workflow {botName} ready to start with the following command: {cfg.chxbotPath} {startInfo.Arguments}");
                    //* Set output and error (asynchronous) handlers
                    p.StartInfo = startInfo;
                    el.WriteEntry("Ready to run afLauncher.exe", EventLogEntryType.Information, 22333);
                    p.Start(); //Asynchronous, returns immediately.  Do not continue until you see the output file.

                    el.WriteEntry("OK afLauncher.exe should have started", EventLogEntryType.Information, 22333);
                    Continue(botName, identifier); //Since p returns immediately, this spins until the bot is complete to ensure that only one runs at once.
                }
                catch(Exception x)
                {
                    el.WriteEntry($"Attempting to launch bot associated with file \"{filePath}\", received error {x.Message}.  Skipping...", EventLogEntryType.Error, 22333); ;
                }
            }
        }

        //Get the command string and the botname out of the file which has the form {botname}|argstring
        private void parseInputFile(string path, out string botName, out string argString)
        {
            botName = string.Empty;
            argString = string.Empty;
            string content = File2String(path);
            int barPos = content.IndexOf('|');
            //If we found no bar, then the file should be a single word--the botname
            if (barPos == -1)
            {
                botName = content.Trim();
                return;
            }
            else
            {
                botName = content.Substring(0, barPos).Trim();
                argString = content.Substring(barPos + 1).Trim();
                if (string.IsNullOrWhiteSpace(argString)) //Nothing after the '|' character
                    argString = string.Empty;  //Add the botGuid which all bots must accept
            }

            if (!validBotName(botName)) throw new Exception($"Found \"{botName}\" as name of bot.  Cannot proceed");
            if (!validArgString(argString)) throw new Exception($"Found \"{argString}\" as argument string.  Cannot proceed");
        }


        //Read a file into a string
        private string File2String(string path)
        {
            string result = string.Empty;
            try
            {
                //Block if the file is still being written
                while (FileinUse(path))
                {
                    el.WriteEntry(string.Format("File {0} in use.  Waiting...", path));
                    System.Threading.Thread.Sleep(100);
                }
                using (StreamReader streamReader = new StreamReader(path, Encoding.UTF8))
                {
                    result = streamReader.ReadToEnd();
                }
            }
            catch (System.IO.IOException x)
            {
                el.WriteEntry($"Error reading file {path}.  Error thrown is {x.Message}.", EventLogEntryType.Error);
                //If the file is in use, wait for a bit and try again.
                System.Threading.Thread.Sleep(1000);
                File2String(path);
            }
            return result;
        }
        //Determine whether the file is available
        protected virtual bool FileinUse(string path)
        {
            FileStream stream = null;
            FileInfo fi = new FileInfo(path);

            try
            {
                stream = fi.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //1.  still being written to
                //2.  being processed by another thread
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }


        //We calculated a bot name.  It should contain only letters and numerals--no spaces or puncuation or special characters
        private bool validBotName(string candidate)
        {
            string pattern = @"^[\w]+$";
            Match m = Regex.Match(candidate, pattern);
            return m.Success;
        }

        //We calculated an argstring.  It should be of the form key1=val1,key2=val2,key3=val3,...key[n]=val[n]
        private bool validArgString(string candidate)
        {
            try
            {
                if (candidate.Trim() == string.Empty)
                    return true;  //Valid to have a bar with nothing after it.

                string keyPattern = @"^[\w]+$";     //Key may not contain spaces
                string valPattern = @"^[\w ]+$";    //Value may contain spaces

                //candidate should look like this:  var1=val1^var2=val2^var3=val3....
                string[] keyvaluepairs = candidate.Split('^');
                foreach (string kv in keyvaluepairs)
                {
                    if (!kv.Contains("=")) return false;
                    string[] KeyValue = kv.Split('=');
                    if (KeyValue.Length != 2) return false;
                    if (!Regex.Match(KeyValue[0].Trim(), keyPattern).Success) return false;
                    if (!Regex.Match(KeyValue[1].Trim(), valPattern).Success) return false;
                }
            }
            catch (Exception x)
            {
                el.WriteEntry($"Unexpected exception in method \"validArgString\" with candidate {candidate}.  Returning false and failing this bot.  Error:  {x.Message}", EventLogEntryType.Error);
            }
            return true;
        }

        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        private void Continue(string botName, Guid identifier)
        {
            el.WriteEntry($"Waiting for completion of bot {botName} with id {identifier}");
            try
            {
                pending = true; //Keep spinning until the file shows up
                string filename = $"{identifier.ToString().Replace("-", "")}.output";
                el.WriteEntry($"Waiting for file {filename}", EventLogEntryType.Information, 22333);
                //Arrange the Monitor (watcher)
                FileSystemWatcher watcher = new FileSystemWatcher();
                watcher.Path = @"C:\Flow\Output"; //Something I read from the registry

                // Watch for changes in LastAccess and LastWrite times, and
                // the renaming of files or directories.
                watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;

                // Waiting for file {guid}.output
                watcher.Filter = filename; 

                // Add event handlers.
                watcher.Created += OnOutputReceived;

                // Begin watching.
                watcher.EnableRaisingEvents = true;

                //Return when the file shows up
                while (pending)
                {
                    System.Threading.Thread.Sleep(300);
                }
                el.WriteEntry($"Bot {botName} with identifier {identifier} is complete.");
            }
            catch (Exception x)
            {
                el.WriteEntry($"Directory watcher failed with the following error:  {x.Message}", EventLogEntryType.Error);
                throw x;  //This is fatal, because the Run() method failed.
            }
        }

        private void OnOutputReceived(object source, FileSystemEventArgs e)
        {
            el.WriteEntry("File Received", EventLogEntryType.Information, 22333);
            pending = false; //Kills the continue loop and completes bot.
        }


        //Single bot has finished
        private void onBotComplete(IAsyncResult ar)
        {
            try
            {
                launchDelegate ld = (launchDelegate)ar.AsyncState;
                ld.EndInvoke(ar);
            }
            catch (Exception x)
            {
                el.WriteEntry($"Method onBotComplete failed with error:  {x.Message}");
            }
        }
    }
}
