using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using WUApiLib;

// Just going to preface this with I haven't coded in ~2-3 years. And I started again with WinAPI and COM. Ugh.

namespace CleanBadUpdates
{
    internal class Program
    {
        private static UpdateSession _session;
        private static IUpdateSearcher _updateSearcher;
        private static AsyncUpdate _callback;
        private static List<IUpdate> _badUpdates;
        private static bool _done;

        private static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException +=
                (sender, eventArgs) => { ExceptionHandler((Exception) eventArgs.ExceptionObject); };

            if (Environment.OSVersion.Platform == PlatformID.Win32NT && Environment.OSVersion.Version.Major == 10)
            {
                Console.WriteLine("Wait a second, you are RUNNING windows 10! You shouldn't be running this!");
                return;
            }
            try
            {
                Console.WriteLine("Creating Windows Update Session...");
                // Because who needs a .Create() method for a COM interop object. Right Microsoft?
                Type t = Type.GetTypeFromProgID("Microsoft.Update.Session");
                _session = (UpdateSession) Activator.CreateInstance(t);
                Console.WriteLine("Creating Search Object...");
                _updateSearcher = _session.CreateUpdateSearcher();
                Console.WriteLine("Starting search for all updates (This might take a while)...");
                // Not joking, takes forever.
                _callback = new AsyncUpdate();
                _updateSearcher.BeginSearch("(IsInstalled=1) OR (IsInstalled=0 AND IsHidden=0)",
                    _callback, null);
                while (!_done)
                    Thread.Sleep(100);
            }
            catch (Exception exception)
            {
                ExceptionHandler(exception);
            }

        }

        private static void ExceptionHandler(Exception e)
        {
            Console.WriteLine(e.ToString());
            _done = true;
        }

        private static void EndSearch(ISearchJob job)
        {
            try
            {
                Console.WriteLine("Retrieving results...");
                ISearchResult searchResult = _updateSearcher.EndSearch(job);

                List<string> badKb = new List<string>
                {
                    // Nagware
                    "KB3035583",
                    "KB3075249",
                    "KB3080149",
                    "Upgrade to Windows 10",
                    // Telemetry / Windows 10 compatability. I hope I got all of these.
                    "KB3022345",
                    "KB3068708",
                    "KB3075249",
                    "KB3080149",
                    "KB2952664",
                    "KB2976978",
                    "KB2977759",
                    "KB3050265",
                    "KB3050267",
                    "KB3068708"
                };

                Console.WriteLine("Searching for Nagware and telemetry patches...");
                //LINQin it up
                _badUpdates =
                    searchResult.Updates.Cast<IUpdate>().Where(update => badKb.Any(update.Title.Contains)).ToList();
                int installedBadUpdates = _badUpdates.Count(update => update.IsInstalled);
                Console.WriteLine("Found a total of {0} updates, {1} of which are installed.",
                    _badUpdates.Count, installedBadUpdates);

                if (installedBadUpdates > 0)
                {
                    Console.WriteLine("Preparing to uninstall updates...");
                    List<string> installedKBs = new List<string>(installedBadUpdates + 1);
                    foreach (IUpdate badUpdate in _badUpdates.Where(badUpdate => badUpdate.IsInstalled))
                    {
                        installedKBs.Add(badUpdate.Title.Substring(badUpdate.Title.LastIndexOf("KB") + 2, 7));
                        Console.WriteLine("Enumberating {0}...", badUpdate.Title);
                    }

                    Console.WriteLine("Uninstalling updates...");
                    Process process = null;
                    for (int i = 0; i < installedKBs.Count; i++)
                    {
                        while (process != null && !process.HasExited)
                        {
                            Thread.Sleep(1000);
                        }
                        ProcessStartInfo start = new ProcessStartInfo("wusa.exe",
                            string.Format("/uninstall /KB:{0} /norestart", installedKBs[i]))
                        {
                            RedirectStandardError = true,
                            RedirectStandardOutput = true,
                            UseShellExecute = false
                        };
                        process = Process.Start(start);
                        Console.WriteLine("Removing KB{0} ({1}/{2})", installedKBs[i], i + 1, installedKBs.Count);
                    }
                    while (process != null && !process.HasExited)
                    {
                        Thread.Sleep(1000);
                    }
                    Console.WriteLine("All done.");
                    HideUpdates();
                }
                else
                {
                    Console.WriteLine("No updates installed, thankfully.");
                    HideUpdates();
                }
            }
            catch (Exception exception)
            {
                ExceptionHandler(exception);
            }
        }

        private static void HideUpdates()
        {
            try
            {
                Console.WriteLine("Hiding updates from Windows Update so it doesn't download them again...");
                foreach (IUpdate update in _badUpdates.Where(update => !update.IsHidden))
                {
                    update.IsHidden = true;
                    Console.WriteLine("Hid {0}...", update.Title);
                }
                Console.WriteLine("{0} updates set as hidden.", _badUpdates.Count);
                KillGwx();
            }
            catch (Exception exception)
            {
                ExceptionHandler(exception);
            }
        }

        private static void KillGwx()
        {
            try
            {
                Console.WriteLine("Killing GWX...");
                foreach (Process process in Process.GetProcessesByName("GWX"))
                {
                    process.Kill();
                    process.WaitForExit();
                    Console.WriteLine("Killed GWX.exe (PID {0})", process.Id);
                }
                RemoveWindows10Data();
            }
            catch (Exception exception)
            {
                ExceptionHandler(exception);
            }
        }

        private static void RemoveWindows10Data()
        {
            try
            {

                Console.WriteLine("Checking for $WINDOWS.~BT...");
                long data = 0;
                if (Directory.Exists("C:\\$WINDOWS.~BT"))
                {
                    Console.WriteLine("Removing $WINDOWS.~BT...");
                    Console.WriteLine("");
                    foreach (
                        string file in Directory.EnumerateFiles("C:\\$WINDOWS.~BT", "*", SearchOption.AllDirectories))
                    {
                        FileInfo info = new FileInfo(file);
                        data += info.Length;
                        info.Delete();
                        Console.WriteLine("\rDeleting: {0}", CompactPath(file, Console.WindowWidth - 11));
                    }
                    // lol forgot to delete the directories
                    Directory.Delete("C:\\$WINDOWS.~BT", true);

                    // stole it
                    string[] sizes = {"B", "KB", "MB", "GB"};
                    double len = data;
                    int order = 0;
                    while (len >= 1024 && order + 1 < sizes.Length)
                    {
                        order++;
                        len = len/1024;
                    }

                    Console.WriteLine("Done. Deleted {0:N2} {1} of Windows 10 data.", len, sizes[order]);
                }
                else
                {
                    Console.WriteLine("Good, no windows 10 download found.");
                }
                Console.WriteLine("All finished. Make sure to restart windows and run this program again, to make sure everything's cleaned out!");
                Console.WriteLine("Press any key to continue...");
                Console.Read();
            }
            catch (Exception exception)
            {
                ExceptionHandler(exception);
            }
            _done = true;
        }

        #region Stolen Code

        [DllImport("shlwapi.dll", CharSet = CharSet.Auto)]
        private static extern bool PathCompactPathEx(
            [Out] StringBuilder pszOut, string szPath, int cchMax, int dwFlags);

        public static string CompactPath(string longPathName, int wantedLength)
        {
            // NOTE: You need to create the builder with the 
            //       required capacity before calling function.
            // See http://msdn.microsoft.com/en-us/library/aa446536.aspx
            StringBuilder sb = new StringBuilder(wantedLength + 1);
            PathCompactPathEx(sb, longPathName, wantedLength + 1, 0);
            return sb.ToString();
        }

        #endregion

        protected class AsyncUpdate : ISearchCompletedCallback
        {
            public void Invoke(ISearchJob searchJob, ISearchCompletedCallbackArgs callbackArgs)
            {
                EndSearch(searchJob);
            }
        }
    }
}
