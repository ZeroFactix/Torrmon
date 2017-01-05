using System;
using System.Linq;
using System.IO;
using System.Diagnostics;

namespace Filemon
{
    class Program
    {

        static string sSource = "Torrmon";
        static string sLog = "Application";
        static int pCount = 0;
        public static string exLocation = @"C:\media\Ready";
        public static string exTempLocation = @"C:\media\Temp";

        static void Main(string[] args)
        {
            pCount = 0;
            //Create Event log if needed
            if (!EventLog.SourceExists(sSource))
                EventLog.CreateEventSource(sSource, sLog);
            WriteEventinfo("TorrMon Started");
            if (args.Count() < 1)
            {
                WriteEventerr("No Arguments Found. Scanning for media conversion");
                mediacheck();
            }
            else //If Directory is found the process files in directory
            {
                string wdir = args[0].ToString();

                if (Directory.Exists(wdir))
                {
                    ProcessFolder(wdir);
                }
                if (File.Exists(wdir))
                {
                    ProcessFiles(wdir);
                }

                if (pCount < 1)
                {
                    WriteEventerr("No files found to process.... Please check: " + wdir);
                }
                WriteEventinfo("TorrMon Closing");
            }

        }
        static void WriteEventinfo(string sMessage)
        {
            EventLog.WriteEntry(sSource, sMessage, EventLogEntryType.Information, 4);
        }
        static void WriteEventerr(string sMessage)
        {
            EventLog.WriteEntry(sSource, sMessage, EventLogEntryType.Error, 4);
        }

        public static void ProcessFolder(String targetdir)
        {
            ProcessFiles(targetdir);


            string[] subDirectoryEntries = Directory.GetDirectories(targetdir);
            foreach (string subDir in subDirectoryEntries)
            {
                if (!subDir.Contains("Sample") || !subDir.Contains("Subs"))
                {
                    ProcessFolder(subDir);
                }
            }

        }
        public static void ProcessFiles(String directory)
        {
            String[] Filelist = Directory.GetFiles(directory);

            foreach (string file in Filelist)
            {
                string fExt = Path.GetExtension(file).ToString().ToLower();
                //if the file is zipped then unzip it.
                if (fExt == ".rar" || fExt == ".zip")
                {
                    pCount++;
                    string wfile = Path.GetFileName(file).ToString().ToLower();
                    WriteEventinfo("Found RAR: " + wfile);
                    //Check for 7zip location
                    if (File.Exists(@"C:\Program Files\7-Zip\7z.exe"))
                    {
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.CreateNoWindow = false;
                        startInfo.UseShellExecute = false;
                        startInfo.FileName = @"C:\Program Files\7-Zip\7z.exe";
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        startInfo.Arguments = "e " + directory + @"\" + wfile + " -y -o" + exLocation;
                        try
                        {
                            // Start the process with the info we specified.
                            // Call WaitForExit and then the using-statement will close.
                            using (Process exeProcess = Process.Start(startInfo))
                            {
                                WriteEventinfo("Processing: " + wfile);
                                exeProcess.WaitForExit();
                            }
                            WriteEventinfo("Finished Processing: " + wfile);
                        }
                        catch (Exception e)
                        {

                            WriteEventerr("Error: " + e.Message.ToString());
                        }
                    }
                    else
                    {
                        WriteEventerr("Can't find 7Zip. Please check");
                    }

                }
                else if (fExt == ".jpg" || fExt == ".mkv" || fExt == ".avi" || fExt == ".mp4")
                {
                    string wFile = Path.GetFileName(file).ToString().ToLower();
                    if (!wFile.Contains("sample"))
                    {
                        pCount++;

                        WriteEventinfo("Found JPG: " + wFile);
                        WriteEventinfo("Copying: " + wFile);
                        if (fExt == ".mp4")
                        {
                            File.Copy(directory + @"\" + wFile, exLocation + @"\" + wFile, true);
                        }
                        else File.Copy(directory + @"\" + wFile, exTempLocation + @"\" + wFile, true);
                        WriteEventinfo("Finished Copying: " + wFile);
                    }
                }
            }
        }

        public static void mediacheck()
        {
            if (File.Exists(@"C:\media\Scripts\HandBrakeCLI\HandBrakeCLI.exe"))
            {
                String[] toConvert = Directory.GetFiles(exTempLocation);
                foreach (string file in toConvert)
                {
                    string fName = Path.GetFileNameWithoutExtension(file);
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.CreateNoWindow = false;
                    startInfo.UseShellExecute = false;
                    startInfo.FileName = @"C:\media\Scripts\HandBrakeCLI\HandBrakeCLI.exe";
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.Arguments = "-i " + file + " -f av_mp4 -e x264 -E av_aac -o " + exLocation + @"\" + fName + ".mp4";
                    try
                    {
                        // Start the process with the info we specified.
                        // Call WaitForExit and then the using-statement will close.
                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            WriteEventinfo("Processing: " + toConvert);
                            exeProcess.WaitForExit();
                        }
                        WriteEventinfo("Finished Processing: " + toConvert);
                    }
                    catch (Exception e)
                    {

                        WriteEventerr("Error: " + e.Message.ToString());
                    }
                }
            }
            else
            {
                WriteEventerr("Can't find HandBrake. Please check");
            }

        }
    }
}
