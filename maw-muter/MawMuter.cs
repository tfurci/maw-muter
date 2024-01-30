using System;
using System.IO;
using NAudio.CoreAudioApi;
using System.Diagnostics;
using System.Linq;
using NAudio.CoreAudioApi.Interfaces;
using System.Collections.Generic;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Maw-Muter v4.0.0");
            Console.WriteLine("Please provide a command (list, mute, mute-other, mute-more, excluded) and optionally an executable name.");
            return;

        }

        string command = args[0].ToLower();

        switch (command)
        {
            case "list":
                ListActiveApps();
                break;

            case "mute":
                if (args.Length < 2)
                {
                    Console.WriteLine("Please provide the executable name to toggle mute.");
                    return;
                }

                string targetExeName = Path.GetFileNameWithoutExtension(args[1]);
                ToggleMute(targetExeName);
                break;

            case "mute-other":
                if (args.Length < 2)
                {
                    Console.WriteLine("Please provide the executable name to keep unmuted.");
                    return;
                }

                string otherExeName = Path.GetFileNameWithoutExtension(args[1]);
                SwitchMuteOther(otherExeName);
                break;

            case "mute-more":
                if (args.Length < 2)
                {
                    Console.WriteLine("Please provide the executable name to mute all instances.");
                    return;
                }

                string targetExeMoreName = Path.GetFileNameWithoutExtension(args[1]);
                MuteAllInstances(targetExeMoreName);
                break;

            case "excluded":
                DisplayExcludedApps();
                break;

            default:
                Console.WriteLine("Invalid command. Please use 'list', 'mute', 'muteother', or 'mute-more'.");
                break;
        }
    }

    private static void ListActiveApps()
    {
        using (var enumerator = new MMDeviceEnumerator())
        {
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);

            foreach (var device in devices)
            {
                var sessionManager = device.AudioSessionManager;
                var sessions = sessionManager.Sessions;

                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    int processId = (int)session.GetProcessID; // Cast uint to int
                    var process = GetProcessById(processId);

                    if (process != null)
                    {
                        Console.WriteLine($"Executable: {process.ProcessName}, Muted: {session.SimpleAudioVolume.Mute}");
                    }
                }
            }
        }
    }

    private static void ToggleMute(string targetExeName)
    {
        using (var enumerator = new MMDeviceEnumerator())
        {
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);

            foreach (var device in devices)
            {
                var sessionManager = device.AudioSessionManager;
                var sessions = sessionManager.Sessions;

                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    int processId = (int)session.GetProcessID; // Cast uint to int
                    var process = GetProcessById(processId);

                    if (process != null && process.ProcessName.Equals(targetExeName, StringComparison.OrdinalIgnoreCase))
                    {
                        session.SimpleAudioVolume.Mute = !session.SimpleAudioVolume.Mute;
                        Console.WriteLine($"{(session.SimpleAudioVolume.Mute ? "Muted" : "Unmuted")}: {process.ProcessName}");
                    }
                }
            }
        }

        Console.WriteLine($"No audio session found for the specified process: {targetExeName}");
    }
    private static void SwitchMuteOther(string targetExeName)
    {
        string[] excludedApps = ReadExcludedApps();

        using (var enumerator = new MMDeviceEnumerator())
        {
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);

            foreach (var device in devices)
            {
                var sessionManager = device.AudioSessionManager;
                var sessions = sessionManager.Sessions;

                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    int processId = (int)session.GetProcessID; // Cast uint to int

                    try
                    {
                        var process = GetProcessById(processId);

                        if (process != null)
                        {
                            string processName = process.ProcessName;
                            string processPath = process.MainModule?.FileName;

                            bool isExcluded = excludedApps.Contains(processName, StringComparer.OrdinalIgnoreCase) ||
                                              excludedApps.Contains(processPath, StringComparer.OrdinalIgnoreCase);

                            if (!isExcluded && !processName.Equals(targetExeName, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    session.SimpleAudioVolume.Mute = !session.SimpleAudioVolume.Mute;
                                    Console.WriteLine($"Switched state for: {processName}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error toggling mute state for {processName}: {ex.Message}");
                                }
                            }
                            else if (isExcluded)
                            {
                                Console.WriteLine($"Skipped (Excluded): {processName}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing session {i + 1}: {ex.Message}");
                    }
                }
            }
        }

        Console.WriteLine($"Muted/Unmuted all instances except: {targetExeName} and excluded apps");
    }


    private static void MuteAllInstances(string targetExeName)
    {
        string[] excludedApps = ReadExcludedApps();

        using (var enumerator = new MMDeviceEnumerator())
        {
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);

            foreach (var device in devices)
            {
                var sessionManager = device.AudioSessionManager;
                var sessions = sessionManager.Sessions;

                // Convert the SessionCollection to a list
                var sessionList = new List<IAudioSessionControl>();
                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    sessionList.Add(session as IAudioSessionControl);
                }

                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    int processId = (int)session.GetProcessID; // Cast uint to int
                    var process = GetProcessById(processId);

                    if (process != null)
                    {
                        bool isExcluded = excludedApps.Contains(process.ProcessName, StringComparer.OrdinalIgnoreCase);

                        if (!isExcluded && process.ProcessName.Equals(targetExeName, StringComparison.OrdinalIgnoreCase))
                        {
                            session.SimpleAudioVolume.Mute = !session.SimpleAudioVolume.Mute;
                            Console.WriteLine($"Switched state for: {process.ProcessName}");
                        }
                    }
                }
            }
        }

        Console.WriteLine($"Muted/Unmuted all instances of the specified process: {targetExeName}");
    }
    private static void DisplayExcludedApps()
    {
        string[] excludedApps = ReadExcludedApps();

        if (excludedApps.Length > 0)
        {
            Console.WriteLine("Excluded Apps:");
            foreach (var excludedApp in excludedApps)
            {
                Console.WriteLine($"- {excludedApp}");
            }
        }
        else
        {
            Console.WriteLine("No apps are currently excluded.");
        }
    }

    private static string[] ReadExcludedApps()
    {
        string excludedFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excluded.txt");

        if (File.Exists(excludedFilePath))
        {
            return File.ReadAllLines(excludedFilePath);
        }

        return new string[0];
    }

    private static Process GetProcessById(int processId)
    {
        try
        {
            return Process.GetProcessById(processId);
        }
        catch
        {
            return null;
        }
    }
}
