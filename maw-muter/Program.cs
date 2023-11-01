using System;
using System.Data;
using System.Diagnostics;
using System.Threading.Tasks;
using CSCore.CoreAudioAPI;

class Program
{
    static async Task Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide a command (mute or list) and optionally a process name.");
            return;
        }

        string command = args[0];

        if (command.Equals("mute", StringComparison.OrdinalIgnoreCase))
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Please provide the process name to mute or unmute as an argument.");
                return;
            }

            string targetProcessName = args[1];
            await MuteOrUnmuteProcess(targetProcessName);
        }
        else if (command.Equals("list", StringComparison.OrdinalIgnoreCase))
        {
            await ListAudioProcesses();
        }
        else
        {
            Console.WriteLine("Invalid command. Please use 'mute' or 'list'.");
        }
    }

    private static async Task MuteOrUnmuteProcess(string targetProcessName)
    {
        bool operationPerformed = false;

        using (var deviceEnumerator = new MMDeviceEnumerator())
        {
            var activeRenderDevices = deviceEnumerator.EnumAudioEndpoints(DataFlow.Render, DeviceState.Active);
            foreach (var device in activeRenderDevices)
            {
                using (var sessionManager = GetAudioSessionManager2(device, DataFlow.Render))
                {
                    var sessionEnumerator = sessionManager.GetSessionEnumerator();
                    foreach (var session in sessionEnumerator)
                    {
                        using (var audioSessionControl = session.QueryInterface<AudioSessionControl2>())
                        {
                            if (audioSessionControl.SessionState == AudioSessionState.AudioSessionStateActive) // Check if the session is active
                            {
                                string exeName = await GetExecutableName(audioSessionControl);
                                if (exeName != null && exeName.Equals(targetProcessName, StringComparison.OrdinalIgnoreCase))
                                {
                                    var simpleVolume = session.QueryInterface<SimpleAudioVolume>();
                                    if (simpleVolume != null)
                                    {
                                        simpleVolume.IsMuted = !simpleVolume.IsMuted;
                                        operationPerformed = true;
                                        break; // Add break here to exit the loop after muting
                                    }
                                }
                            }
                        }
                    }
                }

                if (operationPerformed)
                {
                    break;
                }
            }
        }

        if (!operationPerformed)
        {
            Console.WriteLine("No active audio session found for the specified process.");
        }
    }

    private static async Task ListAudioProcesses()
    {
        using (var deviceEnumerator = new MMDeviceEnumerator())
        {
            var activeDevices = deviceEnumerator.EnumAudioEndpoints(DataFlow.All, DeviceState.Active);
            foreach (var device in activeDevices)
            {
                using (var sessionManager = GetAudioSessionManager2(device, DataFlow.Render))
                {
                    var sessionEnumerator = sessionManager.GetSessionEnumerator();
                    foreach (var session in sessionEnumerator)
                    {
                        using (var audioSessionControl = session.QueryInterface<AudioSessionControl2>())
                        {
                            if (audioSessionControl.SessionState == AudioSessionState.AudioSessionStateActive) // Check if the session is active
                            {
                                string exeName = await GetExecutableName(audioSessionControl);
                                if (exeName != null)
                                {
                                    await Console.Out.WriteLineAsync(exeName);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private static async Task<string> GetExecutableName(AudioSessionControl2 audioSessionControl)
    {
        try
        {
            int processId = audioSessionControl.ProcessID;
            Process process = Process.GetProcessById(processId);
            return await Task.FromResult(process.ProcessName + ".exe");
        }
        catch
        {
            return await Task.FromResult<string>(null);
        }
    }

    private static AudioSessionManager2 GetAudioSessionManager2(MMDevice device, DataFlow dataFlow)
    {
        return AudioSessionManager2.FromMMDevice(device);
    }
}
