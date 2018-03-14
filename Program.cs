using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CSCore.CoreAudioAPI;

namespace VolumeLevelMonitor
{
    class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Enter the index of the program you want to monitor as it appears in Volume Mixer and ignoring System Sounds:");
            var session_index = Int32.Parse(Console.ReadLine());
            Console.WriteLine("Enter polling rate in ms:");
            var polling_rate = Int32.Parse(Console.ReadLine());

            var volume_list = new List<float>();
            var poll_count = 0;

            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    using (var session = sessionEnumerator.ElementAt(session_index))
                    {
                        using (var audioMeterInfomation = session.QueryInterface<AudioMeterInformation>())
                        {
                            Console.Clear();
                            Console.WriteLine("Press esc to stop.");
                            Console.Write("Monitoring...");
                            do
                            {
                                while (!Console.KeyAvailable)
                                {
                                    float peak = audioMeterInfomation.GetPeakValue();
                                    volume_list.Add(peak);
                                    poll_count++;
                                    System.Threading.Thread.Sleep(polling_rate);
                                }
                            } while (Console.ReadKey(true).Key != ConsoleKey.Escape);
                        }
                    }
                }
            }
            Console.Clear();
            Console.WriteLine(poll_count.ToString() + " values captured in " + ((poll_count * polling_rate)/1000f).ToString() + "s.");
            Console.WriteLine("Name of file data will be writen to:");
            var file_path = Console.ReadLine();
            WriteToCSV(file_path, volume_list, polling_rate);

        }

        private static AudioSessionManager2 GetDefaultAudioSessionManager2(DataFlow dataFlow)
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                using (var device = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Multimedia))
                {
                    Console.WriteLine("DefaultDevice: " + device.FriendlyName);
                    var sessionManager = AudioSessionManager2.FromMMDevice(device);
                    return sessionManager;
                }
            }
        }

        private static void WriteToCSV(string file_path, List<float> list, int polling_rate)
        {
            var file_builder = new StringBuilder();
            file_builder.Append(polling_rate.ToString() + " ");
            file_builder.Append(list.Count.ToString() + "\r\n");
            foreach(var e in list)
            {
                file_builder.Append(e.ToString() + ",");
            }
            System.IO.File.WriteAllText(file_path, file_builder.ToString());

            return;
        }

    }
}
