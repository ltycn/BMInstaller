using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using System.Security;

namespace BMInstaller
{
    class Program
    {
        static void Main(string[] args)
        {
            var logo = @"
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

██████╗ ███╗   ███╗██╗███╗   ██╗███████╗████████╗ █████╗ ██╗     ██╗     ███████╗██████╗ 
██╔══██╗████╗ ████║██║████╗  ██║██╔════╝╚══██╔══╝██╔══██╗██║     ██║     ██╔════╝██╔══██╗
██████╔╝██╔████╔██║██║██╔██╗ ██║███████╗   ██║   ███████║██║     ██║     █████╗  ██████╔╝
██╔══██╗██║╚██╔╝██║██║██║╚██╗██║╚════██║   ██║   ██╔══██║██║     ██║     ██╔══╝  ██╔══██╗
██████╔╝██║ ╚═╝ ██║██║██║ ╚████║███████║   ██║   ██║  ██║███████╗███████╗███████╗██║  ██║
╚═════╝ ╚═╝     ╚═╝╚═╝╚═╝  ╚═══╝╚══════╝   ╚═╝   ╚═╝  ╚═╝╚══════╝╚══════╝╚══════╝╚═╝  ╚═╝
                                                                         github.com/ltycn
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
";

            PreventSleep.PreventSleepMode();
            Console.WriteLine(logo);
            Console.WriteLine("Now start Installation, Please do not close this window!");
            Console.WriteLine("_______________________");
            Console.WriteLine("");
            ResourceManager rm = new ResourceManager("BMInstaller.Keys", typeof(Program).Assembly);

            var _3dmKey = rm.GetString("_3DMarkKey");
            var pcmKey = rm.GetString("PCMarkKey");


            List<SoftwareToInstall> softwareList = new List<SoftwareToInstall>
            {
                new SoftwareToInstall("3DMark", @".\3DMark-2.22.7336\3dmark-setup.exe", "/quiet", @"C:\Program Files\UL\3DMark\3DMarkCmd.exe", $"--register={_3dmKey}", runAsAdmin: false),
                new SoftwareToInstall("PCMark 10", @".\PCMark10-v2-1-2506-professional\pcmark10-setup.exe", "/quiet", @"C:\Program Files\UL\PCMark 10\PCMark10Cmd.exe", $"--register={pcmKey}", runAsAdmin: false),
                new SoftwareToInstall("Intel(R)PowerAndThermalAnalysisTool", @".\Intel Power And Thermal Analysis Tool_2.0.0001_Win\Intel(R)PTATWin_2.0.0001.exe", "-S", null, null, runAsAdmin: true), 
                new SoftwareToInstall("7-Zip", @".\7z2301-x64.exe", "/S", null, null, runAsAdmin: true)
            };

            foreach (var software in softwareList)
            {
                if (string.IsNullOrEmpty(software.InstallFilePath) || !File.Exists(software.InstallFilePath))
                {
                    Console.WriteLine($"{software.Name} installation file not found. Skipping installation.");
                    continue;
                }
                if (IsSoftwareInstalled(software.Name))
                {
                    Console.WriteLine($"{software.Name} is already installed.");
                    System.Threading.Thread.Sleep(2000);
                }
                else
                {
                    Console.WriteLine($"Installing {software.Name}...");

                    if (software.RunAsAdmin)
                    {
                        StartProcessAndWaitAsAdmin(software.InstallFilePath, software.InstallArguments);
                    }
                    else
                    {
                        StartProcessAndWait(software.InstallFilePath, software.InstallArguments);
                    }

                    Console.WriteLine($"{software.Name} installation completed.");
                    System.Threading.Thread.Sleep(2000);

                    if (!string.IsNullOrEmpty(software.ActivationFilePath))
                    {
                        Console.WriteLine($"Activating {software.Name}....");
                        StartProcessAndWaitAsAdmin(software.ActivationFilePath, software.ActivationArguments);
                    }
                }
            }

            FinalMove();
            
            // 结束消息
            Console.WriteLine("");
            Console.WriteLine("==///////////////////////   All installations completed   ///////////////////////==");
            Console.WriteLine("");
            PreventSleep.RestoreSleepMode();
            System.Threading.Thread.Sleep(5000);
        }

        static bool IsSoftwareInstalled(string softwareName)
        {
            const string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                if (key != null)
                {
                    foreach (var subKeyName in key.GetSubKeyNames())
                    {
                        using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                        {
                            var displayName = subKey.GetValue("DisplayName") as string;

                            if (!string.IsNullOrEmpty(displayName) && displayName.IndexOf(softwareName, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        public static void FinalMove()
        {
            var sourceDirectory = @"C:\BenchMarkFile";
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            try
            {
                // Move CinebenchR23 directory to the desktop
                var cinebenchDirectory = Path.Combine(sourceDirectory, "CinebenchR23");
                if (Directory.Exists(cinebenchDirectory))
                {
                    Console.WriteLine("Move CinebenchR23 to the desktop...");
                    var destinationCinebenchDirectory = Path.Combine(desktopPath, "CinebenchR23");
                    Directory.Move(cinebenchDirectory, destinationCinebenchDirectory);
                }

                // Move xmltocsv.exe to the desktop
                var xmlToCsvPath = Path.Combine(sourceDirectory, "xmltocsv.exe");
                if (File.Exists(xmlToCsvPath))
                {
                    Console.WriteLine("Move xmltocsv to the desktop...");
                    var destinationXmlToCsvPath = Path.Combine(desktopPath, "xmltocsv.exe");
                    File.Move(xmlToCsvPath, destinationXmlToCsvPath);
                }

                // Delete the entire BenchMarkFile folder
                if (Directory.Exists(sourceDirectory))
                {
                    Console.WriteLine("Cleaning up files...");

                    // Delete operation requires administrator privileges
                    StartProcessAndWaitAsAdmin("cmd.exe", $"/c rmdir /s /q \"{sourceDirectory}\"");

                    Console.WriteLine("Files cleaning completed...");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
        static bool StartProcessAndWait(string filePath, string arguments)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = filePath,
                    Arguments = arguments,
                    CreateNoWindow = true,
                    UseShellExecute = false
                };

                Process process = new Process();
                process.StartInfo = psi;
                process.Start();
                process.WaitForExit();
                return true; // Return true if the process was executed successfully
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return false; // Return false if an error occurred during process execution
            }
        }
        public static void StartProcessAndWaitAsAdmin(string filePath, string arguments)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo(filePath, arguments);
                psi.Verb = "runas"; // This specifies that the process should be started with elevated permissions (run as administrator).
                Process process = new Process();
                process.StartInfo = psi;

                process.Start();
                process.WaitForExit();
            }
            catch (SecurityException)
            {
                Console.WriteLine("Administrator access required, but the user declined the elevation request or is not allowed to run the process as administrator.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while starting the process: {ex.Message}");
            }
        }
    }

    class SoftwareToInstall
    {
        public string Name
        {
            get;
        }
        public string InstallFilePath
        {
            get;
        }
        public string InstallArguments
        {
            get;
        }
        public string ActivationFilePath
        {
            get;
        }
        public string ActivationArguments
        {
            get;
        }
        public bool RunAsAdmin
        {
            get;
        }


        public SoftwareToInstall(string name, string installFilePath, string installArguments, string activationFilePath, string activationArguments, bool runAsAdmin)
        {
            Name = name;
            InstallFilePath = installFilePath;
            InstallArguments = installArguments;
            ActivationFilePath = activationFilePath;
            ActivationArguments = activationArguments;
            RunAsAdmin = runAsAdmin;
            RunAsAdmin = runAsAdmin;
        }
    }

    public class PreventSleep
    {
        // 导入包含SetThreadExecutionState函数的Windows API
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE esFlags);

        // 定义EXECUTION_STATE枚举
        [FlagsAttribute]
        public enum EXECUTION_STATE : uint
        {
            ES_AWAYMODE_REQUIRED = 0x00000040,
            ES_CONTINUOUS = 0x80000000,
            ES_DISPLAY_REQUIRED = 0x00000002,
            ES_SYSTEM_REQUIRED = 0x00000001
            // 可根据需要添加其他标志
        }

        // 阻止计算机进入休眠
        public static void PreventSleepMode()
        {
            SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS |
                                    EXECUTION_STATE.ES_DISPLAY_REQUIRED |
                                    EXECUTION_STATE.ES_SYSTEM_REQUIRED);
        }

        // 恢复计算机休眠状态
        public static void RestoreSleepMode()
        {
            SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS);
        }
    }
}
