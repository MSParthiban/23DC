
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;


namespace _23DC
{
    public partial class acadIntrope : Window
    {


        public acadIntrope()
        {
            bool isAutoCADRunning = Utilities.IsAutoCADRunning();
            if (isAutoCADRunning == false)
            {
                MessageBox.Show("Starting new AutoCAD 2022 instance...this will take a some time.");
                Utilities.StartAutoCADApp();
            }
            else
            {
                MessageBox.Show("AutoCAD already running.");
            }
            Utilities.SendMessage("AutoCAD started from WPF");
            MyDriver.CreateMyProfile();
            Utilities.SendMessage("Profile created:" + Utilities.yourProfileName);
            //MyDriver.NetloadMyApp(@"C:\YourPathFolderPath\custom.Dll");
        }
    }



    public class MyDriver
    {
        public static void CreateMyProfile()
        {
            bool isAutoCADRunning = Utilities.IsAutoCADRunning();
            if (isAutoCADRunning == false)
                Utilities.StartAutoCADApp();
            Utilities.CreateProfile();
        }

        public static void NetloadMyApp(String dllPath)
        {
            bool isAutoCADRunning = Utilities.IsAutoCADRunning();
            if (isAutoCADRunning == false)
                Utilities.StartAutoCADApp();
            Utilities.NetloadDll(dllPath);
        }
    }

    public class Utilities
    {
        [System.Runtime.InteropServices.DllImport("user32")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hwnd, ref IntPtr lpdwProcessId);

        private static readonly string AutoCADProgId = "AutoCAD.Application.24.1";
        private static AcadApplication App;


        public static void SendMessage(String message)
        {
            App.ActiveDocument.SendCommand("(princ \"" + message + "\")(princ)" + Environment.NewLine);
        }


        public static bool IsAutoCADRunning()
        {
            bool isRunning = GetRunningAutoCADInstance();
            return isRunning;
        }

        public static bool ConfigureRunningAutoCADForUsage()
        {
            if (App == null)
                return false;
            MessageFilter.Register();
            SetAutoCADWindowToNormal();
            return true;
        }

        public static bool StartAutoCADApp()
        {
            Type autocadType = System.Type.GetTypeFromCLSID(new Guid("AA46BA8A-9825-40FD-8493-0BA3C4D5CEB5"), true);
            object obj = System.Activator.CreateInstance(autocadType, true);
            AcadApplication appAcad = (AcadApplication)obj;
            App = appAcad;
            MessageFilter.Register();
            SetAutoCADWindowToNormal();
            return true;
        }

        public static bool NetloadDll(string dllPath)
        {
            if (!System.IO.File.Exists(dllPath))
                throw new Exception("Dll does not exist: " + dllPath);
            App.ActiveDocument.SendCommand("(setvar \"secureload\" 0)" + Environment.NewLine);
            dllPath = dllPath.Replace(@"\", @"\\");
            App.ActiveDocument.SendCommand("(command \"_netload\" \"" + dllPath + "\")" + Environment.NewLine);
            return true;
        }


        public static bool CreateProfile()
        {
            if (App == null)
                return false;
            bool profileExists = DoesProfileExist(App, yourProfileName);
            if (profileExists)
            {
                SetYourProfileActive(App, yourProfileName);
                AddTempFolderToTrustedPaths(App);
            }
            else
            {
                CreateYourCustomProfile(App, yourProfileName);
                AddTempFolderToTrustedPaths(App);
            }
            SetYourProfileActive(App, yourProfileName);
            return true;
        }


        public static bool SetAutoCADWindowToNormal()
        {
            if (App == null)
                return false;
            App.WindowState = AcWindowState.acNorm;
            return true;
        }






        private static bool GetRunningAutoCADInstance()
        {
            Type autocadType = System.Type.GetTypeFromProgID(AutoCADProgId, true);
            AcadApplication appAcad;
            try
            {
                object obj = Microsoft.VisualBasic.Interaction.GetObject(null, AutoCADProgId);
                appAcad = (AcadApplication)obj;
                App = appAcad;
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return false;
        }

        public static readonly string yourProfileName = "myCustomProfile";

        private static void SetYourProfileActive(AcadApplication appAcad, string profileName)
        {
            AcadPreferencesProfiles profiles = appAcad.Preferences.Profiles;
            profiles.ActiveProfile = profileName;
        }

        private static void CreateYourCustomProfile(AcadApplication appAcad, string profileName)
        {
            AcadPreferencesProfiles profiles = appAcad.Preferences.Profiles;
            profiles.CopyProfile(profiles.ActiveProfile, profileName);
            profiles.ActiveProfile = profileName;
        }

        private static bool DoesProfileExist(AcadApplication appAcad, string profileName)
        {
            AcadPreferencesProfiles profiles = appAcad.Preferences.Profiles;
            object pNames = null;
            profiles.GetAllProfileNames(out pNames);
            string[] profileNames = (string[])pNames;
            foreach (string name in profileNames)
            {
                if (name.Equals(profileName))
                    return true;
            }
            return false;
        }

        private static void AddTempFolderToTrustedPaths(AcadApplication appAcad)
        {
            string trustedPathsString = System.Convert.ToString(appAcad.ActiveDocument.GetVariable("TRUSTEDPATHS"));
            string tempDirectory = System.IO.Path.GetTempPath();
            List<string> newPaths = new List<string>() { tempDirectory };
            if (!trustedPathsString.Contains(tempDirectory))
                AddTrustedPaths(appAcad, newPaths);
        }

        private static void AddTrustedPaths(AcadApplication appAcad, List<string> newPaths)
        {
            string trustedPathsString = System.Convert.ToString(appAcad.ActiveDocument.GetVariable("TRUSTEDPATHS"));
            List<string> oldPaths = new List<string>();
            oldPaths = trustedPathsString.Split(System.Convert.ToChar(";")).ToList();
            string newTrustedPathsString = trustedPathsString;
            foreach (string newPath in newPaths)
            {
                bool pathAlreadyExists = trustedPathsString.Contains(newPath);
                if (!pathAlreadyExists)
                    newTrustedPathsString = newPath + ";" + newTrustedPathsString;
            }
            appAcad.ActiveDocument.SetVariable("TRUSTEDPATHS", newTrustedPathsString);
        }
    }


    public class MessageFilter : IOleMessageFilter
    {
        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, ref IOleMessageFilter oldFilter);

        public static void Register()
        {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(newFilter, ref oldFilter);
        }
        public static void Revoke()
        {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, ref oldFilter);
        }

        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return 0;
        }

        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == 2)
                // flag = SERVERCALL_RETRYLATER.

                // Retry the thread call immediately if return >=0 & 
                // <100.
                return 99;
            // Too busy; cancel call.
            return -1;
        }

        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return 2;
        }
    }

    [ComImport()]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

}
