using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Linq;

namespace Ptlk_OPC
{
    [RunInstaller(true)]
    public partial class RegisterDll : Installer
    {
        // Get the locaton of systemX86
        private static string systemX86Path = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86);
        // Get the location of regsvr32
        private static string regsvr32Path = Path.Combine(systemX86Path, "regsvr32.exe");
        // Get the location of regasm
        private static string regasmPath = Path.Combine(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "regasm.exe");
        // Get the location of our DLL
        private static string componentPath = typeof(RegisterDll).Assembly.Location;

        public RegisterDll()
        {
            InitializeComponent();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            // Register dependency DLL
            System.Diagnostics.Process.Start(regsvr32Path, Path.Combine(systemX86Path, "/s opccomn_ps.dll")).WaitForExit();
            System.Diagnostics.Process.Start(regsvr32Path, Path.Combine(systemX86Path, "/s opcproxy.dll")).WaitForExit();
            System.Diagnostics.Process.Start(regsvr32Path, Path.Combine(systemX86Path, "/s OPCDAAuto.dll")).WaitForExit();

            // Register our DLL
            System.Diagnostics.Process.Start(regasmPath, $"\"{componentPath}\" /tlb /codebase").WaitForExit();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Uninstall(IDictionary savedState)
        {
            // Unregister our DLL
            System.Diagnostics.Process.Start(regasmPath, $"\"{componentPath}\" /tlb /u").WaitForExit();

            // Delete type library
            FileInfo tlbfile = new FileInfo(componentPath.Replace(".dll", ".tlb"));
            if (tlbfile.Exists) tlbfile.Delete();

            base.Uninstall(savedState);
        }
    }
}
