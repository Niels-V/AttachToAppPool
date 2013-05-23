using System;
using System.Linq;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using System.Collections.Generic;

namespace NielsV.NielsV_AttachToAppPool
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidNielsV_AttachToAppPoolPkgString)]
    public sealed class NielsV_AttachToAppPoolPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public NielsV_AttachToAppPoolPackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        private int baseMRUID = (int)PkgCmdIDList.cmdidNielsVAppPoolList;
        private List<string> appPoolList;
        /////////////////////////////////////////////////////////////////////////////
        // Overriden Package Implementation

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initilaization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Trace.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                this.AddCommand(mcs, PkgCmdIDList.cmdidNielsVRefreshAppPoolList);
                this.InitAppPoolListMenu(mcs);
            }
        }
        private void AddCommand(OleMenuCommandService mcs, uint commandId)
        {
            OleMenuCommand menuItemCommand = new OleMenuCommand(
                delegate(object sender, EventArgs e)
                {
                    RefreshAppPoolMenu(mcs);
                },
                new CommandID(GuidList.guidNielsV_AttachToAppPoolCmdSet, (int)commandId));
            //menuItemCommand.BeforeQueryStatus += (s, e) => menuItemCommand.Visible = isVisible((GeneralOptionsPage)this.GetDialogPage(typeof(GeneralOptionsPage)));
            mcs.AddCommand(menuItemCommand);
        }

        private void RefreshAppPoolMenu(OleMenuCommandService mcs)
        {
            
            var appPools = LookupAppPools.GetAppPoolProcesses();
            int j = 0;
            foreach (var appPool in appPoolList)
            {
                var cmdID = new CommandID(GuidList.guidNielsV_AttachToAppPoolCmdSet, this.baseMRUID + j);
                var menuCommand = mcs.FindCommand(cmdID) as OleMenuCommand;
                if (j >= 0 && j < this.appPoolList.Count && j < appPools.Count)
                {
                    menuCommand.Text = appPools.Values.ElementAt(j);
                    menuCommand.Visible = true;
                }
                else
                {
                    menuCommand.Visible = false;
                }
                j++;
            }
            
        }


        private void InitAppPoolListMenu(OleMenuCommandService mcs)
        {
            if (null == this.appPoolList)
            {
                this.appPoolList = new List<string>();
                if (null != this.appPoolList)
                {
                    for (int i = 0; i < 10; i++) //support max 10 appPools
                    {
                        appPoolList.Add("AppPool" + i);
                    }
                }
            }
            int j = 0;
            foreach (var appPool in appPoolList)
            {
                var cmdID = new CommandID(
                    GuidList.guidNielsV_AttachToAppPoolCmdSet, this.baseMRUID + j);
                var mc = new OleMenuCommand(
                    new EventHandler(OnAppPoolExec), cmdID);
                mc.BeforeQueryStatus += new EventHandler(OnAppPoolQueryStatus);
                mcs.AddCommand(mc);
                j++;
            }
        }

        private void AttachToProcess(int processId)
        {
            DTE dte = (DTE)this.GetService(typeof(DTE));

            foreach (EnvDTE.Process process in dte.Debugger.LocalProcesses)
            {
                if (process.ProcessID == processId)
                {
                    process.Attach();
                    break;
                }
            }
        }

        private void OnAppPoolQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (null != menuCommand)
            {
                var appPools = LookupAppPools.GetAppPoolProcesses();
                int appPoolItemIndex = menuCommand.CommandID.ID - this.baseMRUID;
                if (appPoolItemIndex >= 0 && appPoolItemIndex < this.appPoolList.Count && appPoolItemIndex < appPools.Count)
                {
                    menuCommand.Text = appPools.Values.Skip(appPoolItemIndex).FirstOrDefault();
                    menuCommand.Visible = true;
                }
                else
                {
                    menuCommand.Visible = false;
                }
            }
        }

        private void OnAppPoolExec(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;
            if (null != menuCommand)
            {
                var appPools = LookupAppPools.GetAppPoolProcesses();
                int appPoolItemIndex = menuCommand.CommandID.ID - this.baseMRUID;
                if (appPoolItemIndex >= 0 && appPoolItemIndex < this.appPoolList.Count && appPoolItemIndex < appPools.Count)
                {
                    var appPoolProcId = appPools.Where(ap => ap.Value == menuCommand.Text).FirstOrDefault().Key;
                    AttachToProcess(appPoolProcId);
                }
            }
        }
    }
}
