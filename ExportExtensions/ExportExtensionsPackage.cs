using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;


namespace TTRider.ExportExtensions
{

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidVSExtensionsImportExportPkgString)]
    public sealed class ExportExtensionsPackage : Package
    {
        private Manager manager;

        #region Package Members

        protected override void Initialize()
        {
            base.Initialize();

            var extensionManager = (IVsExtensionManager)GetService(typeof(SVsExtensionManager));
            var dialogFactory = (IVsThreadedWaitDialogFactory)GetService(typeof(SVsThreadedWaitDialogFactory));

            this.manager = new Manager(extensionManager, dialogFactory);

            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                mcs.AddCommand(new MenuCommand(this.manager.ExportExtensions, new CommandID(GuidList.guidVSExtensionsImportExportCmdSet, (int)PkgCmdIDList.cmdidExportExtensionList)));
                
                OleMenuCommand menuItem2 = new OleMenuCommand(this.manager.ExportExtensionsTo, new CommandID(GuidList.guidVSExtensionsImportExportCmdSet, (int)PkgCmdIDList.cmdidExportExtensionListTo));
                menuItem2.ParametersDescription = "p";
                mcs.AddCommand(menuItem2);

                OleMenuCommand menuItem3 = new OleMenuCommand(this.manager.ExportExtensionsToAndExit, new CommandID(GuidList.guidVSExtensionsImportExportCmdSet, (int)PkgCmdIDList.cmdidExportExtensionListToAndExit));
                menuItem3.ParametersDescription = "p";
                mcs.AddCommand(menuItem3);
            }
        }
        #endregion
    }
}
