using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

//using System.Net.Http;

namespace TTRider.ExportExtensions
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
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidVSExtensionsImportExportPkgString)]
    public sealed class VSExtensionsImportExportPackage : Package
    {
        static readonly Guid outputPane = new Guid("{FCC482B4-E7CC-4120-B9D5-04A45CB90A68}");
        IVsExtensionManager vsextm;

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public VSExtensionsImportExportPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                mcs.AddCommand(new MenuCommand(OnExportExtensions, new CommandID(GuidList.guidVSExtensionsImportExportCmdSet, (int)PkgCmdIDList.cmdidExportExtensionList)));
            }

            this.vsextm = (IVsExtensionManager)GetService(typeof(SVsExtensionManager));
        }
        #endregion


        private void OnExportExtensions(object sender, EventArgs e)
        {
            var ofd = new System.Windows.Forms.SaveFileDialog
            {
                AddExtension = true,
                AutoUpgradeEnabled = true,
                CheckPathExists = true,
                DefaultExt = "xml",
                OverwritePrompt = true,
                FileName = "VisualStudioExtensions.cmd",
                SupportMultiDottedExtensions = true,
                Title = Resources.ExportExtensionList,
                Filter = Resources.ImportExportExtensionFilter
            };

            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            ExportExtensionsToVS(ofd.FileName);
        }

        private void ExportExtensionsToVS(string fileName)
        {
            var dialogFactory = GetService(typeof(SVsThreadedWaitDialogFactory)) as IVsThreadedWaitDialogFactory;
            IVsThreadedWaitDialog2 dialog = null;
            if (dialogFactory != null)
            {
                dialogFactory.CreateInstance(out dialog);
            }

            var vsextm = (IVsExtensionManager)GetService(typeof(SVsExtensionManager));
            if (vsextm != null)
            {
                if (dialog != null)
                {
                    dialog.StartWaitDialog(
                        Resources.ProgressTitle, Resources.ProgressHeader,
                        Resources.LoadingList, null,
                        Resources.ProgressTitle,
                        0, false,
                        true);
                }

                try
                {
                    var extensions = ExtensionSetFactory.GetInstalledExtensions(vsextm)
                            .Select(ext =>
                            {
                                if (dialog != null)
                                {
                                    bool canceled;
                                    dialog.UpdateProgress(Resources.ProgressHeader, string.Format(Resources.Loading, ext.Name), string.Format(Resources.Loading, ext.Name), 0, 0, false, out canceled);
                                }
                                return ExtensionSetFactory.GetExtensionDownloadUrl(ext);
                            }).Where(ext => !string.IsNullOrWhiteSpace(ext.DownloadAs));


                    using (var writer = File.CreateText(fileName))
                    {
                        writer.WriteLine(ExtensionSetFactory.GetScriptPart("header.template").ReadToEnd());

                        var index = 0;
                        foreach (var extension in extensions)
                        {
                            bool isCancelled;
                            dialog.HasCanceled(out isCancelled);
                            if (isCancelled)
                            {
                                return;
                            }

                            var cmd = string.Format("echo $ext{0} = @{{Name = '{1}';Description ='{2}';Author ='{3}';Identifier ='{4}';DownloadUrl='{5}';DownloadAs='{6}'}}; >> ext.ps1",
                                index++,
                                CmdEncode(extension.Name),
                                CmdEncode(extension.Description),
                                CmdEncode(extension.Author),
                                CmdEncode(extension.Identifier),
                                CmdEncode(extension.DownloadUrl),
                                CmdEncode(extension.DownloadAs));
                            writer.WriteLine(cmd);
                        }

                        if (index > 0)
                        {
                            var sep = "echo $ext = (";
                            for (int i = 0; i < index; i++)
                            {
                                writer.Write(sep);
                                writer.Write("$ext{0}", i);
                                sep = ",";
                            }
                            writer.WriteLine(");  >> ext.ps1");
                        }

                        var ps = ExtensionSetFactory.GetScriptPart("ps.template");
                        var line = ps.ReadLine();
                        while (line != null)
                        {
                            writer.WriteLine("echo "+CmdEncode(line)+" >> ext.ps1");
                            line = ps.ReadLine();
                        }

                        writer.WriteLine(ExtensionSetFactory.GetScriptPart("footer.template").ReadToEnd());

                    }
                }
                finally
                {
                    if (dialog != null)
                    {
                        int usercancel;
                        dialog.EndWaitDialog(out usercancel);
                    }
                }
            }
        }

        public static string CmdEncode(string str)
        {
            var sb = new StringBuilder(str);
            sb.Replace(">", "^>");
            sb.Replace("<", "^<");
            sb.Replace("%", "^%");
            sb.Replace("&", "^&");
            return sb.ToString();
        }
    }
}
