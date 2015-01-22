using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using TTRider.VSExtensionsImportExport.ExtensionService;

namespace TTRider.VSExtensionsImportExport
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
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                mcs.AddCommand(new MenuCommand(OnExportExtensions, new CommandID(GuidList.guidVSExtensionsImportExportCmdSet, (int)PkgCmdIDList.cmdidExportExtensionList)));
                mcs.AddCommand(new MenuCommand(OnImportExtensions, new CommandID(GuidList.guidVSExtensionsImportExportCmdSet, (int)PkgCmdIDList.cmdidImportExtensionList)));
            }
        }
        #endregion


        private void OnImportExtensions(object sender, EventArgs e)
        {
            var pane = GetOutputWindow();

            var ofd = new System.Windows.Forms.OpenFileDialog
            {
                AddExtension = true,
                AutoUpgradeEnabled = true,
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xml",
                FileName = "VisualStudioExtensions.vsixlist.xml",
                Multiselect = false,
                SupportMultiDottedExtensions = true,
                Title = Resources.ImportExtensionList,
                Filter = Resources.ImportExportExtensionFilter
            };

            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            ImportExtensions(ofd.FileName);
        }


        private void OnExportExtensions(object sender, EventArgs e)
        {
            var pane = GetOutputWindow();

            var ofd = new System.Windows.Forms.SaveFileDialog
            {
                AddExtension = true,
                AutoUpgradeEnabled = true,
                CheckPathExists = true,
                DefaultExt = "xml",
                OverwritePrompt = true,
                FileName = "VisualStudioExtensions.vsixlist.xml",
                SupportMultiDottedExtensions = true,
                Title = Resources.ExportExtensionList,
                Filter = Resources.ImportExportExtensionFilter
            };

            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            ExportExtensions(ofd.FileName);

        }

        private void ExportExtensions(string fileName)
        {
            var output = GetOutputWindow();
            try

            {
                var exs = new ExtensionSet();


                var vsextm = (IVsExtensionManager) GetService(typeof (SVsExtensionManager));
                if (vsextm != null)
                {
                    exs.MachineName = Environment.MachineName;
                    exs.Timestamp = DateTime.UtcNow;
                    exs.Extensions.AddRange(vsextm.GetInstalledExtensions().
                        Where(ext => !ext.Header.SystemComponent)
                        .Select(ext => new ExtensionInfo
                        {
                            Name = ext.Header.Name,
                            LocalizedName = ext.Header.LocalizedName,
                            Description = ext.Header.Description,
                            LocalizedDescription = ext.Header.LocalizedDescription,
                            Author = ext.Header.Author,
                            Identifier = ext.Header.Identifier
                        })
                        .Select(ext =>
                        {
                            output.OutputStringThreadSafe(string.Format("{0} by {1}\r\n", ext.LocalizedName, ext.Author));
                            return ext;
                        }));
                }

                ExtensionSetFactory.Write(fileName, exs);

                output.OutputStringThreadSafe("Saved list to ");
                output.OutputStringThreadSafe(fileName);
            }
            catch (Exception ex)
            {
                output.OutputStringThreadSafe("ERROR: "+ex.Message);
            }
            output.FlushToTaskList();
        }

        private void ImportExtensions(string fileName)
        {
            var output = GetOutputWindow();
            try

            {
                output.OutputStringThreadSafe("Loading list from ");
                output.OutputStringThreadSafe(fileName);
                var exs = ExtensionSetFactory.Read(fileName);

                output.OutputStringThreadSafe(string.Format("\r\nList Time Stamp (UTC): {0}\r\n", exs.Timestamp));
                output.OutputStringThreadSafe(string.Format("List Machine Name: {0}\r\n", exs.MachineName));
            
                // 
                var vsextm = (IVsExtensionManager)GetService(typeof(SVsExtensionManager));
                if (vsextm != null)
                {
                    var installed = vsextm.GetInstalledExtensions().
                        Where(ext => !ext.Header.SystemComponent)
                        .Select(ext => new ExtensionInfo{Identifier = ext.Header.Identifier});

                    var missing = exs.Extensions.Except(installed, ExtansionInfoEqualityComparer.Default);

                    output.OutputStringThreadSafe("Processing missing extentions\r\n");


                    var endpointAddress = new EndpointAddress("https://visualstudiogallery.msdn.microsoft.com/Services/dev12/Extension.svc");
                    var binding = new WSHttpBinding(SecurityMode.Transport);
                    binding.MessageEncoding = WSMessageEncoding.Text;
                    binding.TextEncoding = Encoding.UTF8;
                    

                    var extensionService = new VsIdeServiceClient(binding, endpointAddress);


                    var requests =
                        missing.Select(ex =>
                        {
                            output.OutputStringThreadSafe("Looking up for "+ex.LocalizedName);
                            return extensionService.SearchReleasesAsync("",
                                string.Format("(Project.Metadata['VsixId'] = '{0}')", ex.Identifier),
                                "", null, 0, 10).ContinueWith(t =>
                                {
                                    if (t.IsFaulted)
                                    {
                                        if (t.Exception!=null)
                                        {
                                            foreach (var exc in t.Exception.InnerExceptions)
                                        {
                                            output.OutputStringThreadSafe(string.Format("ERROR: ({0}): {1}\r\n", ex.LocalizedName, (exc != null) ? exc.Message : "Unknown"));
                                        }}
                                        else
                                        {
                                            output.OutputStringThreadSafe(string.Format("ERROR: ({0}): unknown\r\n", ex.LocalizedName));
                                        }
                                        return;
                                    }

                                    

                                });
                        });

                    System.Threading.Tasks.Task.WaitAll(requests.ToArray());

                }
            
            }
            catch (Exception ex)
            {
                output.OutputStringThreadSafe("ERROR: " + ex.Message);
            }
            output.FlushToTaskList();
        }

        /*
         
         var vsextm = (IVsExtensionManager)GetService(typeof(SVsExtensionManager));

            //SVsExtensionRepository

            


            if (vsextm != null)
            {
                var ie = vsextm.GetInstalledExtensions();
                foreach (var ext in ie)
                {
                    var sb = new StringBuilder("=======================================================");
                    sb.AppendFormat("Author:\t\t{0}\r\n", ext.Header.Author);
                    sb.AppendFormat("Description:\t\t{0}\r\n", ext.Header.Description);
                    sb.AppendFormat("GettingStartedGuide:\t\t{0}\r\n", ext.Header.GettingStartedGuide);
                    sb.AppendFormat("GlobalScope:\t\t{0}\r\n", ext.Header.GlobalScope);
                    sb.AppendFormat("Identifier:\t\t{0}\r\n", ext.Header.Identifier);
                    sb.AppendFormat("LocalizedDescription:\t\t{0}\r\n", ext.Header.LocalizedDescription);
                    sb.AppendFormat("LocalizedName:\t\t{0}\r\n", ext.Header.LocalizedName);
                    sb.AppendFormat("MoreInfoUrl:\t\t{0}\r\n", ext.Header.MoreInfoUrl);
                    sb.AppendFormat("Name:\t\t{0}\r\n", ext.Header.Name);
                    sb.AppendFormat("SystemComponent:\t\t{0}\r\n", ext.Header.SystemComponent);
                    sb.AppendFormat("Tags:\t\t{0}\r\n", ext.Header.Tags);
                    sb.AppendFormat("Version:\t\t{0}\r\n", ext.Header.Version);
                    sb.AppendFormat("Type:\t\t{0}\r\n", ext.Type);
                    sb.AppendFormat("State:\t\t{0}\r\n", ext.State);

                    foreach (var itemc in ext.Content)
                    {
                        sb.AppendLine("-----");
                        sb.AppendFormat("Content:ContentTypeName:\t\t{0}\r\n", itemc.ContentTypeName);
                        sb.AppendFormat("Content:RelativePath:\t\t{0}\r\n", itemc.RelativePath);
                        foreach (var attr in itemc.Attributes)
                        {
                            sb.AppendFormat("Content:Attributes:\t\t{0}={1}\r\n", attr.Key, attr.Value);
                        }
                    }
                    //ret.OutputString(sb.ToString());
                    //ret.FlushToTaskList();
                }
            }

            var vsexr = (IVsExtensionRepository)GetService(typeof(SVsExtensionRepository));
            if (vsexr != null)
            {
                var qq = vsexr.CreateQuery<RepoEntry>().Where(re=>re.VsixID == "f4ab1e64-5d35-4f06-bad9-bf414f4b3bbb")
                    .OrderByDescending(v => v.Name).Take(25);

                var q = (IVsExtensionRepositoryQuery<RepoEntry>)qq;

                q.ExecuteCompleted += (s, ee) => {

                    var eeee = ee;

                };

                q.ExecuteAsync();

                //var query = vsexr.CreateQuery<VSGalleryEntry>(false, true)
                //               .OrderByDescending(v => v.Ranking)
                //               .Skip(0)
                //               .Take(25) as IVsExtensionRepositoryQuery<VSGalleryEntry>;

                //query.ExecuteCompleted += (s, ee) => {

                //};
                //query.SearchText = "f4ab1e64-5d35-4f06-bad9-bf414f4b3bbb";
                //query.ExecuteAsync();

            }


            // Show a Message Box to prove we were here
            //IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            //Guid clsid = Guid.Empty;
            //int result;
            //Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
            //           0,
            //           ref clsid,
            //           "Extension Sync",
            //           string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.ToString()),
            //           string.Empty,
            //           0,
            //           OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //           OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
            //           OLEMSGICON.OLEMSGICON_INFO,
            //           0,        // false
            //           out result));
         
         */

        IVsOutputWindowPane GetOutputWindow()
        {
            var localOutputPane = outputPane;

            IVsOutputWindowPane ret = null;
            var wnd = (IVsOutputWindow)GetService(typeof(IVsOutputWindow));
            if (wnd != null)
            {
                wnd.GetPane(ref localOutputPane, out ret);

                if (ret == null)
                {
                    wnd.CreatePane(ref localOutputPane, Resources.OutputPaneName, 0, 1);
                    wnd.GetPane(ref localOutputPane, out ret);
                }
                ret.Clear();
                ret.Activate();
            }
            return ret;
        }


        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Show a Message Box to prove we were here
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            int result;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                       0,
                       ref clsid,
                       "VSExtensionsImportExport",
                       string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.ToString()),
                       string.Empty,
                       0,
                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO,
                       0,        // false
                       out result));
        }

    }
}
