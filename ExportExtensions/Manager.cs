﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Text;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using TTRider.ExportExtensions.ExtensionService;

namespace TTRider.ExportExtensions
{
    internal class Manager
    {
        private readonly IVsExtensionManager extensionManager;
        private readonly IVsThreadedWaitDialogFactory dialogFactory;
        private IVsOutputWindowPane generalPane;

        public Manager(IVsExtensionManager extensionManager, IVsThreadedWaitDialogFactory dialogFactory)
        {
            IVsOutputWindow outWindow = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            Guid generalPaneGuid = Microsoft.VisualStudio.VSConstants.GUID_OutWindowDebugPane;
            outWindow.GetPane(ref generalPaneGuid, out generalPane);
            generalPane.Activate(); // Brings this pane into view

            this.extensionManager = extensionManager;
            this.dialogFactory = dialogFactory;
        }

        private void IDEShutDown(EnvDTE.DTE dte)
        {
            if (dte != null)
            {
                // Add code to dispose of custom objects, save files, 
                // and perform any clean-up tasks.

                // Stop external process debugging.
                if (dte.Mode == EnvDTE.vsIDEMode.vsIDEModeDebug)
                {
                    dte.Debugger.Stop(true);
                }

                // Close the DTE object.
                dte.Quit();
            }
        }

        internal void ExportExtensions(object sender, EventArgs e)
        {
            var ofd = new System.Windows.Forms.SaveFileDialog
            {
                AddExtension = true,
                AutoUpgradeEnabled = true,
                CheckPathExists = true,
                DefaultExt = "cmd",
                OverwritePrompt = true,
                FileName = "SetupVisualStudioExtensions.cmd",
                SupportMultiDottedExtensions = true,
                Title = Resources.ExportExtensionList,
                Filter = Resources.ImportExportExtensionFilter
            };

            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            DoExport(ofd.FileName);
        }

        internal void ExportExtensionsTo(object sender, EventArgs e)
        {
            var args = (OleMenuCmdEventArgs)e;
            if (args.InValue != null && args.InValue.ToString() != "")
            {
                var s = args.InValue.ToString();
                if ((s[0] == '\"') && (s[s.Length - 1] == '\"'))
                {
                    s = args.InValue.ToString().Replace("\"", "");
                }
                else if ((s[0] == '\'') && (s[s.Length - 1] == '\''))
                {
                    s = args.InValue.ToString().Replace("\'", "");
                }
                var fileInfo = new FileInfo(s);
                if ((fileInfo.Directory.ToString() != (new FileInfo("Invalid").Directory.ToString())) &&
                    (fileInfo.Extension == ".cmd"))
                {
                    DoExport(s);
                }

                return;
            } else
            {
                generalPane.OutputString("No path specified for export.\n");
                generalPane.Activate();
            }
        }

        internal void ExportExtensionsToAndExit(object sender, EventArgs e)
        {
            var args = (OleMenuCmdEventArgs)e;
            if (args.InValue != null && args.InValue.ToString() != "")
            {
                var s = args.InValue.ToString();
                if ((s[0] == '\"') && (s[s.Length - 1] == '\"'))
                {
                    s = args.InValue.ToString().Replace("\"", "");
                }
                else if ((s[0] == '\'') && (s[s.Length - 1] == '\''))
                {
                    s = args.InValue.ToString().Replace("\'", "");
                }
                var fileInfo = new FileInfo(s);
                if ((fileInfo.Directory.ToString() != (new FileInfo("Invalid").Directory.ToString())) &&
                    (fileInfo.Extension == ".cmd"))
                {
                    DoExport(s);
                    IDEShutDown(Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(Microsoft.VisualStu‌​dio.Shell.Interop.SD‌​TE)) as EnvDTE.DTE);
                }

                return;
            }
            else
            {
                generalPane.OutputString("No path specified for export.\n");
                generalPane.Activate();
            }
        }

        private IEnumerable<ExtensionInfo> GetInstalledExtensions()
        {
            return this.extensionManager.GetInstalledExtensions().
                Where(ext => !ext.Header.SystemComponent)
                .Select(ext => new ExtensionInfo
                {
                    Name = ext.Header.Name,
                    Description = ext.Header.Description,
                    Author = ext.Header.Author,
                    Identifier = ext.Header.Identifier,
                    State = ext.State
                });
        }

        string CmdEncode(string str, bool extended=true)
        {
            var sb = new StringBuilder(str);
            if (extended) 
            {
                sb.Replace("\r", "");
                sb.Replace("\n", "");
                sb.Replace("'", "''");
            }
            sb.Replace(">", "^>");
            sb.Replace("<", "^<");
            sb.Replace("%", "^%");
            sb.Replace("&", "^&");
            return sb.ToString();
        }

        Release DownloadExtensionDetails(ExtensionInfo ex)
        {
            var endpointAddress = new EndpointAddress("https://visualstudiogallery.msdn.microsoft.com/Services/dev12/Extension.svc");
            var binding = new WSHttpBinding(SecurityMode.Transport)
            {
                MessageEncoding = WSMessageEncoding.Text,
                TextEncoding = Encoding.UTF8
            };
            var extensionService = new VsIdeServiceClient(binding, endpointAddress);

            var entry = extensionService.SearchReleases("",
                string.Format("(Project.Metadata['VsixId'] = '{0}')", ex.Identifier),
                "Project.Metadata['Relevance'] desc", null, 0, 10);
            return entry.Releases.LastOrDefault();
        }

        ExtensionInfo GetExtensionDownloadUrl(ExtensionInfo ex)
        {
            var release = DownloadExtensionDetails(ex);
            //generalPane.OutputString(ex.Name + "\t" +
            //    ((release == null) ? "null" : "not null") + "\t" +
            //    //ex.Description + "\t" +
            //    ex.Author + "\t" +
            //    ex.Identifier + "\t" +
            //    ex.State.ToString() + "\n");
            string url = null;
            if (release != null && !release.Project.Metadata.TryGetValue("DownloadUrl", out url))
            {
                return ex;
            }

            Uri uri;
            if (!Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out uri))
            {
                return ex;
            }

            ex.DownloadUrl = uri.AbsoluteUri;
            ex.DownloadAs = Path.GetFileName(uri.LocalPath);
            return ex;
        }

        TextReader GetScriptPart(string name)
        {
            var stream =
                Assembly.GetExecutingAssembly().GetManifestResourceStream("TTRider.ExportExtensions.templates." + name);
            if (stream == null) throw new ApplicationException();
            return new StreamReader(stream);
        }

        private void DoExport(string fileName)
        {
            IVsThreadedWaitDialog2 dialog;
            dialogFactory.CreateInstance(out dialog);

            dialog.StartWaitDialog(
                Resources.ProgressTitle, Resources.ProgressHeader,
                Resources.LoadingList, null,
                Resources.ProgressTitle,
                0, false,
                true);

            try
            {
                //var extensions = this.GetInstalledExtensions();


                //var count = 0;
                //generalPane.OutputString("Short list:\n");
                //foreach (var extension in extensions)
                //{
                //    count++;
                //    generalPane.OutputString(extension.Name + "\t" +
                //        //extension.Description + "\t" +
                //        extension.Author + "\t" +
                //        extension.Identifier + "\t" +
                //        extension.State.ToString() + "\n");
                //}
                //generalPane.OutputString("Total: " + count + "\n\n");
                //generalPane.Activate(); // Brings this pane into view

                var extensions = this.GetInstalledExtensions()
					.Select(ext =>
                    {
                        if (dialog != null)
                        {
                            bool canceled;
                            dialog.UpdateProgress(
                                Resources.ProgressHeader,
                                string.Format(Resources.Loading, ext.Name),
                                string.Format(Resources.Loading, ext.Name), 0, 0, false, out canceled);
                        }
                        return GetExtensionDownloadUrl(ext);
                    }).Where(ext => !string.IsNullOrWhiteSpace(ext.DownloadAs));


                using (var writer = File.CreateText(fileName))
                {
                    writer.WriteLine(GetScriptPart("header.template").ReadToEnd());

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

                    var ps = GetScriptPart("ps.template");
                    var line = ps.ReadLine();
                    while (line != null)
                    {
                        writer.WriteLine("echo " + CmdEncode(line, false) + " >> ext.ps1");
                        line = ps.ReadLine();
                    }

                    writer.WriteLine(GetScriptPart("footer.template").ReadToEnd());
                }
            }
            finally
            {
                int usercancel;
                dialog.EndWaitDialog(out usercancel);
            }
        }
    }
}
