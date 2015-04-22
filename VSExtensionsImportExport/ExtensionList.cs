using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Text;
using System.Xml.Serialization;
using Microsoft.VisualStudio.ExtensionManager;
using TTRider.ExportExtensions.Service_References.ExtensionService;

namespace TTRider.ExportExtensions
{
    public class ExtensionInfo
    {
        [XmlElement]
        public string Name { get; set; }
        public string Description { get; set; }
        public string Author { get; set; }
        [XmlElement]
        public string Identifier { get; set; }
        [XmlElement]
        public string DownloadUrl { get; set; }
        [XmlElement]
        public string DownloadAs { get; set; }

    }

    public static class ExtensionSetFactory
    {

        public static TextReader GetScriptPart(string name)
        {
            return new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("TTRider.VSExtensionsImportExport.templates."+name));
        }

        public static IEnumerable<ExtensionInfo> GetInstalledExtensions(IVsExtensionManager exm)
        {
            return exm.GetInstalledExtensions().
                Where(ext => !ext.Header.SystemComponent)
                .Select(ext => new ExtensionInfo
                {
                    Name = ext.Header.Name,
                    Description = ext.Header.Description,
                    Author = ext.Header.Author,
                    Identifier = ext.Header.Identifier
                });
        }


        public static Release DownloadExtensionDetails(ExtensionInfo ex)
        {
            var endpointAddress = new EndpointAddress("https://visualstudiogallery.msdn.microsoft.com/Services/dev12/Extension.svc");
            var binding = new WSHttpBinding(SecurityMode.Transport);
            binding.MessageEncoding = WSMessageEncoding.Text;
            binding.TextEncoding = Encoding.UTF8;
            var extensionService = new VsIdeServiceClient(binding, endpointAddress);

            var entry = extensionService.SearchReleases("",
                string.Format("(Project.Metadata['VsixId'] = '{0}')", ex.Identifier),
                "Project.Metadata['Relevance'] desc", null, 0, 10);
            return entry.Releases.LastOrDefault();
        }

        public static ExtensionInfo GetExtensionDownloadUrl(ExtensionInfo ex)
        {
            var release = DownloadExtensionDetails(ex);
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
    }
}
