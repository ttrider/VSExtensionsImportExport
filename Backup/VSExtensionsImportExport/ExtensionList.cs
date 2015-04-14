using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Microsoft.Internal.VisualStudio.PlatformUI;

namespace TTRider.VSExtensionsImportExport
{
    [XmlRoot("ExtensionSet", Namespace = "http://schemas.ttrider.com/schemas/visualstudioextensionlist.xsd")]
    public class ExtensionSet
    {
        public ExtensionSet()
        {
            this.Extensions = new List<ExtensionInfo>();
        }

        [XmlArray]
        [XmlArrayItem("Extension")]
        public List<ExtensionInfo> Extensions { get; private set; }

        [XmlElement]
        public string MachineName { get; set; }
        [XmlElement]
        public DateTimeOffset Timestamp { get; set; }
    }

    public class ExtensionInfo
    {
        [XmlElement]
        public string Name { get; set; }
        [XmlElement]
        public string LocalizedName { get; set; }
        [XmlElement]
        public string Description { get; set; }
        [XmlElement]
        public string LocalizedDescription { get; set; }
        [XmlElement]
        public string Author { get; set; }
        [XmlElement]
        public string Identifier { get; set; }
    }

    public class ExtansionInfoEqualityComparer : IEqualityComparer<ExtensionInfo>
    {
        public static ExtansionInfoEqualityComparer Default = new ExtansionInfoEqualityComparer();

        public bool Equals(ExtensionInfo x, ExtensionInfo y)
        {
            if (x == null && y == null) return true;
            if (x == null || y == null) return false;
            return string.Equals(x.Identifier, y.Identifier, StringComparison.OrdinalIgnoreCase);

        }

        public int GetHashCode(ExtensionInfo obj)
        {
            if (obj == null) return 0;
            if (string.IsNullOrWhiteSpace(obj.Identifier)) return 0;
            return obj.Identifier.GetHashCode();
        }
    }

    public static class ExtensionSetFactory
    {
        static readonly XmlSerializer Serializer = new XmlSerializer(typeof(ExtensionSet));

        public static ExtensionSet Read(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException("filePath");
            using (var reader = File.OpenText(filePath))
            {
                return Serializer.Deserialize(reader) as ExtensionSet;
            }
        }

        public static void Write(string filePath, ExtensionSet extensionSet)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException("filePath");
            using (var writer = File.CreateText(filePath))
            {
                Serializer.Serialize(writer, extensionSet);
            }
        }
    }
}
