namespace TTRider.ExportExtensions
{
    internal class ExtensionInfo
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Author { get; set; }
        public string Identifier { get; set; }
        public string DownloadUrl { get; set; }
        public string DownloadAs { get; set; }
        public Microsoft.VisualStudio.ExtensionManager.EnabledState State { get; set; }
    }
}
