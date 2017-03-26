// Guids.cs
// MUST match guids.h

using System;

namespace TTRider.ExportExtensions
{
    static class GuidList
    {
        public const string guidVSExtensionsImportExportPkgString = "66e5502c-a117-4372-9cde-1295021e1254";
        public const string guidVSExtensionsImportExportCmdSetString = "773c93e0-b116-4db1-bfe7-39863846ba95";

        public static readonly Guid guidVSExtensionsImportExportCmdSet = new Guid(guidVSExtensionsImportExportCmdSetString);
    };
}