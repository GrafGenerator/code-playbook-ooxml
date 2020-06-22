using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace CodePlaybook.OOXML.Models
{
    public interface IDocument: IDisposable
    {
        Stream Contents { get; }
        string MimeType { get; }
        OpenXmlPart DocumentPart { get; }
        void Save();
    }
}