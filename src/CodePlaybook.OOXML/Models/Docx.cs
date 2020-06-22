using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace CodePlaybook.OOXML.Models
{
    public class Docx: DocumentBase
    {
        private bool _documentDisposed;
        
        public override string MimeType => "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        public override OpenXmlPart DocumentPart => Document.MainDocumentPart;

        public override void Save()
        {
            Document.Save();
        }

        public WordprocessingDocument Document { get; }

        public Docx(Stream documentStream) : base(documentStream)
        {
            Contents.Position = 0;
            Document = WordprocessingDocument.Open(Contents, true, new OpenSettings()
            {
                AutoSave = false
            });
        }

        public Docx(byte[] contents, Stream documentStream = null) : base(contents, documentStream)
        {
            Contents.Position = 0;
            Document = WordprocessingDocument.Open(Contents, true, new OpenSettings()
            {
                AutoSave = false
            });
        }

        protected override void Dispose(bool disposing)
        {
            if (_documentDisposed)
            {
                return;
            }

            if (disposing)
            {
                Document.Dispose();
            }

            _documentDisposed = true;
            
            base.Dispose(disposing);
        }
    }
}