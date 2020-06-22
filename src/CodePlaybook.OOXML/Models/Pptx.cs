using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace CodePlaybook.OOXML.Models
{
    public class Pptx: DocumentBase
    {
        private bool _documentDisposed = false;
        
        public override string MimeType => "application/vnd.openxmlformats-officedocument.presentationml.presentation";
        public override OpenXmlPart DocumentPart => Document.PresentationPart;

        public override void Save()
        {
            Document.Save();
        }

        public PresentationDocument Document { get; }

        public Pptx(Stream documentStream) : base(documentStream)
        {
            Contents.Position = 0;
            Document = PresentationDocument.Open(Contents, true, new OpenSettings()
            {
                AutoSave = false
            });
        }
        
        public Pptx(byte[] contents, Stream documentStream = null) : base(contents, documentStream)
        {
            Contents.Position = 0;
            Document = PresentationDocument.Open(Contents, true, new OpenSettings()
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