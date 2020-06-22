using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace CodePlaybook.OOXML.Models
{
    public class Xlsx: DocumentBase
    {
        private bool _documentDisposed = false;
        
        public override string MimeType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public override OpenXmlPart DocumentPart => Document.WorkbookPart;

        public override void Save()
        {
            Document.WorkbookPart.DeletePart(Document.WorkbookPart.CalculationChainPart);
            Document.Save();
        }

        public SpreadsheetDocument Document { get; }

        public Xlsx(Stream documentStream) : base(documentStream)
        {
            Contents.Position = 0;
            Document = SpreadsheetDocument.Open(Contents, true, new OpenSettings()
            {
                AutoSave = false
            });
        }
        
        public Xlsx(byte[] contents, Stream documentStream = null) : base(contents, documentStream)
        {
            Contents.Position = 0;
            Document = SpreadsheetDocument.Open(Contents, true, new OpenSettings()
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