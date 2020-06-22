using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace CodePlaybook.OOXML.Models
{
    public abstract class DocumentBase: IDocument, IDisposable
    {
        private bool _streamDisposed = false;
        
        public Stream Contents => _documentStream;
        public abstract string MimeType { get; }
        public abstract OpenXmlPart DocumentPart { get; }
        public abstract void Save();

        private readonly Stream _documentStream = null;
        private readonly bool _isExternalStream;

        protected DocumentBase(Stream documentStream)
        {
            _documentStream = documentStream ?? throw new ArgumentNullException(nameof(documentStream));
            _documentStream.Position = 0;
            _isExternalStream = true;
        }
        
        protected DocumentBase(byte[] contents, Stream documentStream = null)
        {
            _documentStream = documentStream ?? new MemoryStream();
            _documentStream.Position = 0;
            _documentStream.Write(contents, 0, contents.Length);

            _isExternalStream = documentStream != null;
        }
        
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_streamDisposed)
            {
                return;
            }

            if (disposing)
            {
                if (!_isExternalStream)
                {
                    _documentStream.Dispose();
                }
            }

            _streamDisposed = true;
        }

        ~DocumentBase()
        {
            Dispose(false);
        }
    }
}