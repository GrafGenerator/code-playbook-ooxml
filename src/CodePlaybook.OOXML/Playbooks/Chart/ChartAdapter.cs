using System;
using System.Collections.Generic;
using System.Linq;
using CodePlaybook.OOXML.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CodePlaybook.OOXML.Playbooks.Chart
{
    public class ChartAdapter: IDisposable
    {
        private DocumentFormat.OpenXml.Drawing.Charts.Chart _chart;
        private readonly EmbeddedPackagePart _package;
        private readonly Xlsx _embeddedDocument;
        private readonly Dictionary<string, Sheet> _sheetsMap;

        public ChartAdapter(ChartPart chartPart)
        {
            _chart = chartPart.ChartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>() ??
                     throw new ArgumentException("ChartPart contains no Chart object", nameof(chartPart));

            _package = chartPart.EmbeddedPackagePart;
            if (_package != null)
            {
                using var stream = _package.GetStream();
                var embeddedBytes = new byte[stream.Length];
                stream.Position = 0;
                stream.Read(embeddedBytes, 0, (int) stream.Length);
                _embeddedDocument = new Xlsx(embeddedBytes);

                _sheetsMap = _embeddedDocument.Document?.WorkbookPart?.Workbook?.Sheets?.ChildElements.OfType<Sheet>()
                    .Where(x => !string.IsNullOrEmpty(x.Name?.Value))
                    .ToDictionary(x => x.Name.Value);
                
                if (_sheetsMap == null)
                {
                    throw new ArgumentException("Embedded document is corrupted.");
                }
            }
        }

        public bool TryGetSeries(Range address, out SeriesAdapter seriesAdapter)
        {
            throw new NotImplementedException();
        }

        public void Apply()
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            _embeddedDocument?.Dispose();
        }
    }
}