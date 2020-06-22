using System;
using CodePlaybook.OOXML.Extensions;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace CodePlaybook.OOXML.Playbooks.Chart
{
    public class ChartListStringPoint : IChartListPoint<StringPoint>
    {
        public string Value
        {
            get => Source.GetFirstChild<NumericValue>()?.Text;
            set => Source.EnsureChild<NumericValue>().Text = value;
        }

        public StringPoint Source { get; }

        public ChartListStringPoint(StringPoint source)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
        }
    }
}