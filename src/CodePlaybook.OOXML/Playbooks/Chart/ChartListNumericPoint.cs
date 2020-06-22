using System;
using CodePlaybook.OOXML.Extensions;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace CodePlaybook.OOXML.Playbooks.Chart
{
    public class ChartListNumericPoint : IChartListPoint<NumericPoint>
    {
        public string Value
        {
            get => Source.GetFirstChild<NumericValue>()?.Text;
            set => Source.EnsureChild<NumericValue>().Text = value;
        }

        public NumericPoint Source { get; }

        public ChartListNumericPoint(NumericPoint source)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
        }
    }
}