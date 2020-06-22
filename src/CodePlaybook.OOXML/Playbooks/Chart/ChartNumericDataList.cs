using CodePlaybook.OOXML.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace CodePlaybook.OOXML.Playbooks.Chart
{
    public class ChartNumericDataList : ChartDataList<ChartListNumericPoint, NumericPoint, PointCount> 
    {
        public ChartNumericDataList(OpenXmlElement collectionRoot) 
            : base(collectionRoot, PointFactory, CountUpdate)
        {
        }

        private static ChartListNumericPoint PointFactory(NumericPoint pt) => new ChartListNumericPoint(pt);
        private static void CountUpdate(PointCount pointCount, int count)
        {
            if (pointCount != null)
            {
                pointCount.Val = (uint)count;
            }
        }

        public string Format
        {
            get => CollectionRoot.GetFirstChild<FormatCode>()?.Text;
            set => CollectionRoot.EnsureChild<FormatCode>().Text = value;
        }
    }
}