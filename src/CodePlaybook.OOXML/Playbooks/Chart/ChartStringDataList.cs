using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace CodePlaybook.OOXML.Playbooks.Chart
{
    public class ChartStringDataList : ChartDataList<ChartListStringPoint, StringPoint, PointCount> 
    {
        public ChartStringDataList(OpenXmlElement collectionRoot) 
            : base(collectionRoot, PointFactory, CountUpdate)
        {
        }

        private static ChartListStringPoint PointFactory(StringPoint pt) => new ChartListStringPoint(pt);
        private static void CountUpdate(PointCount pointCount, int count)
        {
            if (pointCount != null)
            {
                pointCount.Val = (uint)count;
            }
        }
    }
}