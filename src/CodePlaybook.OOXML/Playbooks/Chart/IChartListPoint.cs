namespace CodePlaybook.OOXML.Playbooks.Chart
{
    public interface IChartListPoint<T>
    {
        public string Value { get; set; }
        public T Source { get; }
    }
}