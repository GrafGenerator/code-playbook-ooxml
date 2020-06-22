using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace CodePlaybook.OOXML.Playbooks.Chart
{
    public abstract class ChartDataList<T, TPoint, TCount> : IList<string>
        where T : IChartListPoint<TPoint> 
        where TPoint : OpenXmlElement, new()
        where TCount : OpenXmlElement
    {
        protected readonly OpenXmlElement CollectionRoot;
        private readonly Func<TPoint, T> _pointFactoryFn;
        private readonly Action _updateCount;
        protected readonly List<T> Items;

        // todo: reindex points!!!
        
        
        protected ChartDataList(OpenXmlElement collectionRoot, Func<TPoint, T> pointFactoryFn, Action<TCount, int> updateCountFn)
        {
            CollectionRoot = collectionRoot;
            var countElement = collectionRoot.GetFirstChild<TCount>();
            _pointFactoryFn = pointFactoryFn;
            _updateCount = () => updateCountFn(countElement, Items.Count);
            Items = CollectionRoot.ChildElements.OfType<TPoint>()
                .Select(pointFactoryFn)
                .ToList();
        }
        
        public IEnumerator<string> GetEnumerator()
        {
            return Items.Select(x => x.Value).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(string item)
        {
            var point = new TPoint();
            var wrapper = _pointFactoryFn(point);
            wrapper.Value = item;
            Items.Add(wrapper);
            CollectionRoot.AppendChild(point);
            _updateCount();
        }

        public void Clear()
        {
            foreach (var point in Items)
            {
                point.Source.Remove();
            }
            
            Items.Clear();
            _updateCount();
        }

        public bool Contains(string item)
        {
            return Items.Select(x => x.Value).Contains(item);
        }

        public void CopyTo(string[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(string item)
        {
            var itemToRemove = Items.FirstOrDefault(x => x.Value == item);
            if (itemToRemove != null)
            {
                itemToRemove.Source.Remove();
                Items.Remove(itemToRemove);
                _updateCount();
                return true;
            }

            return false;
        }

        public int Count => Items.Count;
        public bool IsReadOnly => false;
        
        public int IndexOf(string item)
        {
            return Items.Select(x => x.Value).ToList().IndexOf(item);
        }

        public void Insert(int index, string item)
        {
            
            var point = new TPoint();
            var wrapper = _pointFactoryFn(point);
            wrapper.Value = item;
            
            Items.Insert(index, wrapper);
            var shiftedElementIndex = index + 1;

            if (shiftedElementIndex < Items.Count)
            {
                Items[shiftedElementIndex].Source.InsertBeforeSelf(point);
            }
            else
            {
                CollectionRoot.AppendChild(point);
            }
            
            _updateCount();
        }

        public void RemoveAt(int index)
        {
            if (index < 0 || index >= Items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var removedPoint = Items[index];
            Items.RemoveAt(index);
            removedPoint.Source.Remove();
            _updateCount();
        }

        public string this[int index]
        {
            get
            {
                if (index < 0 || index >= Items.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }

                return Items[index].Value;
            }
            
            set 
            {
                if (index < 0 || index >= Items.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }

                var point = Items[index];
                point.Value = value;
            }
        }
    }
}