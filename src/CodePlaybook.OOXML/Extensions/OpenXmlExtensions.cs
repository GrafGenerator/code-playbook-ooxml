using System;
using System.Linq;
using DocumentFormat.OpenXml;

namespace CodePlaybook.OOXML.Extensions
{
    public static class OpenXmlExtensions
    {
        public static T EnsureChild<T>(this OpenXmlElement target)
            where T: OpenXmlElement, new()
        {
            return target.GetFirstChild<T>() ?? target.AppendChild(new T());
        }
        
        public static void ClearElementsAfter(this OpenXmlElement target)
        {
            if (target != null)
            {
                foreach (var element in target.ElementsAfter().ToList())
                {
                    element.Remove();
                }
            }
        }

        public static void ClearElementsAfter<T>(this OpenXmlElement target)
            where T: OpenXmlElement
        {
            if (target != null)
            {
                foreach (var element in target.ElementsAfter().OfType<T>().ToList())
                {
                    element.Remove();
                }
            }
        }
        
        public static OpenXmlElement TraverseToRoot(this OpenXmlElement source, params Type[] rootTypes)
        {
            var parent = source.Parent;
            while (parent != null && !rootTypes.Contains(parent.GetType()))
            {
                parent = parent.Parent;
            }

            return parent;
        }

        public static OpenXmlElement TraverseToRoot<TR>(this OpenXmlElement source)
            => TraverseToRoot(source, typeof(TR));
        
        public static OpenXmlElement TraverseToRoot<TR1, TR2>(this OpenXmlElement source)
            => TraverseToRoot(source, typeof(TR1), typeof(TR2));
        
        public static OpenXmlElement TraverseToRoot<TR1, TR2, TR3>(this OpenXmlElement source)
            => TraverseToRoot(source, typeof(TR1), typeof(TR2), typeof(TR3));
    }
}