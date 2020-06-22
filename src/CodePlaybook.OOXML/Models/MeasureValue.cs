using System;

namespace CodePlaybook.OOXML.Models
{
    public class MeasureValue
    {
        // use EMU under the hood
        private readonly decimal _value = 0;

        private MeasureValue(decimal value)
        {
            _value = value;
        }
        
        public long EMUs => Convert.ToInt64(_value);
        public float Inches => Convert.ToSingle(_value) / 914400;
        public float Centimeters => Convert.ToSingle(_value) / 360000;
        public float Millimeters => Convert.ToSingle(_value) / 36000; // 10mm in 1 cm, so x10
        public float Picas => Convert.ToSingle(_value) / 152400; // 6pc in 1in, so x6;
        public float Points => Convert.ToSingle(_value) / 12700; // 72pt in 1 in, so x72;
        public float Pixels(int ppi = 96) => Inches * ppi;

        public static MeasureValue FromEMUs(long value) => new MeasureValue(value);
        public static MeasureValue FromInches(float value) => new MeasureValue(Convert.ToDecimal(Math.Round(value * 914400)));
        public static MeasureValue FromCentimeters(float value) => new MeasureValue(Convert.ToDecimal(Math.Round(value * 360000)));
        public static MeasureValue FromMillimeters(float value) => new MeasureValue(Convert.ToDecimal(Math.Round(value * 36000)));
        public static MeasureValue FromPicas(float value) => new MeasureValue(Convert.ToDecimal(Math.Round(value * 152400)));
        public static MeasureValue FromPoints(float value) => new MeasureValue(Convert.ToDecimal(Math.Round(value * 12700)));
        public static MeasureValue FromPixels(float value, int ppi = 96) => new MeasureValue(Convert.ToDecimal(Math.Round(value / ppi * 914400)));

        public static MeasureValue operator + (MeasureValue a, MeasureValue b)
        {
            if (a == null || b == null)
            {
                throw new ArgumentNullException();
            }
            
            return new MeasureValue(a._value + b._value);
        }
        
        public static MeasureValue operator - (MeasureValue a, MeasureValue b)
        {
            if (a == null || b == null)
            {
                throw new ArgumentNullException();
            }
            
            return new MeasureValue(a._value - b._value);
        }
        
        public static MeasureValue operator * (MeasureValue a, MeasureValue b)
        {
            if (a == null || b == null)
            {
                throw new ArgumentNullException();
            }
            
            return new MeasureValue(a._value * b._value);
        }
        
        public static MeasureValue operator / (MeasureValue a, MeasureValue b)
        {
            if (a == null || b == null)
            {
                throw new ArgumentNullException();
            }

            if (b._value == 0)
            {
                throw new DivideByZeroException();
            }
            
            return new MeasureValue(a._value / b._value);
        }
        
        public static bool operator < (MeasureValue a, MeasureValue b)
        {
            if (a == null || b == null)
            {
                throw new ArgumentNullException();
            }
            
            return a._value < b._value;
        }
        
        public static bool operator > (MeasureValue a, MeasureValue b)
        {
            if (a == null || b == null)
            {
                throw new ArgumentNullException();
            }
            
            return a._value > b._value;
        }
    }
}