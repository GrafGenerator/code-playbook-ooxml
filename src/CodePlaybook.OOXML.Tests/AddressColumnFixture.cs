using System;
using FluentAssertions;
using NUnit.Framework;
using CodePlaybook.OOXML.Models;

namespace CodePlaybook.OOXML.Tests
{
    [TestFixture]
    public class AddressColumnFixture
    {
        [TestCase("A", ExpectedResult = 1)]
        [TestCase("M", ExpectedResult = 13)]
        [TestCase("Z", ExpectedResult = 26)]
        [TestCase("AA", ExpectedResult = 27)]
        [TestCase("AZ", ExpectedResult = 52)]
        [TestCase("BA", ExpectedResult = 53)]
        [TestCase("MM", ExpectedResult = 351)]
        [TestCase("ZZ", ExpectedResult = 702)]
        [TestCase("AAA", ExpectedResult = 703)]
        [TestCase("AAZ", ExpectedResult = 728)]
        [TestCase("ABA", ExpectedResult = 729)]
        [TestCase("AZZ", ExpectedResult = 1378)]
        [TestCase("BAA", ExpectedResult = 1379)]
        [TestCase("MMM", ExpectedResult = 9139)]
        [TestCase("ZZZ", ExpectedResult = 18278)]
        [TestCase("a", ExpectedResult = 1)]
        [TestCase("az", ExpectedResult = 52)]
        public int ColumnReferencePositionCases(string reference)
        {
            return new AddressColumn(reference).NumericPosition;
        }
        
        [TestCase(1, ExpectedResult = "A")]
        [TestCase(13, ExpectedResult = "M")]
        [TestCase(26, ExpectedResult = "Z")]
        [TestCase(27, ExpectedResult = "AA")]
        [TestCase(52, ExpectedResult = "AZ")]
        [TestCase(53, ExpectedResult = "BA")]
        [TestCase(351, ExpectedResult = "MM")]
        [TestCase(702, ExpectedResult = "ZZ")]
        [TestCase(703, ExpectedResult = "AAA")]
        [TestCase(728, ExpectedResult = "AAZ")]
        [TestCase(729, ExpectedResult = "ABA")]
        [TestCase(1378, ExpectedResult = "AZZ")]
        [TestCase(1379, ExpectedResult = "BAA")]
        [TestCase(9139, ExpectedResult = "MMM")]
        [TestCase(18278, ExpectedResult = "ZZZ")]
        [TestCase(1, true, ExpectedResult = "$A")]
        [TestCase(27, true, ExpectedResult = "$AA")]
        [TestCase(703, true, ExpectedResult = "$AAA")]
        [TestCase(18278, true, ExpectedResult = "$ZZZ")]
        public string ColumnNumericPositionCases(int position, bool isFixed = false)
        {
            return new AddressColumn(position, isFixed).ReferencePosition;
            
        }
        
        [TestCase("$A", ExpectedResult = "1,f")]
        [TestCase("$M", ExpectedResult = "13,f")]
        [TestCase("$Z", ExpectedResult = "26,f")]
        [TestCase("$AA", ExpectedResult = "27,f")]
        [TestCase("$AAA", ExpectedResult = "703,f")]
        [TestCase("$ZZZ", ExpectedResult = "18278,f")]
        [TestCase("$a", ExpectedResult = "1,f")]
        [TestCase("$az", ExpectedResult = "52,f")]
        public string ColumnFixedCases(string reference)
        {
            var column = new AddressColumn(reference);
            return $"{column.NumericPosition},{(column.IsFixed ? "f" : "nf")}";
        }
        
        [TestCase("A", 0, ExpectedResult = "1,nf")]
        [TestCase("A", 1, ExpectedResult = "2,nf")]
        [TestCase("A", 100, ExpectedResult = "101,nf")]
        [TestCase("AA", -26, ExpectedResult = "1,nf")]
        [TestCase("AA", -1, ExpectedResult = "26,nf")]
        [TestCase("AA", 0, ExpectedResult = "27,nf")]
        [TestCase("AA", 1, ExpectedResult = "28,nf")]
        [TestCase("AA", 3, ExpectedResult = "30,nf")]
        [TestCase("$A", 0, ExpectedResult = "1,f")]
        [TestCase("$A", 1, ExpectedResult = "2,f")]
        [TestCase("$A", 100, ExpectedResult = "101,f")]
        [TestCase("$AA", -26, ExpectedResult = "1,f")]
        [TestCase("$AA", -1, ExpectedResult = "26,f")]
        [TestCase("$AA", 0, ExpectedResult = "27,f")]
        [TestCase("$AA", 1, ExpectedResult = "28,f")]
        [TestCase("$AA", 3, ExpectedResult = "30,f")]
        public string ColumnMoveCases(string reference, int shift)
        {
            var column = new AddressColumn(reference).Move(shift);
            return $"{column.NumericPosition},{(column.IsFixed ? "f" : "nf")}";
        }
        
        [Test]
        public void ColumnExceptionalCases()
        {
            Action act = () => new AddressColumn("AAAA");
            act.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act2 = () => new AddressColumn("");
            act2.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act3 = () => new AddressColumn("+");
            act3.Should().Throw<ArgumentException>();
            
            Action act4 = () => new AddressColumn("$$A");
            act4.Should().Throw<ArgumentException>();
            
            Action act5 = () => new AddressColumn("$");
            act5.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act6 = () => new AddressColumn(0);
            act6.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act7 = () => new AddressColumn(1).Move(-1);
            act7.Should().Throw<ArgumentOutOfRangeException>();
        }
    }
}