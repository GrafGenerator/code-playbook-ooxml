using System;
using FluentAssertions;
using NUnit.Framework;
using CodePlaybook.OOXML.Models;

namespace CodePlaybook.OOXML.Tests
{
    [TestFixture]
    public class AddressRowFixture
    {
        [TestCase(1, ExpectedResult = "1,nf")]
        [TestCase(100, ExpectedResult = "100,nf")]
        [TestCase(10000000, ExpectedResult = "10000000,nf")]
        [TestCase(1, true, ExpectedResult = "1,f")]
        [TestCase(100, true, ExpectedResult = "100,f")]
        [TestCase(10000000, true, ExpectedResult = "10000000,f")]
        public string RowNumericPositionCases(int number, bool isFixed = false)
        {
            var row = new AddressRow(number, isFixed);
            return $"{row.NumericPosition},{(row.IsFixed ? "f" : "nf")}";
        }
        
        [Test]
        public void RowNumericPositionExceptionalCases()
        {
            Action act = () => new AddressRow(0);
            act.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act2 = () => new AddressRow(-1);
            act2.Should().Throw<ArgumentOutOfRangeException>();
        }
        
        [TestCase("1", ExpectedResult = 1)]
        [TestCase("100", ExpectedResult = 100)]
        [TestCase("10000000", ExpectedResult = 10000000)]
        public int RowReferencePositionCases(string reference)
        {
            return new AddressRow(reference).NumericPosition;
        }
        
        [TestCase("1", ExpectedResult = "1,nf")]
        [TestCase("100", ExpectedResult = "100,nf")]
        [TestCase("10000000", ExpectedResult = "10000000,nf")]
        [TestCase("$1", ExpectedResult = "1,f")]
        [TestCase("$100", ExpectedResult = "100,f")]
        [TestCase("$10000000", ExpectedResult = "10000000,f")]
        public string RowFixedCases(string reference)
        {
            var row = new AddressRow(reference);
            return $"{row.NumericPosition},{(row.IsFixed ? "f" : "nf")}";
        }
        
        [TestCase("1", 0, ExpectedResult = "1,nf")]
        [TestCase("1", 1, ExpectedResult = "2,nf")]
        [TestCase("1", 100, ExpectedResult = "101,nf")]
        [TestCase("100", -99, ExpectedResult = "1,nf")]
        [TestCase("100", -1, ExpectedResult = "99,nf")]
        [TestCase("100", 0, ExpectedResult = "100,nf")]
        [TestCase("100", 1, ExpectedResult = "101,nf")]
        [TestCase("100", 100, ExpectedResult = "200,nf")]
        [TestCase("$1", 0, ExpectedResult = "1,f")]
        [TestCase("$1", 1, ExpectedResult = "2,f")]
        [TestCase("$1", 100, ExpectedResult = "101,f")]
        [TestCase("$100", -99, ExpectedResult = "1,f")]
        [TestCase("$100", -1, ExpectedResult = "99,f")]
        [TestCase("$100", 0, ExpectedResult = "100,f")]
        [TestCase("$100", 1, ExpectedResult = "101,f")]
        [TestCase("$100", 100, ExpectedResult = "200,f")]
        public string RowMoveCases(string position, int shift)
        {
            var row = new AddressRow(position).Move(shift);
            return $"{row.NumericPosition},{(row.IsFixed ? "f" : "nf")}";
        }
        
        [Test]
        public void RowReferenceExceptionalCases()
        {
            Action act = () => new AddressRow("abcd");
            act.Should().Throw<ArgumentException>();
            
            Action act2 = () => new AddressRow("");
            act2.Should().Throw<ArgumentException>();
            
            Action act3 = () => new AddressRow("0");
            act3.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act4 = () => new AddressRow("-1");
            act4.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act5 = () => new AddressRow("$$1");
            act5.Should().Throw<ArgumentException>();
            
            Action act6 = () => new AddressRow("$");
            act6.Should().Throw<ArgumentException>();
            
            Action act7 = () => new AddressRow("$0");
            act7.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act8 = () => new AddressRow("$-1");
            act8.Should().Throw<ArgumentOutOfRangeException>();
            
            Action act9 = () => new AddressRow(1).Move(-1) ;
            act9.Should().Throw<ArgumentOutOfRangeException>();
        }
    }
}