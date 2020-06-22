using System;
using FluentAssertions;
using NUnit.Framework;
using CodePlaybook.OOXML.Models;

namespace CodePlaybook.OOXML.Tests
{
    [TestFixture]
    public class AddressFixture
    {
        [TestCase("A1", ExpectedResult = "A,1")]
        [TestCase("AA111", ExpectedResult = "AA,111")]
        [TestCase("$A1", ExpectedResult = "$A,1")]
        [TestCase("A$1", ExpectedResult = "A,$1")]
        [TestCase("$AA111", ExpectedResult = "$AA,111")]
        [TestCase("AA$111", ExpectedResult = "AA,$111")]
        [TestCase("$AA$111", ExpectedResult = "$AA,$111")]
        [TestCase("a1", ExpectedResult = "A,1")]
        public string AddressCases(string reference)
        {
            var address = new Address(reference);
            return $"{address.Column.ReferencePosition},{address.Row.ReferencePosition}";
        }
        
        [TestCase("A")]
        [TestCase("1")]
        [TestCase("$$A1")]
        [TestCase("$A$$1")]
        public void AddressExceptionalCases(string reference)
        {
            Action act = () => new Address(reference);
            act.Should().Throw<Exception>();
        }
    }
}