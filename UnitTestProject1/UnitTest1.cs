using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using ModernExcel;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Assert.AreEqual("A", MyAddress.GetColumnName(1));
            Assert.AreEqual("B", MyAddress.GetColumnName(2));
            Assert.AreEqual("Z", MyAddress.GetColumnName(26));
            Assert.AreEqual("AA", MyAddress.GetColumnName(27));
            Assert.AreEqual("AB", MyAddress.GetColumnName(28));
            Assert.AreEqual("AZ", MyAddress.GetColumnName(52));
            Assert.AreEqual("BA", MyAddress.GetColumnName(53));
            Assert.AreEqual("AAA", MyAddress.GetColumnName(703));
        }
    }
}
