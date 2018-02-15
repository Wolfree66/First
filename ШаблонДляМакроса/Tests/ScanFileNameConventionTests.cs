using Microsoft.VisualStudio.TestTools.UnitTesting;
using ArchiveInventoryBook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArchiveInventoryBook.Tests
{
    [TestClass()]
    public class ScanFileNameConventionTests
    {
        [TestMethod()]
        public void GetPageNumbersTest()
        {
            string rightString1 = "СКАН_ВГИП.301431.001СБ_Л2.pdf";
            IntervalPageNumbers numbers = ScanFileNameConvention.GetPageNumbersInterval(rightString1);
            Assert.AreEqual(numbers.StartPage, 2);
            Assert.AreEqual(numbers.EndPage, 2);
            string rightString2 = "СКАН_ВГИП.301431.001СБ_Л5-6.pDf";
            IntervalPageNumbers numbers2 = ScanFileNameConvention.GetPageNumbersInterval(rightString2);
            Assert.AreEqual(numbers2.StartPage, 5);
            Assert.AreEqual(numbers2.EndPage, 6);
            //Assert.Fail();
        }

        [TestMethod()]
        public void GetPageNumbersTest2()
        {
            string rightString1 = "СКАН_ВГИП.301431.001СБ.pdf";
            IntervalPageNumbers numbers = ScanFileNameConvention.GetPageNumbersInterval(rightString1);
            Assert.AreEqual(numbers.StartPage, 1);
            Assert.AreEqual(numbers.EndPage, 1);
        }

        [TestMethod()]
        public void GetPageNumbersTest3()
        {
            string rightString1 = "СКАН_ВГИП.301431.001СБ_Л22h.pdf";
            try
            {
                IntervalPageNumbers numbers = ScanFileNameConvention.GetPageNumbersInterval(rightString1);
                // raises AssertionException
                Assert.Fail();
            }
            catch (InvalidCastException)
            { }
            catch (AssertFailedException) { throw; }
 
        }
    }
}