using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ShoppingCart_Excel
{
    [TestClass]
    public class UnitTest1
    {
        public Class1 c = new Class1();
        public Input_Class c1 = new Input_Class();

        [TestInitialize]
        public void Setup()
        {
            c.startBrowser();
        }

        [TestMethod]
        public void GettingDataforTest()
        {
            c.getProductName();
            c.GetBasicDetails();
        }

        [TestMethod]
        public void PerformingTest()
        {
            c.AddProduct();
            c.ShoppingCart();
            c.verifyqauntity();
            c.verifyUnitPrice();
            c.verifyTotalPrice();
        }

        [TestCleanup]
        public void TearDown()
        {
            c.stopBrowser();
        }
    }
}
