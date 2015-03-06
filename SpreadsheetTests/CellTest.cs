// Author: Hannah Brock
// CS 3500, Fall 2012

using SS;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using SpreadsheetUtilities;

namespace SpreadsheetTests
{
    /// <summary>
    ///This is a test class for CellTest and is intended
    ///to contain all CellTest Unit Tests
    ///</summary>
    [TestClass()]
    public class CellTest
    {

        
        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for Cell Constructor
        ///</summary>
        [TestMethod()]
        public void CellConstructorTest()
        {
            Cell target = new Cell("name", "stuff");
            Assert.AreEqual("stuff", target.Contents);
            Assert.AreEqual("name", target.Name);
        }

        /// <summary>
        ///A test for Cell Constructor
        ///</summary>
        [TestMethod()]
        public void CellConstructorTest1()
        {
            Cell target = new Cell("name");
            Assert.AreEqual("", target.Contents);
            Assert.AreEqual("name", target.Name);
        }

        /// <summary>
        ///A test for Equals
        ///</summary>
        [TestMethod()]
        public void EqualsTest()
        {
            Cell c1 = new Cell("c1", "stuff");
            Cell c2 = new Cell("c2", "stuff");
            Cell cc = new Cell("c1", "other stuff");

            Assert.IsTrue(c1.Equals(cc));
            Assert.IsFalse(c1.Equals(c2));
            Assert.IsFalse(c1.Equals(null));
            Assert.IsFalse(c1.Equals(2));
        }

        /// <summary>
        ///A test for GetHashCode
        ///</summary>
        [TestMethod()]
        public void GetHashCodeTest()
        {
            Cell c = new Cell("c", "stuff");
            Cell c2 = new Cell("c", "other stuff");

            Assert.AreEqual(c.GetHashCode(), c2.GetHashCode());
        }

        /// <summary>
        ///A test for Recalculate
        ///</summary>
        [TestMethod()]
        public void RecalculateTest()
        {
            Cell c1 = new Cell("c1", new Formula("2+3", s=>true, s=>s.ToLower()));
            Cell c2 = new Cell("c2", "other stuff");
            Cell c3 = new Cell("c3", 2.5434);

            Assert.AreEqual(2.5434, c3.Recalculate((string x) => 3.5));
            Assert.AreEqual(5.0, c1.Recalculate((string x) => 3.5));
            Assert.AreEqual(0.0, c2.Recalculate((string x) => 3.5));
        }

        /// <summary>
        ///A test for Recalculate
        ///</summary>
        [TestMethod()]
        public void RecalculateTest2()
        {
            Cell c1 = new Cell("c1", new Formula("2+3+s2", s => true, s => s.ToLower()));

            Assert.IsTrue(c1.Recalculate((string x) => { throw new ArgumentException(); }) is FormulaError);
        }

        /// <summary>
        ///A test for GetValue
        ///</summary>
        [TestMethod()]
        public void GetValueTest()
        {
            Cell c3 = new Cell("c3", 2.5434);

            Assert.AreEqual(2.5434, c3.GetValue());
        }
    }
}
