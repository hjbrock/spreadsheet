// Author: Hannah Brock
// CS 3500, Fall 2012

using SS;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Collections.Generic;
using SpreadsheetUtilities;

namespace SpreadsheetTests
{
    /// <summary>
    ///This is a test class for SpreadsheetTest and is intended
    ///to contain all SpreadsheetTest Unit Tests
    ///</summary>
    [TestClass()]
    public class SpreadsheetTest
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
        ///A test for Spreadsheet Constructor
        ///</summary>
        [TestMethod()]
        public void SpreadsheetConstructorTest()
        {
            AbstractSpreadsheet target = new Spreadsheet();
            IEnumerator<String> enumerator = target.GetNamesOfAllNonemptyCells().GetEnumerator();
            enumerator.MoveNext();
            Assert.IsNull(enumerator.Current); //check that our spreadsheet is empty
        }

        /// <summary>
        ///A test for CheckName
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        [ExpectedException(typeof(InvalidNameException))]
        public void CheckNameTest()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.CheckName("2x");
        }

        /// <summary>
        ///A test for CheckName
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        [ExpectedException(typeof(InvalidNameException))]
        public void CheckNameTest2()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.CheckName("x2x");
        }

        /// <summary>
        ///A test for CheckName
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        [ExpectedException(typeof(InvalidNameException))]
        public void CheckNameTest3()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.CheckName("x;;2");
        }

        /// <summary>
        ///A test for CheckName, check that valid name doesn't throw an exception
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        public void CheckNameTest4()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.CheckName("xsdj552");
        }

        /// <summary>
        ///A test for CheckName, check null
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        [ExpectedException(typeof(InvalidNameException))]
        public void CheckNameTest5()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.CheckName(null);
        }

        /// <summary>
        ///A test for GetCellContents
        ///</summary>
        [TestMethod()]
        public void GetCellContentsTest()
        {
            AbstractSpreadsheet target = new Spreadsheet();

            //add a formula, text, and double to separate cells
            target.SetContentsOfCell("a1", "=2+3");
            target.SetContentsOfCell("b1", "This is a string");
            target.SetContentsOfCell("c1", "2.534");

            //check that we get them back correctly
            Assert.AreEqual(new Formula("2 + 3", s => true, s => s.ToUpper()), target.GetCellContents("a1"));
            Assert.AreEqual("This is a string", target.GetCellContents("b1"));
            Assert.AreEqual(2.534, target.GetCellContents("c1"));
            //check that we get an empty string for a cell that isn't there yet
            Assert.AreEqual("", target.GetCellContents("zzz34658"));
        }

        /// <summary>
        ///A test for GetCellContents
        ///</summary>
        [TestMethod()]
        public void GetCellContentsTest2()
        {
            AbstractSpreadsheet target = new Spreadsheet();

            target.SetContentsOfCell("a1", "=2+3");
            Assert.AreEqual(new Formula("2+3", s => true, s=> s.ToUpper()), target.GetCellContents("a1"));
        }

        /// <summary>
        ///A test for GetCellContents
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void GetCellContentsTest3()
        {
            AbstractSpreadsheet target = new Spreadsheet();
            target.GetCellContents("11a1"); //expect to get an invalid name exception
        }

        /// <summary>
        ///A test for GetCellContents
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void GetCellContentsTest4()
        {
            AbstractSpreadsheet target = new Spreadsheet();
            target.GetCellContents("a111a"); //expect to get an invalid name exception
        }

        /// <summary>
        ///A test for GetDirectDependents
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        public void GetDirectDependentsTest()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            HashSet<String> expected = new HashSet<String>(new String[] { "B1", "C2" });

            target.SetContentsOfCell("A1", "2.435");
            target.SetContentsOfCell("B1", "=A1*2");
            target.SetContentsOfCell("C2", "=5+A1");
            target.SetContentsOfCell("D1", "=B1-C2");

            foreach (string name in target.GetDirectDependents("A1"))
            {
                Assert.IsTrue(expected.Contains(name));
                expected.Remove(name);
            }
            Assert.AreEqual(0, expected.Count);
        }

        /// <summary>
        ///A test for GetDirectDependents
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        [ExpectedException(typeof(ArgumentNullException))]
        public void GetDirectDependentsNullTest()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();

            target.GetDirectDependents(null);
        }

        /// <summary>
        ///A test for GetNamesOfAllNonemptyCells
        ///</summary>
        [TestMethod()]
        public void GetNamesOfAllNonemptyCellsTest()
        {
            Spreadsheet target = new Spreadsheet();
            HashSet<String> expected = new HashSet<String>(new String[] { "d1", "b1", "c2" });

            target.SetContentsOfCell("a1", "2.435");
            target.SetContentsOfCell("b1", "=a1*2");
            target.SetContentsOfCell("c2", "=5+a1");
            target.SetContentsOfCell("d1", "=b1-c2");
            target.SetContentsOfCell("a1", "");

            foreach (string name in target.GetNamesOfAllNonemptyCells())
            {
                Assert.IsTrue(expected.Contains(name));
                expected.Remove(name);
            }
            Assert.AreEqual(0, expected.Count);
        }

        /// <summary>
        ///A test for SetCellContents
        ///</summary>
        [TestMethod()]
        public void SetCellContentsTest()
        {
            Spreadsheet target = new Spreadsheet();
            HashSet<String> expected = new HashSet<String>(new String[] { "a1", "d1", "b1", "c2" });

            target.SetContentsOfCell("a1", "2.435");
            target.SetContentsOfCell("b1", "=a1*2");
            target.SetContentsOfCell("c2", "=5+a1");
            target.SetContentsOfCell("d1", "=b1-c2");
            target.SetContentsOfCell("f5", "not related");

            foreach (string cell in target.SetContentsOfCell("a1", "5.5"))
            {
                Assert.IsTrue(expected.Contains(cell));
                expected.Remove(cell);
            }
            Assert.AreEqual(0, expected.Count);
        }

        /// <summary>
        ///A test for SetCellContents
        ///</summary>
        [TestMethod()]
        public void SetCellContentsTest1()
        {
            Spreadsheet target = new Spreadsheet();
            HashSet<String> expected = new HashSet<String>(new String[] { "a1", "d1", "b1", "c2" });

            target.SetContentsOfCell("a1", "2.435");
            target.SetContentsOfCell("b1", "=a1*2");
            target.SetContentsOfCell("c2", "=5+a1");
            target.SetContentsOfCell("d1", "=b1-c2");
            target.SetContentsOfCell("f5", "not related");

            foreach (string cell in target.SetContentsOfCell("a1", "this is now a string"))
            {
                Assert.IsTrue(expected.Contains(cell));
                expected.Remove(cell);
            }
            Assert.AreEqual(0, expected.Count);
            Assert.AreEqual("this is now a string", target.GetCellContents("a1"));

            //remove b1 from dependents
            target.SetContentsOfCell("b1", "5.5");
            Assert.AreEqual(5.5, target.GetCellContents("b1"));
            expected = new HashSet<String>(new String[] { "a1", "d1", "c2" });

            foreach (string cell in target.SetContentsOfCell("a1", "=2+3"))
            {
                Assert.IsTrue(expected.Contains(cell));
                expected.Remove(cell);
            }
            Assert.AreEqual(0, expected.Count);
            Assert.AreEqual(new Formula("2+3", s => true, s => s.ToLower()), target.GetCellContents("a1"));
        }

        /// <summary>
        ///A test for SetCellContents
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SetCellContentsNullTest()
        {
            Spreadsheet target = new Spreadsheet();
            string s = null;

            target.SetContentsOfCell("a1", s);
        }

        /// <summary>
        ///A test for SetCellContents, checking for circular formulas
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(CircularException))]
        public void SetCellContentsCircularTest()
        {
            Spreadsheet target = new Spreadsheet();

            target.SetContentsOfCell("a1", "=b2*c2");
            target.SetContentsOfCell("b2", "=c2");
            target.SetContentsOfCell("c2", "=b2");
        }

        /// <summary>
        ///A test for SetCellContents, checking for circular formulas
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(CircularException))]
        public void SetCellContentsCircularTest2()
        {
            Spreadsheet target = new Spreadsheet();

            target.SetContentsOfCell("a1", "=a1");
        }

        /// <summary>
        ///A test for SetCellContents, checking for circular formulas
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(CircularException))]
        public void SetCellContentsCircularTest3()
        {
            Spreadsheet target = new Spreadsheet();

            target.SetContentsOfCell("a1", "=b2*c2");
            target.SetContentsOfCell("b2", "=d3+d5");
            target.SetContentsOfCell("d3", "=f2*3");
            target.SetContentsOfCell("f2", "=b2+2");
        }

        /// <summary>
        ///A test for GetCellsToRecalculate
        ///</summary>
        [TestMethod()]
        [DeploymentItem("Spreadsheet.dll")]
        public void GetCellsToRecalculateTest()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();

            target.SetContentsOfCell("A1", "=B2*C2");
            target.SetContentsOfCell("B2", "=D3+D5");
            target.SetContentsOfCell("F2", "=B2+A1");

            string[] expected = { "B2", "A1", "F2" };

            int i = 0;
            foreach (string name in target.GetCellsToRecalculate("B2"))
            {
                Assert.AreEqual(expected[i], name);
                i++;
            }
            Assert.AreEqual(3, i);
        }

        /// <summary>
        ///A test for Save, Reading, Saved Version
        ///</summary>
        [TestMethod()]
        public void ReadWriteTest()
        {
            Spreadsheet target = new Spreadsheet(s=>true, s=> s.ToUpper(), "test version 12345");

            target.SetContentsOfCell("a1", "=b2*c2");
            target.SetContentsOfCell("b2", "=d3+d5");
            target.SetContentsOfCell("f2", "=b2+a1");
            target.SetContentsOfCell("d6", "this is a string");
            target.SetContentsOfCell("d7", "6.54");

            target.Save("test.xml");

            Assert.AreEqual("test version 12345", target.GetSavedVersion("test.xml"));

            AbstractSpreadsheet newSheet = new Spreadsheet("test.xml", s=>true, s=> s.ToUpper(), "test version 12345");

            Assert.AreEqual(new Formula("b2*c2", s => true, s => s.ToUpper()), newSheet.GetCellContents("a1"));
            Assert.AreEqual(new Formula("d3+d5", s => true, s => s.ToUpper()), newSheet.GetCellContents("b2"));
            Assert.AreEqual(new Formula("b2+a1", s => true, s => s.ToUpper()), newSheet.GetCellContents("f2"));
            Assert.AreEqual("this is a string", newSheet.GetCellContents("d6"));
            Assert.AreEqual(6.54, newSheet.GetCellContents("d7"));
        }

        /// <summary>
        ///A test for Save
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void SaveTest2()
        {
            Spreadsheet target = new Spreadsheet();

            //test non-existent file path
            target.Save("C:\\Users\\Stuff\\Hannah\\Desktop\\test.xml");
        }

        /// <summary>
        /// From PS4 grading tests, check that we don't change the cell contents when a circular
        /// exception would be encountered.
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(CircularException))]
        public void PS4GradingTestsTest17()
        {
            AbstractSpreadsheet s = new Spreadsheet();
            try
            {
                s.SetContentsOfCell("A1", "=A2+A3");
                s.SetContentsOfCell("A2", "15");
                s.SetContentsOfCell("A3", "30");
                s.SetContentsOfCell("A2", "=A3*A1");
            }
            catch (CircularException e)
            {
                Assert.AreEqual(15, (double)s.GetCellContents("A2"), 1e-9);
                throw e;
            }
        }

        /// <summary>
        /// Test GetCellValue
        /// </summary>
        [TestMethod()]
        public void GetCellValueTest()
        {
            AbstractSpreadsheet s = new Spreadsheet();

            s.SetContentsOfCell("A1", "=A2+A3");
            s.SetContentsOfCell("A2", "15");
            s.SetContentsOfCell("A3", "30");
            s.SetContentsOfCell("b3", "this is a string");

            Assert.AreEqual(15.0, s.GetCellValue("A2"));
            Assert.AreEqual(30.0, s.GetCellValue("A3"));
            Assert.AreEqual(45.0, s.GetCellValue("A1"));
            Assert.AreEqual("this is a string", s.GetCellValue("b3"));
        }

        /// <summary>
        /// Test GetCellValue for FormulaErrors
        /// </summary>
        [TestMethod()]
        public void GetCellValueTest2()
        {
            AbstractSpreadsheet s = new Spreadsheet();

            s.SetContentsOfCell("A1", "=1+A3");

            Assert.IsTrue(s.GetCellValue("A1") is FormulaError);
        }

        /// <summary>
        /// Test GetCellValue for invalid names
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(InvalidNameException))]
        public void GetCellValueTest3()
        {
            AbstractSpreadsheet s = new Spreadsheet();

            s.GetCellValue("1a1");
        }

        /// <summary>
        /// Test reading a spreadsheet with an invalid version
        /// </summary>
        [TestMethod()]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void ReadWriteTest2()
        {
            AbstractSpreadsheet sheet = new Spreadsheet(s=>true,s=>s.ToUpper(),"version");
            sheet.Save("ReadWriteTest2.xml");

            sheet = new Spreadsheet("ReadWriteTest2.xml", s => true, s => s.ToUpper(), "different version");
        }

        /// <summary>
        /// Test reading spreadsheet with name only in cell
        /// </summary>
        [TestMethod()]
        public void ReadWriteTest3()
        {
            DirectoryInfo dir = new DirectoryInfo(TestContext.TestDir);
            dir = dir.Parent.Parent; //move to test project dir
            AbstractSpreadsheet sheet = new Spreadsheet(dir.FullName + "\\SpreadsheetTests\\badFormat.xml", s => true, s => s.ToUpper(), "1.0");
            Assert.AreEqual("", sheet.GetCellContents("f2"));
            Assert.AreEqual("this is a string", sheet.GetCellContents("d6"));
            Assert.AreEqual(6.54, sheet.GetCellContents("d7"));
        }

        /// <summary>
        /// Test reading spreadsheet with contents only in cell
        /// </summary>
        [TestMethod()]
        public void ReadWriteTest4()
        {
            DirectoryInfo dir = new DirectoryInfo(TestContext.TestDir);
            dir = dir.Parent.Parent; //move to test project dir
            try
            {
                AbstractSpreadsheet sheet = new Spreadsheet(dir.FullName + "\\SpreadsheetTests\\badFormat3.xml", s => true, s => s.ToUpper(), "1.0");
            }
            catch (SpreadsheetReadWriteException e)
            {
                Assert.AreEqual("XML formatted incorrectly", e.Message);
            }
        }

        /// <summary>
        /// Test reading spreadsheet with contents outside of cell
        /// </summary>
        [TestMethod()]
        public void ReadWriteTest5()
        {
            DirectoryInfo dir = new DirectoryInfo(TestContext.TestDir);
            dir = dir.Parent.Parent; //move to test project dir
            try
            {
                AbstractSpreadsheet sheet = new Spreadsheet(dir.FullName + "\\SpreadsheetTests\\badFormat2.xml", s => true, s => s.ToUpper(), "1.0");
            }
            catch (SpreadsheetReadWriteException e)
            {
                Assert.AreEqual("XML formatted incorrectly", e.Message);
            }
        }

        /// <summary>
        /// Check that empty cell doesn't result in errors
        /// </summary>
        [TestMethod()]
        public void ReadWriteTest6()
        {
            DirectoryInfo dir = new DirectoryInfo(TestContext.TestDir);
            dir = dir.Parent.Parent; //move to test project dir
            AbstractSpreadsheet sheet = new Spreadsheet(dir.FullName + "\\SpreadsheetTests\\badFormat4.xml", s => true, s => s.ToUpper(), "1.0");
        }
    }
}
