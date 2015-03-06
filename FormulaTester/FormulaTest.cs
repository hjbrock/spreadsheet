// Author: Hannah Brock
// CS 3500, Fall 2012

using SpreadsheetUtilities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace FormulaTester
{


    /// <summary>
    ///This is a test class for FormulaTest and is intended
    ///to contain all FormulaTest Unit Tests
    ///</summary>
    [TestClass()]
    public class FormulaTest
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
        ///A test for Formula Constructor with too many left parens
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorLtParensTest()
        {
            Formula target = new Formula("(2 + 5 * 4",s=>true,s=>s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with too many left parens
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorLtParensTest2()
        {
            Formula target = new Formula("(2 + 5) * (4",s=>true,s=>s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with too many right parens
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorRtParensTest()
        {
            Formula target = new Formula("(2 + 5 * 4))", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with too many right parens
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorRtParensTest2()
        {
            Formula target = new Formula("((2 + 5)) * 4)", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with operators directly before an opening parens
        ///</summary>
        [TestMethod()]
        public void FormulaConstructorOpParensTest1()
        {
            try
            {
                Formula target = new Formula("(+2+5)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(*2+5)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(/2+5)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(-2+5)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("-2+5", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }
        }

        /// <summary>
        ///A test for Formula Constructor with operators directly before a closing parens
        ///</summary>
        [TestMethod()]
        public void FormulaConstructorOpParensTest2()
        {
            try
            {
                Formula target = new Formula("(2+5+)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(2+5)2", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(2+5*)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(2+5/)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(2+5-)", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("2+5-", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }

            try
            {
                Formula target = new Formula("(2+5)+2-()", s => true, s => s.ToUpper());
                Assert.Fail();
            }
            catch (FormulaFormatException) { }
        }

        /// <summary>
        ///A test for Formula Constructor with no tokens
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorEmptyTest()
        {
            Formula target = new Formula("", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with number followed by a number
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorNumTest()
        {
            Formula target = new Formula("2+.945.95456", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with invalid variable as token
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorVariableTest()
        {
            Formula target = new Formula("xx", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with invalid variable as token
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorVariableTest2()
        {
            Formula target = new Formula("!x", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with invalid variable as token
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorVariableTest3()
        {
            Formula target = new Formula("xxx33334x3", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with invalid variable as token
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorVariableTest4()
        {
            Formula target = new Formula("x2x2", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with invalid variable as token
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(FormulaFormatException))]
        public void FormulaConstructorVariableTest5()
        {
            Formula target = new Formula("x..9", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Formula Constructor with correct parens matching.
        ///All constructions should complete without exceptions.
        ///</summary>
        [TestMethod()]
        public void FormulaConstructorCorrectParensTest()
        {
            Formula target = new Formula("(2 + 5 * (4))", s => true, s => s.ToUpper());
            target = new Formula("((2 + 5) * (4))", s => true, s => s.ToUpper());
            target = new Formula("(2 + 5 * (4))", s => true, s => s.ToUpper());
            target = new Formula("(2 + 5 * (4))", s => true, s => s.ToUpper());
            target = new Formula("((2 + 5)) * 4", s => true, s => s.ToUpper());
            target = new Formula("(2) + (5) * (4)", s => true, s => s.ToUpper());
        }

        /// <summary>
        ///A test for Equals
        ///</summary>
        [TestMethod()]
        public void EqualsTest()
        {
            Formula target = new Formula("2+7", s => true, s => s.ToUpper());
            Formula compare = new Formula("2 + 7", s => true, s => s.ToUpper());
            Assert.IsTrue(target.Equals(compare));

            target = new Formula("x1+y1", s => true, s => s.ToUpper());
            compare = new Formula("x1 + y1", s => true, s => s.ToUpper());
            Assert.IsTrue(target.Equals(compare));

            target = new Formula("x1+y1", s => true, s => s.ToUpper());
            compare = new Formula("y1 + x1", s => true, s => s.ToUpper());
            Assert.IsFalse(target.Equals(compare));

            target = new Formula("2.0000+y1", s => true, s => s.ToUpper());
            compare = new Formula("2.00 + y1", s => true, s => s.ToUpper());
            Assert.IsTrue(target.Equals(compare));

            target = new Formula("(x1+y1)", s => true, s => s.ToUpper());
            compare = new Formula("x1 + y1", s => true, s => s.ToUpper());
            Assert.IsFalse(target.Equals(compare));

            target = new Formula("x1 * y1", s => true, s => s.ToUpper());
            compare = new Formula("(x1) * y1", s => true, s => s.ToUpper());
            Assert.IsFalse(target.Equals(compare));

            target = new Formula("(5.0 + x1) * ((2+3)/(4.70))", s => true, s => s.ToUpper());
            compare = new Formula(" (5 + x1) * ((2+3)/(4.7))", s => true, s => s.ToUpper());
            Assert.IsTrue(target.Equals(compare));

            target = new Formula("((5.0 + x1)) * ((2+3)/(4.70))", s => true, s => s.ToUpper());
            compare = new Formula(" (5 + x1) * ((2+3)/(4.7))", s => true, s => s.ToUpper());
            Assert.IsFalse(target.Equals(compare));

            target = new Formula("(5.0 + x1) * ((2+3)/(4.60))", s => true, s => s.ToUpper());
            compare = new Formula(" (5 + x1) * ((2+3)/(4.7))", s => true, s => s.ToUpper());
            Assert.IsFalse(target.Equals(compare));

            target = new Formula("(5.0 + x1) * ((2+3)/4.70)", s => true, s => s.ToUpper());
            compare = new Formula(" (5 + x1) * ((2+3)/x2)", s => true, s => s.ToUpper());
            Assert.IsFalse(target.Equals(compare));

            Assert.IsFalse(target.Equals(null));
            Assert.IsFalse(target.Equals(1));
        }

        /// <summary>
        /// Test single value and variable lookup.
        /// </summary>
        [TestMethod]
        public void TestEvaluateVariable()
        {
            Assert.AreEqual(20.0, (new Formula("b2", s => true, s => s.ToLower()).Evaluate(Lookup)));
        }

        /// <summary>
        /// Test the order of operations
        /// </summary>
        [TestMethod]
        public void TestOrderOps1()
        {
            Assert.AreEqual(9.0, (new Formula("2+4*3-5", s => true, s => s.ToUpper())).Evaluate(Lookup));
            Assert.AreEqual(-10.0, (new Formula("(2+4)*(3-5)+2", s => true, s => s.ToUpper())).Evaluate(Lookup));
        }

        /// <summary>
        /// Test the order of operations
        /// </summary>
        [TestMethod]
        public void TestOrderOps2()
        {
            Assert.AreEqual(4.0, (new Formula("2+4/2", s => true, s => s.ToUpper())).Evaluate(Lookup));
        }

        /// <summary>
        /// Test a combination of parenthesis, operators, and variables
        /// </summary>
        [TestMethod]
        public void TestEvaluateCombo()
        {
            Assert.AreEqual(47.0, (new Formula("(2 + X6) * (5 + 0) + 2", s => true, s => s.ToUpper())).Evaluate(Lookup));
            Assert.AreEqual(16.0, (new Formula("(2 + X6) + (5 + 0) + 2", s => true, s => s.ToUpper())).Evaluate(Lookup));
            Assert.AreEqual(14.0, (new Formula("(2 * X6)", s => true, s => s.ToUpper())).Evaluate(Lookup));
        }

        /// <summary>
        /// Test that we throw an exception for dividing by zero
        /// </summary>
        [TestMethod]
        public void TestEvaluateDivideByZero()
        {
            object response = (new Formula("2/0", s => true, s => s.ToUpper())).Evaluate(Lookup);
            Assert.IsInstanceOfType(response, typeof(FormulaError));
        }

        /// <summary>
        /// Test non-existent variable
        /// </summary>
        [TestMethod]
        public void TestBadVariable()
        {
            object response = (new Formula("2+x3", s => true, s => s.ToUpper())).Evaluate(Lookup);
            Assert.IsInstanceOfType(response, typeof(FormulaError));
        }

        /// <summary>
        /// Lookup delegate to utilize for evaluator tests.
        /// </summary>
        /// <param name="v">The variable name</param>
        /// <returns>Returns the value of the variable</returns>
        private double Lookup(string v)
        {
            if (v.Equals("a"))
                return 1;
            if (v.Equals("b2"))
                return 20;
            if (v.Equals("c"))
                return 5;
            if (v.Equals("X6"))
                return 7;
            else
                throw new ArgumentException();
        }

        /// <summary>
        ///A test for GetHashCode
        ///</summary>
        [TestMethod()]
        public void GetHashCodeTest()
        {
            Formula target = new Formula("(2+5.5)-2", s => true, s => s.ToUpper());
            Formula compare = new Formula("( 2.0 + 5.500 )- 2", s => true, s => s.ToUpper());
            Assert.AreEqual(target.GetHashCode(), compare.GetHashCode());
        }

        /// <summary>
        ///A test for GetVariables
        ///</summary>
        [TestMethod()]
        public void GetVariablesTest()
        {
            Formula target = new Formula("(a2+b3)*b4*b3", s => true, s => s.ToUpper());
            IEnumerator<string> vars = target.GetVariables().GetEnumerator();
            vars.MoveNext();
            Assert.AreEqual("A2", vars.Current);
            vars.MoveNext();
            Assert.AreEqual("B3", vars.Current);
            vars.MoveNext();
            Assert.AreEqual("B4", vars.Current);
            Assert.IsFalse(vars.MoveNext());
        }

        /// <summary>
        ///A test for ToString
        ///</summary>
        [TestMethod()]
        public void ToStringTest()
        {
            Formula target = new Formula("( 2.0 + 5 )- 2", s => true, s => s.ToUpper());
            Formula compare = new Formula(target.ToString(), s => true, s => s.ToUpper());
            Assert.IsTrue(target.Equals(compare));
            Assert.IsFalse(target.ToString().Contains(" "));

            foreach (char c in target.ToString())
                Assert.AreNotEqual(' ', c);

            target = new Formula("( 2 + b2 )- (x3*2.5 )", s => true, s => s.ToUpper());
            compare = new Formula(target.ToString(), s => true, s => s.ToUpper());
            Assert.IsTrue(target.Equals(compare));
            Assert.IsFalse(target.ToString().Contains(" "));
        }

        /// <summary>
        ///A test for op_Equality
        ///</summary>
        [TestMethod()]
        public void op_EqualityTest()
        {
            Formula target = new Formula("(5.0 + x1) * ((2+3)/(4.70))", s => true, s => s.ToUpper());
            Formula compare = new Formula(" (5 + x1) * ((2+3)/(4.7))", s => true, s => s.ToUpper());
            Assert.IsTrue(target == compare);

            target = new Formula("2+3", s => true, s => s.ToUpper());
            compare = null;
            Assert.IsFalse(target == compare);

            target = null;
            compare = new Formula("2+3", s => true, s => s.ToUpper());
            Assert.IsFalse(target == compare);

            target = null;
            compare = null;
            Assert.IsTrue(target == compare);

            target = new Formula("((5.0 + x1)) * ((2+3)/(4.70))", s => true, s => s.ToUpper());
            compare = new Formula(" (5 + x1) * ((2+3)/(4.7))", s => true, s => s.ToUpper());
            Assert.IsFalse(target == compare);
        }

        /// <summary>
        ///A test for op_Inequality
        ///</summary>
        [TestMethod()]
        public void op_InequalityTest()
        {
            Formula target = new Formula("(5.0 + x1) * ((2+3)/(4.70))", s => true, s => s.ToUpper());
            Formula compare = new Formula(" (5 + x1) * ((2+3)/(4.7))", s => true, s => s.ToUpper());
            Assert.IsFalse(target != compare);

            target = new Formula("2+3", s => true, s => s.ToUpper());
            compare = null;
            Assert.IsTrue(target != compare);

            target = null;
            compare = new Formula("2+3", s => true, s => s.ToUpper());
            Assert.IsTrue(target != compare);

            target = null;
            compare = null;
            Assert.IsFalse(target != compare);

            target = new Formula("((5.0 + x1)) * ((2+3)/(4.70))", s => true, s => s.ToUpper());
            compare = new Formula(" (5 + x1) * ((2+3)/(4.7))", s => true, s => s.ToUpper());
            Assert.IsTrue(target != compare);
        }
    }
}
