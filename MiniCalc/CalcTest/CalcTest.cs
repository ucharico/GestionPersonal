using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MiniCalc;

namespace CalcTest
{
    [TestClass]
    public class CalcTest
    {
        Calculator testCalc;
        [TestInitialize]
        public void Initialize()
        {
            testCalc = new Calculator();
        }
        [TestCleanup]
        public void Cleanup()
        {
            testCalc = null;
        }
        [TestMethod]
        public void TestAdd()
        {
            Console.Out.WriteLine("TestAddition called");
            //test for case of zeros'
            Assert.AreEqual(0, testCalc.Add(0, 0), "Adding 0 to 0 should produce 0");
            //test that param ordering isn't important
            Assert.AreEqual(1, testCalc.Add(1, 0), "Adding 1 to 0 should produce 1");
            Assert.AreEqual(1, testCalc.Add(0, 1), "Adding 0 to 1 should produce 1");
            //test for non zero case
            Assert.AreEqual(3, testCalc.Add(1, 2), "Adding 1 to 2 should produce 3");
            int nResult;
            try
            {
                nResult = testCalc.Add(int.MaxValue, int.MaxValue);
                Assert.Fail("Should throw a ResultOutofRangeException");
            }
            catch (ResultOutOfRangeException)
            {
            }
            testCalc = null;
        }
        [TestMethod]
        public void TestAddNegatives()
        {
            Console.Out.WriteLine("TestAddNegatives called");
            int nResult;
            try
            {
                nResult = testCalc.Add(-1, 0);
                Assert.Fail("Should throw a NegativeParameterException");
            }
            catch (NegativeParameterException)
            { }
            try
            {
                nResult = testCalc.Add(0, -1);
                Assert.Fail("Should throw a NegativeParameterException");
            }
            catch (NegativeParameterException)
            { }
            try
            {
                nResult = testCalc.Add(int.MinValue, int.MinValue);
                Assert.Fail("Should throw a NegativeParameterException");
            }
            catch (NegativeParameterException)
            { }
            testCalc = null;
        }
        [TestMethod]
        public void TestSubtract()
        {
            Console.Out.WriteLine("TestSubtract called");
            int nResult;
            Assert.AreEqual(0, testCalc.Subtract(0, 0), "Subtracting 0 from 0 should produce 0");
            Assert.AreEqual(1, testCalc.Subtract(0, 1), "Subtracting 0 from 1 should produce 1");
            try
            {
                nResult = testCalc.Subtract(1, 0);
                Assert.Fail("Subtracting 1 from 0 should throw a ResultOutofRangeException");
            }
            catch (ResultOutOfRangeException)
            { }
            Assert.AreEqual(0, testCalc.Subtract(int.MaxValue, int.MaxValue), "Subtracting max value from max value should produce 0");
            try
            {
                nResult = testCalc.Subtract(-1, 0);
                Assert.Fail("Should throw a NegativeParameterException");
            }
            catch (NegativeParameterException)
            { }
            try
            {
                nResult = testCalc.Subtract(0, -1);
                Assert.Fail("Should throw a NegativeParameterException");
            }
            catch (NegativeParameterException)
            { }
            try
            {
                nResult = testCalc.Subtract(int.MinValue, int.MinValue);
                Assert.Fail("Should throw a NegativeParameterException");
            }
            catch (NegativeParameterException)
            { }
        }
        [TestMethod]
        public void TestCheckForNegativeNumbers()
        {
            Console.Out.WriteLine("TestSubtract called");
            try
            {
                testCalc.CheckForNegativeNumbers(0, 0);
            }
            catch (NegativeParameterException)
            {
                Assert.Fail("Zeros are not negative numbers");
            }
            try
            {
                testCalc.CheckForNegativeNumbers(1, 1);
            }
            catch (NegativeParameterException)
            {
                Assert.Fail("1's are not negative numbers");
            }
            try
            {
                testCalc.CheckForNegativeNumbers(
                int.MaxValue, int.MaxValue);
            }
            catch (NegativeParameterException)
            {
                Assert.Fail("Max Vals are not negative numbers");
            }
            try
            {
                testCalc.CheckForNegativeNumbers(-1, -1);
                Assert.Fail("-1's are negative numbers");
            }
            catch (NegativeParameterException)
            { }
            try
            {
                testCalc.CheckForNegativeNumbers(
                int.MinValue, int.MinValue);
                Assert.Fail("Min Vals are negative numbers");
            }
            catch (NegativeParameterException)
            { }
        }
        [TestMethod]
        public void TestCalculate()
        {
            Console.Out.WriteLine("TestCalculate called");
            Assert.AreEqual(2, testCalc.Calculate(1, CalcOperation.eAdd, 1), "Adding 1 to 1 failed");
            Assert.AreEqual(0, testCalc.Calculate(1, CalcOperation.eSubtract, 1), "Subtracting 1 from 1 failed");
        }
    }


}
