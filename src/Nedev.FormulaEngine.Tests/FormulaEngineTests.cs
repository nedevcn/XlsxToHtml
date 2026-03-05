using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nedev.FormulaEngine;

namespace Nedev.FormulaEngine.Tests
{
    [TestClass]
    public class FormulaEngineTests
    {
        [TestMethod]
        public void BasicArithmetic_Works()
        {
            var engine = new FormulaEngine();
            Assert.AreEqual("3", engine.Evaluate("1+2"));
            Assert.AreEqual("6", engine.Evaluate("2*3"));
            Assert.AreEqual("1", engine.Evaluate("5-4"));
            Assert.AreEqual("4", engine.Evaluate("8/2"));
        }

        [TestMethod]
        public void CellReference_Resolves()
        {
            var engine = new FormulaEngine();
            engine.CellResolver = cell => cell == "A1" ? 10 : 0;
            Assert.AreEqual("15", engine.Evaluate("A1+5"));
        }

        [TestMethod]
        public void SumRange_Works()
        {
            var engine = new FormulaEngine();
            engine.CellResolver = cell => cell switch
            {
                "A1" => 1,
                "A2" => 2,
                "A3" => 3,
                _ => 0
            };
            Assert.AreEqual("6", engine.Evaluate("SUM(A1:A3)"));
        }

        [TestMethod]
        public void AverageFunction_Works()
        {
            var engine = new FormulaEngine();
            engine.CellResolver = cell => cell switch
            {
                "B1" => 2,
                "B2" => 4,
                _ => 0
            };
            Assert.AreEqual("3", engine.Evaluate("AVERAGE(B1,B2)"));
        }

        [TestMethod]
        public void MinMax_Works()
        {
            var engine = new FormulaEngine();
            engine.CellResolver = cell => cell switch
            {
                "C1" => 5,
                "C2" => -1,
                "C3" => 10,
                _ => 0
            };
            Assert.AreEqual("-1", engine.Evaluate("MIN(C1:C3)"));
            Assert.AreEqual("10", engine.Evaluate("MAX(C1:C3)"));
        }
    }
}