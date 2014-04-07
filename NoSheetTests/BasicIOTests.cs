using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Reflection;
using NoSheet;

namespace NoSheetTests
{
    [TestClass]
    public class BasicIOTests
    {
        public const string SIMPLE_WB = "SimpleWorkbook.xlsx";

        [TestMethod]
        public void OpenReadClose()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(SIMPLE_WB);
            
            // load spreadsheet
            using (var ss = new ExcelSpreadsheet(filename))
            {
                var values = ss.GetAllValues();
                var formulas = ss.GetAllFormulas();

                // there should be 23 data cells
                Assert.AreEqual(23, values.Count);

                // there should be 3 formulas
                Assert.AreEqual(3, formulas.Count);
            }
        }
    }
}
