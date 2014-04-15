using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoSheet;
using SpreadsheetAST;

namespace NoSheetTests
{
    [TestClass]
    public class GraphTests
    {
        public const string RVARIATIONS = "RangeVariations.xlsx";

        [TestMethod]
        public void ExcludesNonPerturbableRange()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(RVARIATIONS);

            using (var ss = new ExcelSpreadsheet(filename))
            {
                var rs = ss.HomogeneousInputs;

                foreach (SpreadsheetAST.Range r in rs)
                {
                    var rstr = r.A1FullyQualified();
                }

                Assert.IsTrue(rs.Length > 0);
            }
        }
    }
}
