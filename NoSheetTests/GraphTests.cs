using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoSheet;
using SpreadsheetAST;
using System.Linq;

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
                // the names of the homogeneous ranges
                var rngname1 = "[" + filename + "]Sheet1!A3:A4";
                var rngname2 = "[" + filename + "]Sheet1!A1:A4";

                // there should be a two homogeneous input ranges
                string[] rs = ss.HomogeneousInputs.Select((SpreadsheetAST.Range rng) => rng.A1FullyQualified()).ToArray();

                // check for the known ranges
                Assert.IsTrue(rs.Contains(rngname1));
                Assert.IsTrue(rs.Contains(rngname2));

                Assert.IsTrue(rs.Length == 2);
            }
        }
    }
}
