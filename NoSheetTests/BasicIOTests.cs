using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Reflection;
using NoSheet;
using SpreadsheetAST;

namespace NoSheetTests
{
    [TestClass]
    public class BasicIOTests
    {
        public const string SIMPLE_WB = "SimpleWorkbook.xlsx";

        // Note that testing that IDisposable objects do the right
        // thing is not really feasible because 1) GC is nondeterministic,
        // and in DEBUG mode, manually calling 2) GC.Collect() does not
        // do what you'd expect, because scoping rules are different.
        [TestMethod]
        public void OpenReadClose()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(SIMPLE_WB);

            // load spreadsheet
            using (var ss = new ExcelSpreadsheet(filename))
            {
                var values = ss.Values;
                var formulas = ss.Formulas;

                // there should be 23 data cells
                Assert.AreEqual(23, values.Count);

                // there should be 3 formulas
                Assert.AreEqual(3, formulas.Count);
            }
        }

        [TestMethod]
        public void OpenSaveAs()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(SIMPLE_WB);

            // get SaveAs path
            string newfilename = Path.Combine(Path.GetDirectoryName(filename),
                                              Path.GetFileNameWithoutExtension(filename) + "2.xlsx");

            // ensure that the file doesn't already exist
            Assert.IsFalse(File.Exists(newfilename));

            // load spreadsheet
            using (var ss = new ExcelSpreadsheet(filename))
            {
                // the first time we save with the new filename,
                // the method should save, because the file
                // does not exist
                Assert.IsTrue(ss.SaveAs(newfilename));

                // file should exist
                Assert.IsTrue(File.Exists(newfilename));

                // but the second time we save with the same
                // filename, the method should fail
                Assert.IsFalse(ss.SaveAs(newfilename));
            }

            // cleanup so that we can run this test again
            File.Delete(newfilename);
        }

        [TestMethod]
        public void OpenUpdateFlushClose()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(SIMPLE_WB);

            // make a copy of the file to the following path
            string newfilename = Path.Combine(Path.GetDirectoryName(filename),
                                              Path.GetFileNameWithoutExtension(filename) + "_flushtest.xlsx");

            // ensure that the file doesn't already exist
            Assert.IsFalse(File.Exists(newfilename));

            // load spreadsheet
            using (var ss = new ExcelSpreadsheet(filename))
            {
                // save with new name
                ss.SaveAs(newfilename);

                // some addresses
                var addr_a1 = Address.FromR1C1(1, 1, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_b1 = Address.FromR1C1(1, 2, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_c1 = Address.FromR1C1(1, 3, "Sheet1", ss.WorkbookName, ss.Directory);

                // get the value in Sheet1!A1
                var a1_orig = ss.ValueAt(addr_a1);

                // get the values of the function outputs at Sheet1!B1 and Sheet1!C1
                var b1_orig = ss.ValueAt(addr_b1);
                var c1_orig = ss.ValueAt(addr_c1);

                // the new value
                var newval = System.Convert.ToString(System.Convert.ToDouble(a1_orig) + 1);

                // change the value
                ss.SetValueAt(addr_a1, newval);

                // get the value again
                var a1_new = ss.ValueAt(addr_a1);

                // get the new function outputs
                var b1_new = ss.ValueAt(addr_b1);
                var c1_new = ss.ValueAt(addr_c1);

                // the values should be different
                Assert.AreNotEqual(a1_orig, a1_new);

                // specifically, a1_new should be the value we stuck in there
                Assert.AreEqual(a1_new, newval);

                // the formula outputs should also be different
                Assert.AreNotEqual(b1_orig, b1_new);
                Assert.AreNotEqual(c1_orig, c1_new);
            }

            // cleanup
            File.Delete(newfilename);
        }

        [TestMethod]
        public void NonexistentFileOpenAttempt()
        {
            // this file should not exist
            string filename = @"C:\Stuff\FOOBAZ.xls";

            // make sure that it doesn't
            Assert.IsFalse(File.Exists(filename));

            // exercise file open exception
            try
            {
                using (var ss = new ExcelSpreadsheet(filename))
                {
                    // should never get here
                }
            }
            catch (FileNotFoundException)
            {
                return;
            }

            // if we got here, then our ExcelSpreadsheet constructor
            // failed to throw the appropriate exception
            Assert.Fail("ExcelSpreadsheet should throw a FileNotFoundException when a file does not exist.");
        }
    }
}
