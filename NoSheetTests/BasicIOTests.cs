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
        public const string INTERLEAVED_WB = "Interleaved.xlsx";

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
                                              Path.GetFileNameWithoutExtension(filename) + "_temp.xlsx");

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
                                              Path.GetFileNameWithoutExtension(filename) +
                                              "_temp.xlsx");

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

        [TestMethod]
        public void OpenUpdateRangeFlushClose()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(SIMPLE_WB);

            // make a copy of the file to the following path
            string newfilename = Path.Combine(Path.GetDirectoryName(filename),
                                              Path.GetFileNameWithoutExtension(filename) +
                                              "_temp.xlsx");

            // ensure that the file doesn't already exist
            Assert.IsFalse(File.Exists(newfilename));

            // load spreadsheet
            using (var ss = new ExcelSpreadsheet(filename))
            {
                // save with new name
                ss.SaveAs(newfilename);

                // some addresses
                var addr_a1 = Address.FromR1C1(1, 1, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_a2 = Address.FromR1C1(2, 1, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_a3 = Address.FromR1C1(3, 1, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_a5 = Address.FromR1C1(5, 1, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_b1 = Address.FromR1C1(1, 2, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_c1 = Address.FromR1C1(1, 3, "Sheet1", ss.WorkbookName, ss.Directory);

                // get the value in Sheet1!A1
                var a1_orig = ss.ValueAt(addr_a1);
                var a2_orig = ss.ValueAt(addr_a2);
                var a3_orig = ss.ValueAt(addr_a3);
                var a5_orig = ss.ValueAt(addr_a5);

                // get the values of the function outputs at Sheet1!B1 and Sheet1!C1
                var b1_orig = ss.ValueAt(addr_b1);
                var c1_orig = ss.ValueAt(addr_c1);

                // the new values
                var a1_newval = System.Convert.ToString(System.Convert.ToDouble(a1_orig) + 1);
                var a2_newval = System.Convert.ToString(System.Convert.ToDouble(a2_orig) + 1);
                var a3_newval = System.Convert.ToString(System.Convert.ToDouble(a3_orig) + 1);
                var a5_newval = System.Convert.ToString(System.Convert.ToDouble(a5_orig) + 1);

                // change the values
                ss.SetValueAt(addr_a1, a1_newval);
                ss.SetValueAt(addr_a2, a2_newval);
                ss.SetValueAt(addr_a3, a3_newval);
                ss.SetValueAt(addr_a5, a5_newval);

                // get the values again
                var a1_new = ss.ValueAt(addr_a1);
                var a2_new = ss.ValueAt(addr_a2);
                var a3_new = ss.ValueAt(addr_a3);
                var a5_new = ss.ValueAt(addr_a5);

                // get the new function outputs
                var b1_new = ss.ValueAt(addr_b1);
                var c1_new = ss.ValueAt(addr_c1);

                // the values should be different
                Assert.AreNotEqual(a1_orig, a1_new);
                Assert.AreNotEqual(a2_orig, a2_new);
                Assert.AreNotEqual(a3_orig, a3_new);
                Assert.AreNotEqual(a5_orig, a5_new);

                // specifically, the new values should be the values we stuck in there
                Assert.AreEqual(a1_new, a1_newval);
                Assert.AreEqual(a2_new, a2_newval);
                Assert.AreEqual(a3_new, a3_newval);
                Assert.AreEqual(a5_new, a5_newval);

                // the formula outputs should also be different
                Assert.AreNotEqual(b1_orig, b1_new);
                Assert.AreNotEqual(c1_orig, c1_new);
            }

            // cleanup
            File.Delete(newfilename);
        }

        [TestMethod]
        public void InterleavedFunctionValueWrite()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(INTERLEAVED_WB);

            // make a copy of the file to the following path
            string newfilename = Path.Combine(Path.GetDirectoryName(filename),
                                              Path.GetFileNameWithoutExtension(filename) +
                                              "_temp.xlsx");

            // ensure that the file doesn't already exist
            Assert.IsFalse(File.Exists(newfilename));

            // load spreadsheet
            using (var ss = new ExcelSpreadsheet(filename))
            {
                // save with new name
                ss.SaveAs(newfilename);

                // some value addresses
                var addr_a1 = Address.FromR1C1(1, 1, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_c4= Address.FromR1C1(3, 4, "Sheet1", ss.WorkbookName, ss.Directory);

                // some function addresses
                var addr_b1 = Address.FromR1C1(1, 2, "Sheet1", ss.WorkbookName, ss.Directory);
                var addr_c1 = Address.FromR1C1(1, 3, "Sheet1", ss.WorkbookName, ss.Directory);

                // grab formula strings from B1 and C1
                var form_b1_1 = ss.FormulaAsStringAt(addr_b1);
                var form_c1_1 = ss.FormulaAsStringAt(addr_c1);

                // grab initial values from B1 and C1
                var data_b1_1 = ss.ValueAt(addr_b1);
                var data_c1_1 = ss.ValueAt(addr_c1);

                // update A1 and C4; this should cause a range write
                // the next time we read
                ss.SetValueAt(addr_a1, "1000");
                ss.SetValueAt(addr_c4, "1000");

                // read the formula strings again
                var form_b1_2 = ss.FormulaAsStringAt(addr_b1);
                var form_c1_2 = ss.FormulaAsStringAt(addr_c1);

                // read values from B1 and C1 again
                var data_b1_2 = ss.ValueAt(addr_b1);
                var data_c1_2 = ss.ValueAt(addr_c1);

                // are the formula strings the same?
                Assert.AreEqual(form_b1_1, form_b1_2);
                Assert.AreEqual(form_c1_1, form_c1_2);

                // are the computed values correct?
                Assert.AreNotEqual(data_b1_1, data_b1_2);
                Assert.AreEqual("1009", data_b1_2);
                Assert.AreNotEqual(data_c1_1, data_c1_2);
                Assert.AreEqual("1027", data_c1_2);
            }

            // cleanup
            File.Delete(newfilename);
        }

        [TestMethod]
        public void ExcelAddressCorrectness()
        {
            // get test file path
            string filename = ResourceLoader.GetResourcePath(SIMPLE_WB);

            // load spreadsheet
            using (var ss = new ExcelSpreadsheet(filename))
            {
                // get address
                var addr_a1 = Address.FromR1C1(1, 1, "Sheet1", ss.WorkbookName, ss.Directory);
                
                // get fully-qualified spreadsheet address
                var a1fq = addr_a1.A1FullyQualified();

                // the correct string
                var correct = "[" + filename + "]Sheet1!A1";

                // the strings should be the same
                Assert.AreEqual(correct, a1fq);
            }
        }
    }
}
