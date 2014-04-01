using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.FSharp.Core;
using System.IO;
using System.Text.RegularExpressions;

namespace NoSheet
{
    public class ExcelSpreadsheet : ISpreadsheet, IDisposable
    {
        // COM handles
        private Excel.Application _app;
        private Excel.Workbook _wb;
        private HashSet<Excel.Worksheet> _wss = new HashSet<Excel.Worksheet>();

        // data storage
        private Dictionary<AST.Address, string> _data = new Dictionary<AST.Address, string>();
        private Dictionary<AST.Address, string> _formulas = new Dictionary<AST.Address, string>();

        // formula string regex
        private Regex fpatt = new Regex("^=", RegexOptions.Compiled);

        // All of the following private enums are poorly documented
        private enum XlCorruptLoad
        {
            NormalLoad = 0,
            RepairFile = 1,
            ExtractData = 2
        }

        private enum XlUpdateLinks
        {
            Yes = 2,
            No = 0
        }

        private enum XlPlatform
        {
            Macintosh = 1,
            Windows = 2,
            MSDOS = 3
        }

        private enum CellType
        {
            Data,
            Formula
        }

        public ExcelSpreadsheet(string filename)
        {
            _app = ExcelSingleton.Instance;
            _wb = OpenWorkbook(filename);
        }

        private void TrackWorksheet(Excel.Worksheet w)
        {
            if (!_wss.Contains(w))
            {
                _wss.Add(w);
            }
        }

        private void FastReadAll(CellType ct)
        {
            var wbpath = FSharpOption<string>.Some(Path.GetDirectoryName(_wb.FullName));
            var wbname = FSharpOption<string>.Some(_wb.Name);

            foreach (Excel.Worksheet ws in _wb.Worksheets)
            {
                // keep track of this worksheet
                TrackWorksheet(ws);

                // get worksheet name
                var wsname = FSharpOption<string>.Some(ws.Name);

                // get used range
                Excel.Range ur = ws.UsedRange;

                // sometimes the used range is null
                if (ur.Value2 != null) { continue; }

                // calculate offsets
                var left = ur.Column;
                var right = ur.Columns.Count + left - 1;
                var top = ur.Row;
                var bottom = ur.Rows.Count + top - 1;

                // adjust offsets for Excel 1-based addressing
                var x_del = left - 1;
                var y_del = top - 1;

                // sometimes the Used Range is a range
                if (left != right || top != bottom)
                {
                    // y is the first index
                    // x is the second index
                    object[,] buf2d = ur.Value2;        // fast array read for data

                    // copy cells in data to Cell objects
                    for (int i = 1; i <= buf2d.GetLength(0); i++)
                    {
                        for (int j = 1; j <= buf2d.GetLength(1); j++)
                        {
                            if (buf2d[i, j] != null)
                            {
                                // calculate address
                                var addr = AST.Address.NewFromR1C1(i + y_del, j + x_del, wsname, wbname, wbpath);

                                // data case
                                if (ct == CellType.Data)
                                {
                                    InsertValue(addr, System.Convert.ToString(buf2d[i, j]));
                                }
                                // formula case
                                else if (!String.IsNullOrWhiteSpace((String)buf2d[i, j])
                                         && fpatt.IsMatch((String)buf2d[i,j]))
                                {
                                    InsertFormula(addr, System.Convert.ToString(buf2d[i, j]));
                                }
                            }
                        }
                    }
                }
                // and other times it is a single cell
                else
                {
                    // calculate address
                    var addr = AST.Address.NewFromR1C1(top, left, wsname, wbname, wbpath);

                    // value2
                    var v2 = System.Convert.ToString(ur.Value2);

                    // data case
                    if (ct == CellType.Data)
                    {
                        InsertValue(addr, v2);
                    }
                    // formula case
                    else if (!String.IsNullOrWhiteSpace(v2)
                             && fpatt.IsMatch(v2))
                    {
                        InsertFormula(addr, v2);
                    }
                }
            }
        }

        private Excel.Workbook OpenWorkbook(string filename)
        {
            // we need to disable all alerts, e.g., password prompts, etc.
            _app.DisplayAlerts = false;

            // disable macros
            _app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // This call is stupid. See:
            // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
            _app.Workbooks.Open(filename, // FileName (String)
                                XlUpdateLinks.Yes, // UpdateLinks (XlUpdateLinks enum)
                                true, // ReadOnly (Boolean)
                                Missing.Value, // Format (int?)
                                "thisisnotapassword", // Password (String)
                                Missing.Value, // WriteResPassword (String)
                                true, // IgnoreReadOnlyRecommended (Boolean)
                                Missing.Value, // Origin (XlPlatform enum)
                                Missing.Value, // Delimiter; if the filetype is txt (String)
                                Missing.Value, // Editable; not what you think (Boolean)
                                false, // Notify (Boolean)
                                Missing.Value, // Converter(int)
                                false, // AddToMru (Boolean)
                                Missing.Value, // Local; really "use my locale?" (Boolean)
                                XlCorruptLoad.RepairFile); // CorruptLoad (XlCorruptLoad enum)

            return _app.Workbooks[1];
        }

        public void Dispose()
        {
            // release each worksheet COM object
            foreach (Excel.Worksheet w in _wss)
            {
                Marshal.ReleaseComObject(w);
            }

            // nullify worksheet collection
            _wss = null;

            // close workbook
            _wb.Close();

            // release COM object
            Marshal.ReleaseComObject(_wb);

            // nullify ref
            _wb = null;

            // poke GC
            GC.Collect();
        }

        public void InsertValue(AST.Address address, string value)
        {
            throw new NotImplementedException();
        }

        public string GetValue(AST.Address address)
        {
            throw new NotImplementedException();
        }

        public void InsertFormula(AST.Address address, string formula)
        {
            throw new NotImplementedException();
        }

        public string GetFormula(AST.Address address)
        {
            throw new NotImplementedException();
        }

        public bool IsFormula(AST.Address address)
        {
            throw new NotImplementedException();
        }
    }

    public class ExcelSingleton
    {
        private static Excel.Application _instance;
        private ExcelSingleton() { }
        public static Excel.Application Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new Excel.Application();
                }
                return _instance;
            }
        }

        ~ExcelSingleton()
        {
            _instance.Quit();
            _instance = null;
            GC.Collect();
        }
    }
}
