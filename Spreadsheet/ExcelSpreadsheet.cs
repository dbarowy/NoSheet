using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.FSharp.Core;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;

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

        // init dirty bits (key is A1 worksheet name)
        private Dictionary<string, bool> _dirty_sheets = new Dictionary<string, bool>();

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
            // init Excel resources
            _app = ExcelSingleton.Instance;
            _wb = OpenWorkbook(filename);

            // do initial reads
            FastReadAll(CellType.Data);
            FastReadAll(CellType.Formula);
        }

        private void TrackWorksheet(Excel.Worksheet w)
        {
            if (!_wss.Contains(w))
            {
                _wss.Add(w);
                _dirty_sheets.Add(w.Name, true);
            }
        }

        private void __ArrayRead(CellType celltype,
                                 Excel.Range usedrange,
                                 int x_del,
                                 int y_del,
                                 FSharpOption<string> wsname,
                                 FSharpOption<string> wbname,
                                 FSharpOption<string> wbpath)
        {
            // y is the first index
            // x is the second index
            object[,] buf2d = usedrange.Value2; // do read

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
                        if (celltype == CellType.Data)
                        {
                            InsertValue(addr, System.Convert.ToString(buf2d[i, j]));
                        }
                        // formula case
                        else if (!String.IsNullOrWhiteSpace((String)buf2d[i, j])
                                 && fpatt.IsMatch((String)buf2d[i, j]))
                        {
                            InsertFormula(addr, System.Convert.ToString(buf2d[i, j]));
                        }
                    }
                }
            }
        }

        private void __CellRead(CellType celltype,
                                Excel.Range cell,
                                int left,
                                int top,
                                FSharpOption<string> wsname,
                                FSharpOption<string> wbname,
                                FSharpOption<string> wbpath)
        {
            // calculate address
            var addr = AST.Address.NewFromR1C1(top, left, wsname, wbname, wbpath);

            // value2
            var v2 = System.Convert.ToString(cell.Value2);

            // data case
            if (celltype == CellType.Data)
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

        private void FastReadAll(CellType ct)
        {
            var wbpath_o = FSharpOption<string>.Some(Path.GetDirectoryName(_wb.FullName));
            var wbname_o = FSharpOption<string>.Some(_wb.Name);

            foreach (Excel.Worksheet ws in _wb.Worksheets)
            {
                // keep track of this worksheet
                TrackWorksheet(ws);

                // only read sheet if dirty bit is set
                // the above call to TrackWorksheet will
                // init the dirty bit to true
                if (!_dirty_sheets[ws.Name]) { continue; }

                // get used range
                Excel.Range ur = ws.UsedRange;

                // sometimes the used range is null
                if (ur.Value2 != null) { continue; }

                // get worksheet name
                var wsname = ws.Name;
                var wsname_o = FSharpOption<string>.Some(wsname);

                // calculate offsets
                var left = ur.Column;
                var right = ur.Columns.Count + left - 1;
                var top = ur.Row;
                var bottom = ur.Rows.Count + top - 1;

                // sometimes the Used Range is a range
                if (left != right || top != bottom)
                {
                    // adjust offsets for Excel 1-based addressing
                    var x_del = left - 1;
                    var y_del = top - 1;

                    __ArrayRead(ct, ur, x_del, y_del, wsname_o, wbname_o, wbpath_o);
                }
                // and other times it is a single cell
                else
                {
                    __CellRead(ct, ur, left, top, wsname_o, wbname_o, wbpath_o);
                }

                // unset dirty bit
                _dirty_sheets[wsname] = false;
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
            // insert into local storage
            if (_data.ContainsKey(address))
            {
                _data[address] = value;
            }
            else
            {
                _data.Add(address, value);
            }

            // set dirty bit for worksheet
            _dirty_sheets[address.A1Worksheet()] = true;
        }

        public string GetValue(AST.Address address)
        {
            if (_dirty_sheets.ContainsValue(true))
            {
                FastReadAll(CellType.Data);
            }
            string value;
            if (_data.TryGetValue(address, out value))
            {
                return value;
            }
            else
            {
                return String.Empty;
            }
        }

        public void InsertFormula(AST.Address address, string formula)
        {
            throw new NotImplementedException();
        }

        public string GetFormula(AST.Address address)
        {
            string formula;
            if (_formulas.TryGetValue(address, out formula))
            {
                return formula;
            }
            else
            {
                return String.Empty;
            }
        }

        public bool IsFormula(AST.Address address)
        {
            return _formulas.ContainsKey(address);
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
