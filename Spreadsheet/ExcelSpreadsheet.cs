using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.FSharp.Core;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using Addr = SpreadsheetAST.Address;
using Expr = SpreadsheetAST.Expression;

namespace NoSheet
{
    public class InvalidRangeException : Exception
    {
        public InvalidRangeException(string message) : base(message) { }
    }

    public class ExcelSpreadsheet : ISpreadsheet, IDisposable
    {
        // COM handles
        private Excel.Application _app;
        private Excel.Workbook _wb;
        private Dictionary<string, Excel.Worksheet> _wss = new Dictionary<string, Excel.Worksheet>();

        // data storage
        private Dictionary<Addr, string> _data = new Dictionary<Addr, string>();
        private Dictionary<Addr, Expr> _formulas = new Dictionary<Addr, Expr>();
        private Graph.DirectedAcyclicGraph _graph;

        // init dirty bits (key is A1 worksheet name)
        private Dictionary<string, bool> _needs_write = new Dictionary<string, bool>();
        private Dictionary<string, bool> _needs_read = new Dictionary<string, bool>();

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
            FastRead(CellType.Data);
            FastRead(CellType.Formula);

            // construct DAG
            _graph = new Graph.DirectedAcyclicGraph(_formulas, _data);
        }

        private void TrackWorksheet(Excel.Worksheet w)
        {
            if (!_wss.ContainsKey(w.Name))
            {
                _wss.Add(w.Name, w);
                _needs_read.Add(w.Name, true);
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
                        var addr = Addr.NewFromR1C1(i + y_del, j + x_del, wsname, wbname, wbpath);

                        // read value
                        var v = System.Convert.ToString(buf2d[i, j]);

                        // data case
                        if (celltype == CellType.Data)
                        {
                            if (String.IsNullOrEmpty(v))
                            {
                                InsertValue(addr, String.Empty);
                            }
                            else
                            {
                                InsertValue(addr, v);
                            }
                        }
                        // formula case
                        else if (!String.IsNullOrWhiteSpace((String)buf2d[i, j])
                                 && fpatt.IsMatch((String)buf2d[i, j]))
                        {
                            InsertFormulaAsString(addr, System.Convert.ToString(buf2d[i, j]));
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
            var addr = Addr.NewFromR1C1(top, left, wsname, wbname, wbpath);

            // value2
            var v2 = System.Convert.ToString(cell.Value2);

            // data case
            if (celltype == CellType.Data)
            {
                if (String.IsNullOrEmpty(v2))
                {
                    InsertValue(addr, String.Empty);
                }
                else
                {
                    InsertValue(addr, v2);
                }
            }
            // formula case
            else if (!String.IsNullOrWhiteSpace(v2)
                     && fpatt.IsMatch(v2))
            {
                InsertFormulaAsString(addr, v2);
            }
        }

        /// <summary>
        /// Fluses all cached data with dirty worksheet bits to
        /// the backing Excel file.  Unsets dirty write bits and
        /// sets dirty read bits.
        /// </summary>
        private void FastWrite()
        {
            // worksheets with dirty write
            var ds = _needs_write.Where(pair => pair.Value == true).Select(pair => pair.Key);

            // changed input list
            var changed_addrs = new List<Addr>();

            foreach(string wsname in ds)
            {
                // filter dictionary to include only matching addresses
                IEnumerable<KeyValuePair<Addr,string>> fdata = _data.Where(pair => pair.Key.A1Worksheet().Equals(wsname));

                // collect addresses
                IEnumerable<Addr> addrl = fdata.Select(pair => pair.Key);

                // get used range
                SpreadsheetAST.Range ur = GetRegion(addrl);

                // calculate deltas to store in zero-based array
                int x_del = -ur.getXLeft();
                int y_del = -ur.getYTop();

                // construct write object
                // first coord is y
                // second coord is x
                string[,] output = new string[ur.Height, ur.Width];

                // fill with data
                foreach (KeyValuePair<Addr, string> pair in fdata)
                {
                    var addr = pair.Key;
                    var value = pair.Value;

                    // write
                    output[addr.Y + y_del, addr.X + x_del] = value;

                    // note changed address
                    changed_addrs.Add(addr);
                }

                // get corresponding COM object
                Excel.Range rng = GetCOMRange(ur);

                // write to COM object
                rng.Value2 = output;

                // unset dirty write bit
                _needs_write[wsname] = false;

                // set dirty read bit
                _needs_read[wsname] = true;
            }

            // ensure that dirty read bit is set for all
            // outputs of the inputs just written
            foreach (Addr a in changed_addrs)
            {
                foreach (Addr faddr in _graph.GetOutputDependencies(a))
                {
                    _needs_read[faddr.A1Worksheet()] = true;
                }
            }
        }

        private SpreadsheetAST.Range GetRegion(IEnumerable<Addr> addresses)
        {
            int leftmost = 0;
            int rightmost = 0;
            int topmost = 0;
            int bottommost = 0;

            FSharpOption<string> wsname = FSharpOption<string>.None;
            FSharpOption<string> wbname = FSharpOption<string>.None;
            FSharpOption<string> wbpath = FSharpOption<string>.None;

            foreach (Addr a in addresses)
            {
                if (FSharpOption<string>.get_IsNone(wsname))
                {
                    wsname = a.WorksheetName;
                }
                else if (!wsname.Equals(a.WorksheetName))
                {
                    throw new InvalidRangeException(
                        String.Format(
                            "Range contains references to worksheets \"{0}\" and \"{1}\"",
                            wsname.Value,
                            a.WorksheetName.Value
                        )
                    );
                }

                if (a.X < leftmost)
                {
                    leftmost = a.X;
                }
                if (a.X > rightmost)
                {
                    rightmost = a.X;
                }
                if (a.Y < topmost)
                {
                    topmost = a.Y;
                }
                if (a.Y > bottommost)
                {
                    bottommost = a.Y;
                }
            }

            // get topleft and bottomright address
            Addr tl = Addr.NewFromR1C1(topmost, leftmost, wsname, wbname, wbpath);
            Addr br = Addr.NewFromR1C1(bottommost, rightmost, wsname, wbname, wbpath);

            // return corresponding range
            return new SpreadsheetAST.Range(tl, br);
        }

        /// <summary>
        /// Reads all values whose worksheets have dirty read bits set to true,
        /// or all values if worksheet has never been read.  Unsets dirty read bits.
        /// </summary>
        /// <param name="ct"></param>
        private void FastRead(CellType ct)
        {
            var wbpath_o = FSharpOption<string>.Some(Path.GetDirectoryName(_wb.FullName));
            var wbname_o = FSharpOption<string>.Some(_wb.Name);

            foreach (Excel.Worksheet ws in _wb.Worksheets)
            {
                // keep track of this worksheet
                TrackWorksheet(ws);

                // get worksheet name
                var wsname = ws.Name;
                var wsname_o = FSharpOption<string>.Some(wsname);

                // only read sheet if dirty bit is set
                // the above call to TrackWorksheet will
                // init the dirty bit to true
                if (!_needs_read[wsname]) { continue; }

                // get used range
                Excel.Range ur = ws.UsedRange;

                // sometimes the used range is null
                if (ur.Value2 != null) {
                    // unset dirty bit then continue
                    _needs_read[wsname] = false;
                    continue;
                }

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

                // unset needs read bit
                _needs_read[wsname] = false;
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

        private Excel.Range GetCOMCell(Addr address)
        {
            var cell_ws = address.A1Worksheet();
            return _wss[cell_ws].get_Range(address.A1Local());
        }

        private Excel.Range GetCOMRange(SpreadsheetAST.Range range)
        {
            return _wss[range.GetWorksheetName()].get_Range(range.TopLeftAddress(), range.BottomRightAddress());
        }

        private static Addr AddressFromCOMObject(Excel.Range com, Excel.Workbook wb) {
            var wsname = com.Worksheet.Name;
            var wbname = wb.Name;
            var path = System.IO.Path.GetDirectoryName(wb.FullName);
            var addr = com.get_Address(true,
                                       true,
                                       Excel.XlReferenceStyle.xlR1C1,
                                       Type.Missing,
                                       Type.Missing);
             return Addr.FromString(addr,
                                           FSharpOption<string>.Some(wsname),
                                           FSharpOption<string>.Some(wbname),
                                           FSharpOption<string>.Some(path));
        }

        private IEnumerable<SpreadsheetAST.Range> GetReferencesFromFormula(string formula, Excel.Workbook wb, Excel.Worksheet ws)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            // release each worksheet COM object
            foreach (KeyValuePair<string,Excel.Worksheet> pair in _wss)
            {
                Marshal.ReleaseComObject(pair.Value);
            }

            // nullify worksheet collection
            _wss.Clear();
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

        public void InsertValue(Addr address, string value)
        {
            // insert into local storage
            if (_data.ContainsKey(address))
            {
                // if the string is null or empty,
                // remove from the dict because
                // the default value to return when
                // the key is not in the dictionary
                // is the empty string
                if (String.IsNullOrEmpty(value))
                {
                    _data.Remove(address);
                }
                // don't bother updating if the values
                // are the same
                else if (_data[address].Equals(value))
                {
                    // return without setting dirty write bit
                    return;
                } else {
                    _data[address] = value;
                }
            }
            else
            {
                _data.Add(address, value);
            }

            // set dirty write bit for worksheet
            _needs_write[address.A1Worksheet()] = true;
        }

        public string GetValue(Addr address)
        {
            if (_needs_write.ContainsValue(true))
            {
                FastWrite();
                FastRead(CellType.Data);
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

        public void InsertFormula(Addr address, Expr formula)
        {
            throw new NotImplementedException();
        }

        public Expr GetFormula(Addr address)
        {
            throw new NotImplementedException();
        }

        public void InsertFormulaAsString(Addr address, string formula)
        {
            // parse formula
            var pf = ExcelParserUtility.ParseFormulaWithAddress(formula, address);

            // insert
            if (_formulas.ContainsKey(address))
            {
                _formulas[address] = pf;
            }
            else
            {
                _formulas.Add(address, pf);
            }
        }

        public string GetFormulaAsString(Addr address)
        {
            // TODO: this should actually use an Excel-specific
            //       visitor, since SpreadsheetAST is supposed
            //       to be a spreadsheet-agnostic IR
            return _formulas[address].ToString();
        }

        public bool IsFormula(Addr address)
        {
            return _formulas.ContainsKey(address);
        }

        public Dictionary<Addr, string> GetAllValues()
        {
            return _data;
        }
        public Dictionary<Addr, Expr> GetAllFormulas()
        {
            return _formulas;
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
