﻿using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.FSharp.Core;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using System.Threading;
using Addr = SpreadsheetAST.Address;
using Expr = SpreadsheetAST.Expression;

// Make all internal methods in this assembly visible to test code
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("NoSheetTests")]

namespace NoSheet
{
    public class InvalidRangeException : Exception
    {
        public InvalidRangeException(string message) : base(message) { }
    }
    
    public class ExcelSpreadsheet : ISpreadsheet, IDisposable
    {
        // reader-writer lock object
        private object _lock = new object();

        // interlock for safe disposal
        private int _disposed = 0;

        // COM handles
        private Excel.Application _app;
        private Excel.Workbook _wb;
        private Dictionary<string, Excel.Worksheet> _wss = new Dictionary<string, Excel.Worksheet>();

        // data storage
        private Dictionary<Addr, string> _data = new Dictionary<Addr, string>();
        private Dictionary<Addr, Expr> _formulas = new Dictionary<Addr, Expr>();
        private Dictionary<Addr, string> _formula_strings = new Dictionary<Addr, string>();
        private Graph.DirectedAcyclicGraph _graph;

        // init dirty bits (key is A1 worksheet name)
        private HashSet<Addr> _pending_writes = new HashSet<Addr>();
        private Dictionary<string, bool> _needs_data_read = new Dictionary<string, bool>();
        private Dictionary<string, bool> _needs_formula_read = new Dictionary<string, bool>();
        // TODO: At the moment, there is no support for writing formulas

        // formula string regex
        private readonly Regex ISFORMULA = new Regex("^=", RegexOptions.Compiled);

        // names
        private string _workbook_directory;
        private string _workbook_name;

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

        public enum FileFormat
        {
            AddIn = Excel.XlFileFormat.xlAddIn,
            AddIn8 = Excel.XlFileFormat.xlAddIn8,
            CSV = Excel.XlFileFormat.xlCSV,
            CSV_Mac = Excel.XlFileFormat.xlCSVMac,
            CSV_MSDOS = Excel.XlFileFormat.xlCSVMSDOS,
            CSV_Windows = Excel.XlFileFormat.xlCSVWindows,
            CurrentPlatformText = Excel.XlFileFormat.xlCurrentPlatformText,
            DBF2 = Excel.XlFileFormat.xlDBF2,
            DBF3 = Excel.XlFileFormat.xlDBF3,
            DBF4 = Excel.XlFileFormat.xlDBF4,
            DIF = Excel.XlFileFormat.xlDIF,
            Excel12 = Excel.XlFileFormat.xlExcel12,
            Excel2 = Excel.XlFileFormat.xlExcel2,
            Excel2FarEast = Excel.XlFileFormat.xlExcel2FarEast,
            Excel3 = Excel.XlFileFormat.xlExcel3,
            Excel4 = Excel.XlFileFormat.xlExcel4,
            Excel4Workbook = Excel.XlFileFormat.xlExcel4Workbook,
            Excel5 = Excel.XlFileFormat.xlExcel5,
            Excel7 = Excel.XlFileFormat.xlExcel7,
            Excel8 = Excel.XlFileFormat.xlExcel8,
            Excel9795 = Excel.XlFileFormat.xlExcel9795,
            HTML = Excel.XlFileFormat.xlHtml,
            IntlAddIn = Excel.XlFileFormat.xlIntlAddIn,
            IntlMacro = Excel.XlFileFormat.xlIntlMacro,
            OpenDocumentSpreadsheet = Excel.XlFileFormat.xlOpenDocumentSpreadsheet,
            OpenXMLAddIn = Excel.XlFileFormat.xlOpenXMLAddIn,
            OpenXMLTemplate = Excel.XlFileFormat.xlOpenXMLTemplate,
            OpenXMLTemplaceMacroEnabled = Excel.XlFileFormat.xlOpenXMLTemplateMacroEnabled,
            OpenXMLWorkbook = Excel.XlFileFormat.xlOpenXMLWorkbook,
            OpenXMLWorkbookMacroEnabled = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            SYLK = Excel.XlFileFormat.xlSYLK,
            Template = Excel.XlFileFormat.xlTemplate,
            Template8 = Excel.XlFileFormat.xlTemplate8,
            TextMac = Excel.XlFileFormat.xlTextMac,
            TextMSDOS = Excel.XlFileFormat.xlTextMSDOS,
            TextPrinter = Excel.XlFileFormat.xlTextPrinter,
            TextWindows = Excel.XlFileFormat.xlTextWindows,
            UnicodeText = Excel.XlFileFormat.xlUnicodeText,
            WebArchive = Excel.XlFileFormat.xlWebArchive,
            WJ2WD1 = Excel.XlFileFormat.xlWJ2WD1,
            WJ3 = Excel.XlFileFormat.xlWJ3,
            WJ3FJ3 = Excel.XlFileFormat.xlWJ3FJ3,
            WK1 = Excel.XlFileFormat.xlWK1,
            WK1ALL = Excel.XlFileFormat.xlWK1ALL,
            WK1FMT = Excel.XlFileFormat.xlWK1FMT,
            WK3 = Excel.XlFileFormat.xlWK3,
            WK3FM3 = Excel.XlFileFormat.xlWK3FM3,
            WK4 = Excel.XlFileFormat.xlWK4,
            WKS = Excel.XlFileFormat.xlWKS,
            WorkbookDefault = Excel.XlFileFormat.xlWorkbookDefault,
            WorkbookNormal = Excel.XlFileFormat.xlWorkbookNormal,
            Works2FarEast = Excel.XlFileFormat.xlWorks2FarEast,
            WQ1 = Excel.XlFileFormat.xlWQ1,
            XMLSpreadsheet = Excel.XlFileFormat.xlXMLSpreadsheet
        }

        /// <summary>
        /// Constructor. Initializes an ExcelSpreadsheet, starting up an Excel instance if necessary.
        /// ExcelSpreadsheet reads values from the backing store lazily, but when a request
        /// is required, it reads all worksheets at once in order to amortize the cost.
        /// </summary>
        /// <param name="filename"></param>
        public ExcelSpreadsheet(string filename)
        {
            // init Excel resources
            _app = ExcelSingleton.Instance;
            _wb = OpenWorkbook(filename);

            // grab names, which only change when we call Save As.
            _workbook_directory = Path.GetDirectoryName(_wb.FullName);
            _workbook_name = _wb.Name;

            // get worksheets
            foreach (Excel.Worksheet ws in _wb.Worksheets)
            {
                TrackWorksheet(ws);
            }

            // force a full Excel recalculation to ensure
            // that no stale values are used.
            _app.CalculateFullRebuild();
        }

        /// <summary>
        /// Returns the PID of the singleton Excel Application instance.
        /// </summary>
        /// <returns></returns>
        internal int GetProcessID()
        {
            return ExcelSingleton.ProcessID;
        }

        /// <summary>
        /// Internal use only. Not thread-safe. Register COM worksheet name
        /// and dirty bits. The dirty read bit
        /// is initialized to true since the worksheet needs to be read
        /// and the dirty write bit is initialized to false since nothing
        /// should be read just yet. Should only be called by constructor.
        /// </summary>
        /// <param name="w">Reference to an Excel COM Worksheet object.</param>
        private void TrackWorksheet(Excel.Worksheet w)
        {
            if (!_wss.ContainsKey(w.Name))
            {
                _wss.Add(w.Name, w);
                _needs_data_read.Add(w.Name, true);
                _needs_formula_read.Add(w.Name, true);
            }
        }

        /// <summary>
        /// Internal use only. Reads multiple values from the spreadsheet backing store.
        /// Not thread-safe! Should only ever be called by FastRead.
        /// </summary>
        /// <param name="celltype"></param>
        /// <param name="usedrange"></param>
        /// <param name="x_del"></param>
        /// <param name="y_del"></param>
        /// <param name="wsname"></param>
        /// <param name="wbname"></param>
        /// <param name="wbpath"></param>
        private void __ArrayRead(CellType celltype,
                                 Excel.Range usedrange,
                                 int x_del,
                                 int y_del,
                                 FSharpOption<string> wsname,
                                 FSharpOption<string> wbname,
                                 FSharpOption<string> wbpath)
        {
            // do fast read
            // y is the first index
            // x is the second index
            object[,] buf2d = celltype == CellType.Data ? usedrange.Value2 : usedrange.Formula;

            // calculate height and width once
            int height = buf2d.GetLength(0);
            int width = buf2d.GetLength(1);

            // copy cells in data to Cell objects
            for (int i = 1; i <= height; i++)
            {
                for (int j = 1; j <= width; j++)
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
                            // note that we ignore write signal
                            // on fast read since we only read
                            // on initial open and after writes
                            __CacheValue(addr, v);
                        }
                        // formula case
                        else
                        {
                            __CacheFormula(addr, System.Convert.ToString(buf2d[i, j]));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Internal use only. Reads a single value from the spreadsheet backing store.
        /// Not thread-safe! Should only ever be called by FastRead.
        /// </summary>
        /// <param name="celltype"></param>
        /// <param name="cell"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="wsname"></param>
        /// <param name="wbname"></param>
        /// <param name="wbpath"></param>
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
                // note that we ignore write signal
                // when doing a fast read
                __CacheValue(addr, v2);
            }
            // formula case
            else
            {
                __CacheFormula(addr, v2);
            }
        }

        /// <summary>
        /// Internal use only. Updates a single value in the spreadsheet backing store.
        /// Not thread-safe! Should only ever be called by FastUpdate.
        /// </summary>
        /// <param name="addr"></param>
        private void __CellUpdate(Addr addr)
        {
            // get COM object
            Excel.Range cell = GetCOMCell(addr);

            // write value
            cell.Value2 = _data[addr];
        }

        /// <summary>
        /// Internal use only. Updates multiple values in the spreadsheet backing store.
        /// Not thread-safe! Should only ever be called by FastUpdate.
        /// </summary>
        /// <param name="addrs"></param>
        private void __RegionUpdate(IEnumerable<Addr> addrs)
        {
            // get the smallest region that includes all of our updates
            SpreadsheetAST.Range region = GetRegion(addrs);

            // calculate deltas to adjust addresses
            // fo region bounds
            int x_del = region.getXLeft();
            int y_del = region.getYTop();

            // get corresponding COM object
            Excel.Range rng = GetCOMRange(region);

            // save all of the original values
            object[,] data = rng.Value2;

            // fill with data
            foreach (Addr addr in addrs)
            {
                var value = _data[addr];

                // calculate addresses for offset and
                // 1-based addressing
                var y = addr.Y - y_del + 1;
                var x = addr.X - x_del + 1;

                // update array
                data[y, x] = value;
            }

            // write data to COM object
            rng.Value2 = data;

            // find formulas that we may have overwritten
            var oform = _formula_strings.Where(pair => region.ContainsAddress(pair.Key));

            // fix each one
            foreach (KeyValuePair<Addr, string> fa in oform)
            {
                var addr = fa.Key;
                var value = fa.Value;

                GetCOMCell(addr).Formula = value;
            }
        }

        /// <summary>
        /// Flushes all cached data marked as changed to
        /// the backing Excel file.  Unsets dirty write bits and
        /// sets dirty read bits. This method should only ever
        /// be called by FastRead(). Not thread-safe.
        /// </summary>
        private void WriteToBackingStore()
        {
            // quick check for pending writes
            if (_pending_writes.Count == 0)
            {
                return;
            }

            // iterate through the worksheets
            foreach (KeyValuePair<string, Excel.Worksheet> kvp in _wss)
            {
                string wsname = kvp.Key;
                Excel.Workbook wb = kvp.Value.Parent;
                string wbname = wb.Name;
                string wbpath = Path.GetDirectoryName(wb.FullName);

                // filter pending writes to include only addresses for this worksheet
                var pw_filt =
                    _pending_writes.Where(addr =>
                        addr.A1Worksheet().Equals(wsname) &&
                        addr.A1Workbook().Equals(wbname) &&
                        addr.A1Path().Equals(wbpath)
                    );

                // move on if there are no applicable Addresses for this worksheet
                if (pw_filt.Count() == 0) { continue; }

                // if the value is a singleton, then don't do a range
                // write; Excel will throw a runtime exception
                if (pw_filt.Count() == 1)
                {
                    __CellUpdate(pw_filt.First());
                }
                // otherwise do range write
                else
                {
                    __RegionUpdate(pw_filt);
                }

                // set dirty read bit
                _needs_data_read[wsname] = true;
            }

            // ensure that dirty read bit is set for all
            // worksheets containing the outputs
            // of the inputs just written
            foreach (Addr a in _pending_writes)
            {
                foreach (Addr faddr in _graph.GetOutputDependencies(a))
                {
                    _needs_data_read[faddr.A1Worksheet()] = true;
                }
            }

            // clear pending writes
            _pending_writes.Clear();
        }

        /// <summary>
        /// Returns a Region given a collection of Addresses. Thread-safe.
        /// </summary>
        /// <param name="addresses"></param>
        /// <returns></returns>
        private static SpreadsheetAST.Range GetRegion(IEnumerable<Addr> addresses)
        {
            if (addresses.Count() == 0)
            {
                throw new Exception("IEnumerable must contain at least one Address.");
            }

            int leftmost = Int32.MaxValue;
            int rightmost = Int32.MinValue;
            int topmost = Int32.MaxValue;
            int bottommost = Int32.MinValue;

            FSharpOption<string> wsname = addresses.First().WorksheetName;
            FSharpOption<string> wbname = addresses.First().WorkbookName;
            FSharpOption<string> wbpath = addresses.First().Path;

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
        /// Not thread-safe.
        /// </summary>
        /// <param name="ct"></param>
        private void ReadFromBackingStore(CellType ct)
        {
            // always start by ensuring that pending writes have been flushed
            WriteToBackingStore();

            // quick check to see if we need to read at all
            if (ct == CellType.Data &&
                !_needs_data_read.ContainsValue(true)) { return; }
            if (ct == CellType.Formula &&
                !_needs_formula_read.ContainsValue(true)) { return; }

            var wbpath_o = FSharpOption<string>.Some(_workbook_directory);
            var wbname_o = FSharpOption<string>.Some(_workbook_name);

            foreach (KeyValuePair<string,Excel.Worksheet> pair in _wss)
            {
                // get worksheet name
                var wsname = pair.Key;
                var wsname_o = FSharpOption<string>.Some(wsname);

                // only read sheet if dirty bit is set
                // the above call to TrackWorksheet will
                // init the dirty bit to true on inital read
                if (ct == CellType.Data)
                {
                    if (!_needs_data_read[wsname]) { continue; }
                }
                else
                {
                    if (!_needs_formula_read[wsname]) { continue; }
                }

                // get used range
                Excel.Range ur = pair.Value.UsedRange;

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
                if (ct == CellType.Data)
                {
                    _needs_data_read[wsname] = false;
                }
                else
                {
                    _needs_formula_read[wsname] = false;
                }
            }

            // if we just reread formulas, we need to rebuild the graph
            if (ct == CellType.Formula)
            {
                _graph = new Graph.DirectedAcyclicGraph(_formulas, _data);
            }
        }

        /// <summary>
        /// For internal use only.  Opens an Excel Workbook given a path string
        /// and returns a COM Interop Excel.Workbook handle.  Not thread-safe.
        /// Should only be called by the ExcelSpreadsheet constructor.
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private Excel.Workbook OpenWorkbook(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException(filename);
            }

            // we need to disable all alerts, e.g., password prompts, etc.
            _app.DisplayAlerts = false;

            // disable macros
            _app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // This call is stupid. See:
            // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
            _app.Workbooks.Open(filename,                   // FileName (String)
                                XlUpdateLinks.Yes,          // UpdateLinks (XlUpdateLinks enum)
                                true,                       // ReadOnly (Boolean)
                                Missing.Value,              // Format (int?)
                                "thisisnotapassword",       // Password (String)
                                Missing.Value,              // WriteResPassword (String)
                                true,                       // IgnoreReadOnlyRecommended (Boolean)
                                Missing.Value,              // Origin (XlPlatform enum)
                                Missing.Value,              // Delimiter; if the filetype is txt (String)
                                Missing.Value,              // Editable; not what you think (Boolean)
                                false,                      // Notify (Boolean)
                                Missing.Value,              // Converter(int)
                                false,                      // AddToMru (Boolean)
                                Missing.Value,              // Local; really "use my locale?" (Boolean)
                                XlCorruptLoad.RepairFile);  // CorruptLoad (XlCorruptLoad enum)

            return _app.Workbooks[1];
        }

        /// <summary>
        /// For internal use only.  Use breaks thread-safe abstraction provided by
        /// ExcelSpreadsheet API.
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        internal Excel.Range GetCOMCell(Addr address)
        {
            var cell_ws = address.A1Worksheet();
            return _wss[cell_ws].get_Range(address.A1Local());
        }

        /// <summary>
        /// For internal use only.  Use breaks thread-safe abstraction provided by
        /// ExcelSpreadsheet API.
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        internal Excel.Range GetCOMRange(SpreadsheetAST.Range range)
        {
            var tla = range.TopLeftAddress();
            var bra = range.BottomRightAddress();

            return _wss[range.GetWorksheetName()].get_Range(tla.A1Local(), bra.A1Local());
        }

        /// <summary>
        /// For internal use only.  Returns the Address for a given COM object. Thread-safe.
        /// </summary>
        /// <param name="com"></param>
        /// <param name="wb"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Internal use only.  Parses a formula and caches the expression
        /// tree.  Should only ever be called by __ArrayRead or __CellRead.
        /// Not thread-safe!
        /// </summary>
        /// <param name="address"></param>
        /// <param name="value"></param>
        private bool __CacheValue(Addr address, string value)
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
                    // return without signaling write
                    return false;
                } else {
                    _data[address] = value;
                }
            }
            else
            {
                // only add if not null or empty
                if (String.IsNullOrEmpty(value))
                {
                    // return without signaling write
                    return false;
                }
                else
                {
                    _data.Add(address, value);
                }
            }

            return true;
        }

        public void SetValueAt(Addr address, string value)
        {
            lock (_lock)
            {
                // if a dependency graph has never been built,
                // now is the time to do so.
                ReadFromBackingStore(CellType.Formula);

                // we do not write to formula outputs
                if (_formulas.ContainsKey(address))
                {
                    throw new FormulaOverwriteException(address);
                }

                if (__CacheValue(address, value))
                {
                    _pending_writes.Add(address);
                }
            }
        }

        public string ValueAt(Addr address)
        {
            lock (_lock)
            {
                ReadFromBackingStore(CellType.Data);

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
        }

        public Expr FormulaAt(Addr address)
        {
            lock (_lock)
            {
                ReadFromBackingStore(CellType.Formula);

                Expr f;
                if (_formulas.TryGetValue(address, out f))
                {
                    return f;
                }
                else
                {
                    throw new IsNotFormulaException(address);
                }
            }
        }

        /// <summary>
        /// Internal use only.  Parses a formula and caches the expression
        /// tree.  Should only ever be called by __ArrayRead or __CellRead.
        /// Not thread-safe!
        /// </summary>
        /// <param name="address"></param>
        /// <param name="formula"></param>
        private void __CacheFormula(Addr address, string formula)
        {
            if (!String.IsNullOrWhiteSpace(formula)
                && ISFORMULA.IsMatch(formula))
            {
                // parse formula
                var pf = ExcelParserUtility.ParseFormulaWithAddress(formula, address);

                // insert
                if (_formulas.ContainsKey(address))
                {
                    _formulas[address] = pf;
                    _formula_strings[address] = formula;
                }
                else
                {
                    _formulas.Add(address, pf);
                    _formula_strings.Add(address, formula);
                }
            }
        }

        public string FormulaAsStringAt(Addr address)
        {
            lock (_lock)
            {
                ReadFromBackingStore(CellType.Formula);

                // TODO: this should actually use an Excel-specific
                //       visitor, since SpreadsheetAST is supposed
                //       to be a spreadsheet-agnostic IR
                string f;
                if (_formula_strings.TryGetValue(address, out f))
                {
                    return f;
                }
                else
                {
                    throw new IsNotFormulaException(address);
                }
            }
        }

        public bool IsFormulaAt(Addr address)
        {
            lock (_lock)
            {
                ReadFromBackingStore(CellType.Formula);

                return _formulas.ContainsKey(address);

            }
        }

        public Dictionary<Addr, string> Values
        {
            get {
                lock (_lock)
                {
                    ReadFromBackingStore(CellType.Data);

                    return _data;
                }
            }
        }

        public Dictionary<Addr, Expr> Formulas
        {
            get {
                lock (_lock)
                {
                    ReadFromBackingStore(CellType.Formula);

                    return _formulas;
                }
            }
        }

        /// <summary>
        /// Save changes to the backing file.
        /// </summary>
        public void Save()
        {
            lock (_lock)
            {
                ReadFromBackingStore(CellType.Formula);
                ReadFromBackingStore(CellType.Data);

                _wb.Save();
            }
        }

        /// <summary>
        /// Save the file with a different filename. If the
        /// file already exists, SaveAs returns false and saves nothing.
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public bool SaveAs(string filename)
        {
            return SaveAs(filename, FileFormat.WorkbookDefault);
        }

        /// <summary>
        /// Save the file with a different filename and/or file format. If the
        /// file already exists, SaveAs returns false and saves nothing.
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="fileformat"></param>
        /// <returns></returns>
        public bool SaveAs(string filename, FileFormat fileformat)
        {
            lock (_lock) {
            ReadFromBackingStore(CellType.Formula);
            ReadFromBackingStore(CellType.Data);

                if (File.Exists(filename))
                {
                    return false;
                }

                _wb.SaveAs(filename,                                            // filename
                           fileformat,                                          // FileFormat enum
                           Type.Missing,                                        // password
                           Type.Missing,                                        // write reservation password
                           false,                                               // readonly recommended
                           false,                                               // create backup
                           Excel.XlSaveAsAccessMode.xlExclusive,                // access mode
                           Excel.XlSaveConflictResolution.xlLocalSessionChanges,// conflict resolution policy
                           false,                                               // add to MRU list
                           Type.Missing,                                        // codepage (ignored)
                           Type.Missing,                                        // visual layout (ignored)
                           true                                                 // true == "Excel language"; false == "VBA language"
                          );

                // when someone changes the name of the workbook, we need to reread the
                // cells so that our addresses are modified appropriately
                foreach (KeyValuePair<string, bool> pair in _needs_data_read)
                {
                    _needs_data_read[pair.Key] = true;
                }
                foreach (KeyValuePair<string, bool> pair in _needs_formula_read)
                {
                    _needs_formula_read[pair.Key] = true;
                }
                _workbook_directory = Path.GetDirectoryName(_wb.FullName);
                _workbook_name = _wb.Name;
            }

            return true;
        }

        /// <summary>
        /// Returns an array of the worksheets contained in the spreadsheet.
        /// </summary>
        /// <returns></returns>
        public string[] WorksheetNames
        {
            get {
                lock(_lock) {
                    return _wss.Select(pair => pair.Key).ToArray();
                }
            }
        }
        
        /// <summary>
        /// Returns the name of the workbook represented by this spreadsheet.
        /// </summary>
        public string WorkbookName
        {
            get {
                lock (_lock) {
                    return _workbook_name;
                }
            }
        }

        /// <summary>
        /// Returns the directory of the current spreadsheet file.
        /// </summary>
        public string Directory
        {
            get
            {
                lock (_lock) {
                    return _workbook_directory;
                }
            }
        }

        /// <summary>
        /// Returns an array of homogeneous input ranges.
        /// </summary>
        public SpreadsheetAST.Range[] HomogeneousInputs
        {
            get
            {
                // read formulas if they need to be updated
                ReadFromBackingStore(CellType.Formula);
                return _graph.HomogeneousInputs;
            }
        }

        public void Dispose()
        {
            // don't call this more than once
            if (Interlocked.Increment(ref _disposed) == 1)
            {
                // we will block here indefinitely until all readers
                // and writers release their locks
                lock (_lock)
                {
                    // release each worksheet COM object
                    foreach (KeyValuePair<string, Excel.Worksheet> pair in _wss)
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
            }
        }
    }

    public class ExcelSingleton
    {
        // P/Invoke call to get PID
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowThreadProcessId(HandleRef handle, out int processId);

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
        public static int ProcessID
        {
            get
            {
                // force singleton startup
                var i = ExcelSingleton.Instance;

                // get PID
                HandleRef hwnd = new HandleRef(_instance, (IntPtr)_instance.Hwnd);
                int pid;
                GetWindowThreadProcessId(hwnd, out pid);
                return pid;
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
