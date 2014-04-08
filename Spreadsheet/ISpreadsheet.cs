using System.Collections.Generic;
using SpreadsheetAST;

namespace NoSheet
{
    public interface ISpreadsheet
    {
        // Properties
        string                          Directory { get; }
        Dictionary<Address, Expression> Formulas { get; }
        Dictionary<Address, string>     Values { get; }
        string                          WorkbookName { get; }
        string[]                        WorksheetNames { get; }

        // Methods
        string      FormulaAsStringAt(Address address);
        Expression  FormulaAt(Address address);
        bool        IsFormulaAt(Address address);
        void        Save();
        bool        SaveAs(string filename);
        void        SetValueAt(Address address, string value);
        string      ValueAt(Address address);
    }
}
