using System;
using System.Collections.Generic;
using SpreadsheetAST;

namespace NoSheet
{
    public class FormulaOverwriteException : Exception
    {
        public FormulaOverwriteException(Address address)
            : base(String.Format("Can't overwrite formula output at {0}.", address.A1FullyQualified())) { }
    }

    public class IsNotFormulaException : Exception
    {
        public IsNotFormulaException(Address address)
            : base(String.Format("Can't get formula for non-formula cell at {0}.", address.A1FullyQualified())) { }
    }

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
