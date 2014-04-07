using System.Collections.Generic;

namespace NoSheet
{
    public interface ISpreadsheet
    {
        SpreadsheetAST.Expression GetFormula(SpreadsheetAST.Address address);
        Dictionary<SpreadsheetAST.Address, string> GetAllValues();
        Dictionary<SpreadsheetAST.Address, SpreadsheetAST.Expression> GetAllFormulas();
        string GetFormulaAsString(SpreadsheetAST.Address address);
        string GetValue(SpreadsheetAST.Address address);
        void InsertFormula(SpreadsheetAST.Address address, SpreadsheetAST.Expression ast);
        void InsertFormulaAsString(SpreadsheetAST.Address address, string formula);
        void InsertValue(SpreadsheetAST.Address address, string value);
        bool IsFormula(SpreadsheetAST.Address address);
    }    
}
