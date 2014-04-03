namespace NoSheet
{
    public interface ISpreadsheet
    {
        void InsertValue(SpreadsheetAST.Address address, string value);
        string GetValue(SpreadsheetAST.Address address);
        void InsertFormula(SpreadsheetAST.Address address, string formula);
        string GetFormula(SpreadsheetAST.Address address);
        bool IsFormula(SpreadsheetAST.Address address);
    }    
}
