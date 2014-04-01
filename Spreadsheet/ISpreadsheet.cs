namespace NoSheet
{
    public interface ISpreadsheet
    {
        void InsertValue(AST.Address address, string value);
        string GetValue(AST.Address address);
        void InsertFormula(AST.Address address, string formula);
        string GetFormula(AST.Address address);
        bool IsFormula(AST.Address address);
    }    
}
