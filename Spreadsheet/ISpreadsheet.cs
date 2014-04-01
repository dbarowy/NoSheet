namespace NoSheet
{
    public interface ISpreadsheet
    {
        public void InsertValue(AST.Address address, string value);
        public string GetValue(AST.Address address);
        public string GetFormula(AST.Address address);
        public bool IsFormula(AST.Address address);
    }    
}
