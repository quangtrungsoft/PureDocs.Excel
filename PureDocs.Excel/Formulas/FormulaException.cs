namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Exception thrown for formula parsing or evaluation errors. 
/// These correspond to Excel error codes like #VALUE!, #REF!, #DIV/0!, etc.
/// </summary>
public sealed class FormulaException : Exception
{
    public FormulaException(string message) : base(message) { }
    public FormulaException(string message, Exception innerException) : base(message, innerException) { }
}
