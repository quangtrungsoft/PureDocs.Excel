namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Represents a token produced by the formula lexer.
/// </summary>
internal sealed class FormulaToken
{
    public FormulaTokenType Type { get; }
    public string Value { get; }
    public int Position { get; }

    public FormulaToken(FormulaTokenType type, string value, int position)
    {
        Type = type;
        Value = value;
        Position = position;
    }

    public override string ToString() => $"[{Type}] '{Value}' @{Position}";
}

/// <summary>
/// Types of tokens in an Excel formula.
/// </summary>
internal enum FormulaTokenType
{
    /// <summary>Numeric literal (integer or decimal)</summary>
    Number,

    /// <summary>String literal enclosed in double quotes</summary>
    String,

    /// <summary>Boolean literal (TRUE/FALSE)</summary>
    Boolean,

    /// <summary>Cell reference (e.g., A1, $B$2)</summary>
    CellReference,

    /// <summary>Range operator (:) — separates two cell references</summary>
    Colon,

    /// <summary>Function name (e.g., SUM, VLOOKUP)</summary>
    Function,

    /// <summary>Opening parenthesis</summary>
    LeftParen,

    /// <summary>Closing parenthesis</summary>
    RightParen,

    /// <summary>Comma (argument separator)</summary>
    Comma,

    /// <summary>Plus operator</summary>
    Plus,

    /// <summary>Minus operator</summary>
    Minus,

    /// <summary>Multiply operator</summary>
    Multiply,

    /// <summary>Divide operator</summary>
    Divide,

    /// <summary>Power operator (^)</summary>
    Power,

    /// <summary>Percent operator (%)</summary>
    Percent,

    /// <summary>Concatenation operator (&amp;)</summary>
    Ampersand,

    /// <summary>Equal comparison (=)</summary>
    Equal,

    /// <summary>Not equal comparison (&lt;&gt;)</summary>
    NotEqual,

    /// <summary>Less than (&lt;)</summary>
    LessThan,

    /// <summary>Less than or equal (&lt;=)</summary>
    LessThanOrEqual,

    /// <summary>Greater than (&gt;)</summary>
    GreaterThan,

    /// <summary>Greater than or equal (&gt;=)</summary>
    GreaterThanOrEqual,

    /// <summary>Sheet-prefixed reference (e.g., Sheet1!A1, 'My Sheet'!B2)</summary>
    SheetReference,

    /// <summary>Named range identifier (e.g., SalesTotal)</summary>
    NamedRange,

    /// <summary>Exclamation mark (!) — used in sheet references</summary>
    Exclamation,

    /// <summary>Structured reference: Table1[Column1]</summary>
    StructuredReference,

    /// <summary>3D sheet range reference: Sheet1:Sheet3</summary>
    ThreeDSheetRange,

    /// <summary>@ implicit intersection operator</summary>
    AtSign,

    /// <summary>Error literal (e.g., #N/A, #REF!, #VALUE!)</summary>
    ErrorLiteral,

    /// <summary>End of formula</summary>
    EOF
}
