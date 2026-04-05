using System.Text;

namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Tokenizes an Excel formula string into a sequence of tokens.
/// Supports: sheet-prefixed refs (Sheet1!A1, 'My Sheet'!A1), _xlfn. prefix,
/// named ranges, quoted sheet names, standard operators, error literals (#N/A, #REF!, etc.).
/// </summary>
/// <remarks>
/// <para>
/// Known limitations:
/// - Array constants {1,2,3} are not supported. Formulas like =MATCH(1,{1,2,3},0) will fail.
/// - Consider implementing array constant support in a future version.
/// </para>
/// </remarks>
internal sealed class FormulaLexer
{
    private readonly string _formula;
    private int _pos;

    public FormulaLexer(string formula)
    {
        _formula = formula ?? throw new ArgumentNullException(nameof(formula));
        _pos = 0;
    }

    /// <summary>
    /// Tokenizes the entire formula.
    /// </summary>
    public List<FormulaToken> Tokenize()
    {
        var tokens = new List<FormulaToken>();
        _pos = 0;

        while (_pos < _formula.Length)
        {
            SkipWhitespace();
            if (_pos >= _formula.Length) break;

            var token = ReadNextToken(tokens);
            if (token != null)
                tokens.Add(token);
        }

        tokens.Add(new FormulaToken(FormulaTokenType.EOF, "", _pos));
        return tokens;
    }

    private FormulaToken? ReadNextToken(List<FormulaToken> previousTokens)
    {
        char c = _formula[_pos];

        // String literal
        if (c == '"')
            return ReadString();

        // Error literal: #N/A, #REF!, #VALUE!, etc.
        if (c == '#')
            return ReadErrorLiteral();

        // Quoted sheet name: 'Sheet Name'!A1
        if (c == '\'')
            return ReadQuotedSheetRef();

        // Number
        if (char.IsDigit(c) || (c == '.' && _pos + 1 < _formula.Length && char.IsDigit(_formula[_pos + 1])))
            return ReadNumber();

        // Identifier: cell reference, function name, boolean, named range, or sheet!ref
        if (char.IsLetter(c) || c == '_' || c == '$')
            return ReadIdentifier();

        // Operators and punctuation
        int startPos = _pos;
        switch (c)
        {
            case '(':
                _pos++;
                return new FormulaToken(FormulaTokenType.LeftParen, "(", startPos);
            case ')':
                _pos++;
                return new FormulaToken(FormulaTokenType.RightParen, ")", startPos);
            case ',':
                _pos++;
                return new FormulaToken(FormulaTokenType.Comma, ",", startPos);
            case ':':
                _pos++;
                return new FormulaToken(FormulaTokenType.Colon, ":", startPos);
            case '+':
                _pos++;
                return new FormulaToken(FormulaTokenType.Plus, "+", startPos);
            case '-':
                _pos++;
                return new FormulaToken(FormulaTokenType.Minus, "-", startPos);
            case '*':
                _pos++;
                return new FormulaToken(FormulaTokenType.Multiply, "*", startPos);
            case '/':
                _pos++;
                return new FormulaToken(FormulaTokenType.Divide, "/", startPos);
            case '^':
                _pos++;
                return new FormulaToken(FormulaTokenType.Power, "^", startPos);
            case '%':
                _pos++;
                return new FormulaToken(FormulaTokenType.Percent, "%", startPos);
            case '&':
                _pos++;
                return new FormulaToken(FormulaTokenType.Ampersand, "&", startPos);
            case '=':
                _pos++;
                return new FormulaToken(FormulaTokenType.Equal, "=", startPos);
            case '<':
                _pos++;
                if (_pos < _formula.Length)
                {
                    if (_formula[_pos] == '>')
                    {
                        _pos++;
                        return new FormulaToken(FormulaTokenType.NotEqual, "<>", startPos);
                    }
                    if (_formula[_pos] == '=')
                    {
                        _pos++;
                        return new FormulaToken(FormulaTokenType.LessThanOrEqual, "<=", startPos);
                    }
                }
                return new FormulaToken(FormulaTokenType.LessThan, "<", startPos);
            case '>':
                _pos++;
                if (_pos < _formula.Length && _formula[_pos] == '=')
                {
                    _pos++;
                    return new FormulaToken(FormulaTokenType.GreaterThanOrEqual, ">=", startPos);
                }
                return new FormulaToken(FormulaTokenType.GreaterThan, ">", startPos);
            case '!':
                _pos++;
                return new FormulaToken(FormulaTokenType.Exclamation, "!", startPos);
            case '@':
                _pos++;
                return new FormulaToken(FormulaTokenType.AtSign, "@", startPos);
            default:
                throw new FormulaException($"Unexpected character '{c}' at position {_pos}");
        }
    }

    /// <summary>Reads error literals: #N/A, #REF!, #VALUE!, #DIV/0!, #NAME?, #NUM!, #NULL!</summary>
    private FormulaToken ReadErrorLiteral()
    {
        int start = _pos;
        var sb = new StringBuilder();
        sb.Append(_formula[_pos]); // #
        _pos++;

        // Read until we hit a terminating character
        // Error literals end with ! or ? (except #N/A)
        while (_pos < _formula.Length)
        {
            char c = _formula[_pos];
            if (char.IsLetterOrDigit(c) || c == '/' || c == '!' || c == '?')
            {
                sb.Append(c);
                _pos++;
                // ! and ? are terminators for most errors
                if (c == '!' || c == '?')
                    break;
            }
            else
            {
                break;
            }
        }

        string errorValue = sb.ToString().ToUpperInvariant();
        
        // Validate it's a known error
        if (IsKnownError(errorValue))
            return new FormulaToken(FormulaTokenType.ErrorLiteral, errorValue, start);
        
        throw new FormulaException($"Unknown error literal '{errorValue}' at position {start}");
    }

    private static bool IsKnownError(string s) => s switch
    {
        "#NULL!" or "#DIV/0!" or "#VALUE!" or "#REF!" or 
        "#NAME?" or "#NUM!" or "#N/A" or "#CALC!" or "#SPILL!" => true,
        _ => false
    };

    /// <summary>Reads 'Sheet Name'!A1 or 'Sheet Name'!A1:B5 style references.</summary>
    private FormulaToken ReadQuotedSheetRef()
    {
        int start = _pos;
        _pos++; // skip opening '
        var sb = new StringBuilder();

        while (_pos < _formula.Length)
        {
            char c = _formula[_pos];
            if (c == '\'')
            {
                _pos++;
                // Escaped '' inside quoted name
                if (_pos < _formula.Length && _formula[_pos] == '\'')
                {
                    sb.Append('\'');
                    _pos++;
                }
                else
                {
                    // End of quoted name — expect !
                    if (_pos < _formula.Length && _formula[_pos] == '!')
                    {
                        _pos++; // skip !
                        string sheetName = sb.ToString();

                        // Read the cell reference after !
                        var cellRefSb = new StringBuilder();
                        while (_pos < _formula.Length)
                        {
                            char rc = _formula[_pos];
                            if (char.IsLetterOrDigit(rc) || rc == '$')
                            { cellRefSb.Append(rc); _pos++; }
                            else break;
                        }

                        string cellRef = cellRefSb.ToString().Replace("$", "");
                        
                        // Check for range: 'Sheet Name'!A1:B5
                        // Don't consume the colon - let parser handle range construction
                        // The colon will be tokenized separately and parser will build the range
                        
                        return new FormulaToken(FormulaTokenType.SheetReference,
                            sheetName + "!" + cellRef, start);
                    }
                    throw new FormulaException($"Expected '!' after quoted sheet name at position {start}");
                }
            }
            else
            {
                sb.Append(c);
                _pos++;
            }
        }

        throw new FormulaException($"Unterminated quoted sheet name starting at position {start}");
    }

    private FormulaToken ReadString()
    {
        int start = _pos;
        _pos++; // skip opening quote
        var sb = new StringBuilder();

        while (_pos < _formula.Length)
        {
            char c = _formula[_pos];
            if (c == '"')
            {
                _pos++;
                // Excel uses "" for escaped quote
                if (_pos < _formula.Length && _formula[_pos] == '"')
                {
                    sb.Append('"');
                    _pos++;
                }
                else
                {
                    return new FormulaToken(FormulaTokenType.String, sb.ToString(), start);
                }
            }
            else
            {
                sb.Append(c);
                _pos++;
            }
        }

        throw new FormulaException($"Unterminated string starting at position {start}");
    }

    private FormulaToken ReadNumber()
    {
        int start = _pos;
        var sb = new StringBuilder();
        bool hasDot = false;
        bool hasE = false;

        while (_pos < _formula.Length)
        {
            char c = _formula[_pos];

            if (char.IsDigit(c))
            {
                sb.Append(c);
                _pos++;
            }
            else if (c == '.' && !hasDot && !hasE)
            {
                hasDot = true;
                sb.Append(c);
                _pos++;
            }
            else if ((c == 'E' || c == 'e') && !hasE && sb.Length > 0)
            {
                hasE = true;
                sb.Append(c);
                _pos++;
                // optional sign after E
                if (_pos < _formula.Length && (_formula[_pos] == '+' || _formula[_pos] == '-'))
                {
                    sb.Append(_formula[_pos]);
                    _pos++;
                }
            }
            else
            {
                break;
            }
        }

        return new FormulaToken(FormulaTokenType.Number, sb.ToString(), start);
    }

    private FormulaToken ReadIdentifier()
    {
        int start = _pos;
        var sb = new StringBuilder();

        // Read the identifier (letters, digits, $, _, .)
        while (_pos < _formula.Length)
        {
            char c = _formula[_pos];
            if (char.IsLetterOrDigit(c) || c == '_' || c == '$' || c == '.')
            {
                sb.Append(c);
                _pos++;
            }
            else
            {
                break;
            }
        }

        string value = sb.ToString();
        string upper = value.ToUpperInvariant();

        // Strip _xlfn. prefix (future function prefix in Excel)
        if (upper.StartsWith("_XLFN."))
        {
            value = value[6..];
            upper = upper[6..];
        }

        // Check for booleans
        if (upper == "TRUE")
            return new FormulaToken(FormulaTokenType.Boolean, "TRUE", start);
        if (upper == "FALSE")
            return new FormulaToken(FormulaTokenType.Boolean, "FALSE", start);

        // Check for sheet reference: Sheet1!A1
        if (_pos < _formula.Length && _formula[_pos] == '!')
        {
            string sheetName = value;
            _pos++; // skip !

            // Read cell reference after !
            var cellRefSb = new StringBuilder();
            while (_pos < _formula.Length)
            {
                char c = _formula[_pos];
                if (char.IsLetterOrDigit(c) || c == '$')
                { cellRefSb.Append(c); _pos++; }
                else break;
            }
            string cellRef = cellRefSb.ToString().Replace("$", "");
            return new FormulaToken(FormulaTokenType.SheetReference,
                sheetName + "!" + cellRef, start);
        }

        // Check if it's a function call (identifier followed by open paren)
        SkipWhitespace();
        if (_pos < _formula.Length && _formula[_pos] == '(')
            return new FormulaToken(FormulaTokenType.Function, upper, start);

        // Otherwise it's a cell reference or named range
        // If it looks like a cell reference (letter(s) + digit(s)), return CellReference
        string stripped = value.Replace("$", "");
        if (IsCellReference(stripped))
            return new FormulaToken(FormulaTokenType.CellReference, stripped, start);

        // Otherwise return as NamedRange token
        return new FormulaToken(FormulaTokenType.NamedRange, value, start);
    }

    /// <summary>
    /// Checks if a string looks like a cell reference (e.g., A1, AB123).
    /// </summary>
    /// <remarks>
    /// This check prioritizes cell references over named ranges when there's
    /// a conflict. For example, "AB12" looks like both a valid cell reference
    /// and a potential named range. Excel uses the same priority: cell refs win.
    /// Users should avoid naming ranges with patterns that match cell refs.
    /// </remarks>
    private static bool IsCellReference(string s)
    {
        if (s.Length < 2) return false;
        int i = 0;
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == 0 || i >= s.Length) return false;
        while (i < s.Length && char.IsDigit(s[i])) i++;
        return i == s.Length;
    }

    private void SkipWhitespace()
    {
        while (_pos < _formula.Length && char.IsWhiteSpace(_formula[_pos]))
            _pos++;
    }
}
