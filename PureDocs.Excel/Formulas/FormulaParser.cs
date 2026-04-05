namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Recursive descent parser that converts formula tokens into an AST.
/// Follows Excel operator precedence:
///   1. Percent (%) — postfix
///   2. Power (^) — right-associative
///   3. Negation (-) — unary prefix
///   4. Multiplication, Division (*, /)
///   5. Addition, Subtraction (+, -)
///   6. Concatenation (&amp;)
///   7. Comparison (=, &lt;&gt;, &lt;, &lt;=, &gt;, &gt;=)
/// </summary>
internal sealed class FormulaParser
{
    private readonly List<FormulaToken> _tokens;
    private int _pos;

    public FormulaParser(List<FormulaToken> tokens)
    {
        _tokens = tokens ?? throw new ArgumentNullException(nameof(tokens));
        _pos = 0;
    }

    /// <summary>
    /// Parses the token stream into an AST.
    /// </summary>
    public FormulaNode Parse()
    {
        _pos = 0;
        var node = ParseComparison();

        if (Current.Type != FormulaTokenType.EOF)
            throw new FormulaException($"Unexpected token '{Current.Value}' at position {Current.Position}");

        return node;
    }

    // ── Token helpers ──────────────────────────────────────────────────

    private FormulaToken Current => _pos < _tokens.Count
        ? _tokens[_pos]
        : new FormulaToken(FormulaTokenType.EOF, "", -1);

    private FormulaToken Advance()
    {
        var token = Current;
        _pos++;
        return token;
    }

    private FormulaToken Expect(FormulaTokenType type)
    {
        if (Current.Type != type)
            throw new FormulaException(
                $"Expected {type} but got {Current.Type} ('{Current.Value}') at position {Current.Position}");
        return Advance();
    }

    // ── Precedence levels ──────────────────────────────────────────────

    /// <summary>Comparison: =, &lt;&gt;, &lt;, &lt;=, &gt;, &gt;=</summary>
    private FormulaNode ParseComparison()
    {
        var left = ParseConcatenation();

        while (true)
        {
            BinaryOperator? op = Current.Type switch
            {
                FormulaTokenType.Equal => BinaryOperator.Equal,
                FormulaTokenType.NotEqual => BinaryOperator.NotEqual,
                FormulaTokenType.LessThan => BinaryOperator.LessThan,
                FormulaTokenType.LessThanOrEqual => BinaryOperator.LessThanOrEqual,
                FormulaTokenType.GreaterThan => BinaryOperator.GreaterThan,
                FormulaTokenType.GreaterThanOrEqual => BinaryOperator.GreaterThanOrEqual,
                _ => null
            };

            if (op == null) break;
            Advance();
            var right = ParseConcatenation();
            left = new BinaryOpNode(left, op.Value, right);
        }

        return left;
    }

    /// <summary>Concatenation: &amp;</summary>
    private FormulaNode ParseConcatenation()
    {
        var left = ParseAddSub();

        while (Current.Type == FormulaTokenType.Ampersand)
        {
            Advance();
            var right = ParseAddSub();
            left = new BinaryOpNode(left, BinaryOperator.Concatenate, right);
        }

        return left;
    }

    /// <summary>Addition/Subtraction: +, -</summary>
    private FormulaNode ParseAddSub()
    {
        var left = ParseMulDiv();

        while (true)
        {
            BinaryOperator? op = Current.Type switch
            {
                FormulaTokenType.Plus => BinaryOperator.Add,
                FormulaTokenType.Minus => BinaryOperator.Subtract,
                _ => null
            };

            if (op == null) break;
            Advance();
            var right = ParseMulDiv();
            left = new BinaryOpNode(left, op.Value, right);
        }

        return left;
    }

    /// <summary>Multiplication/Division: *, /</summary>
    private FormulaNode ParseMulDiv()
    {
        var left = ParseUnary();

        while (true)
        {
            BinaryOperator? op = Current.Type switch
            {
                FormulaTokenType.Multiply => BinaryOperator.Multiply,
                FormulaTokenType.Divide => BinaryOperator.Divide,
                _ => null
            };

            if (op == null) break;
            Advance();
            var right = ParseUnary();
            left = new BinaryOpNode(left, op.Value, right);
        }

        return left;
    }

    /// <summary>Unary: +, -, @ (prefix)</summary>
    private FormulaNode ParseUnary()
    {
        if (Current.Type == FormulaTokenType.Minus)
        {
            Advance();
            var operand = ParsePower();
            return new UnaryOpNode(UnaryOperator.Negate, operand);
        }

        if (Current.Type == FormulaTokenType.Plus)
        {
            Advance();
            var operand = ParsePower();
            return new UnaryOpNode(UnaryOperator.Plus, operand);
        }

        if (Current.Type == FormulaTokenType.AtSign)
        {
            Advance();
            var operand = ParsePower();
            return new ImplicitIntersectionNode(operand);
        }

        return ParsePower();
    }

    /// <summary>Power: ^ (right-associative)</summary>
    private FormulaNode ParsePower()
    {
        var left = ParsePercent();

        if (Current.Type == FormulaTokenType.Power)
        {
            Advance();
            var right = ParseUnary(); // right-associative
            return new BinaryOpNode(left, BinaryOperator.Power, right);
        }

        return left;
    }

    /// <summary>Percent: % (postfix)</summary>
    private FormulaNode ParsePercent()
    {
        var node = ParsePrimary();

        while (Current.Type == FormulaTokenType.Percent)
        {
            Advance();
            node = new UnaryOpNode(UnaryOperator.Percent, node);
        }

        return node;
    }

    /// <summary>Primary: literals, cell refs, ranges, functions, parenthesized exprs</summary>
    private FormulaNode ParsePrimary()
    {
        var token = Current;

        switch (token.Type)
        {
            case FormulaTokenType.Number:
                Advance();
                if (!double.TryParse(token.Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double num))
                    throw new FormulaException($"Invalid number: {token.Value}");
                return new NumberNode(num);

            case FormulaTokenType.String:
                Advance();
                return new StringNode(token.Value);

            case FormulaTokenType.Boolean:
                Advance();
                return new BooleanNode(token.Value == "TRUE");

            case FormulaTokenType.ErrorLiteral:
                Advance();
                var error = FormulaValue.ErrorFromString(token.Value);
                return new ErrorNode(error);

            case FormulaTokenType.CellReference:
                Advance();
                // Check for range (A1:B5)
                if (Current.Type == FormulaTokenType.Colon)
                {
                    Advance();
                    var endRef = Expect(FormulaTokenType.CellReference);
                    return new RangeReferenceNode(token.Value, endRef.Value);
                }
                return new CellReferenceNode(token.Value);

            case FormulaTokenType.SheetReference:
                Advance();
                // value is "SheetName!CellRef" — split and create SheetReferenceNode
                var sheetParts = token.Value.Split('!', 2);
                string sheetRefCell = sheetParts.Length > 1 ? sheetParts[1] : sheetParts[0];
                string sheetName = sheetParts.Length > 1 ? sheetParts[0] : "";
                // Check for range: Sheet1!A1:B5
                if (Current.Type == FormulaTokenType.Colon)
                {
                    Advance();
                    // After colon, expect cell reference (without sheet prefix)
                    var sheetRangeEnd = Expect(FormulaTokenType.CellReference);
                    return new SheetRangeReferenceNode(sheetName, sheetRefCell, sheetRangeEnd.Value);
                }
                return new SheetCellReferenceNode(sheetName, sheetRefCell);

            case FormulaTokenType.NamedRange:
                Advance();
                return new NamedRangeNode(token.Value);

            case FormulaTokenType.Function:
                return ParseFunction();

            case FormulaTokenType.LeftParen:
                Advance();
                var expr = ParseComparison();
                Expect(FormulaTokenType.RightParen);
                return expr;

            default:
                throw new FormulaException(
                    $"Unexpected token '{token.Value}' ({token.Type}) at position {token.Position}");
        }
    }

    /// <summary>Function call: NAME(arg1, arg2, ...) - supports empty arguments like IF(A1,,0)</summary>
    private FormulaNode ParseFunction()
    {
        var nameToken = Advance(); // function name
        Expect(FormulaTokenType.LeftParen);

        var args = new List<FormulaNode>();

        if (Current.Type != FormulaTokenType.RightParen)
        {
            // First argument - check if empty (starts with comma or is closing paren)
            if (Current.Type == FormulaTokenType.Comma)
            {
                args.Add(new BlankNode());
            }
            else
            {
                args.Add(ParseComparison());
            }

            while (Current.Type == FormulaTokenType.Comma)
            {
                Advance(); // consume comma
                
                // Check for empty argument: consecutive comma or closing paren
                if (Current.Type == FormulaTokenType.Comma || Current.Type == FormulaTokenType.RightParen)
                {
                    args.Add(new BlankNode());
                }
                else
                {
                    args.Add(ParseComparison());
                }
            }
        }

        Expect(FormulaTokenType.RightParen);
        return new FunctionCallNode(nameToken.Value, args);
    }
}
