namespace TVE.PureDocs.Excel.Formulas;

// ── AST Nodes ──────────────────────────────────────────────────────

internal abstract class FormulaNode
{
    public abstract FormulaValue Evaluate(FormulaContext context);
}

internal sealed class NumberNode(double value) : FormulaNode
{
    public double Value => value;
    public override FormulaValue Evaluate(FormulaContext context) => FormulaValue.Number(value);
}

internal sealed class StringNode(string value) : FormulaNode
{
    public string Value => value;
    public override FormulaValue Evaluate(FormulaContext context) => FormulaValue.Text(value);
}

internal sealed class BooleanNode(bool value) : FormulaNode
{
    public bool Value => value;
    public override FormulaValue Evaluate(FormulaContext context) => FormulaValue.Boolean(value);
}

/// <summary>Error literal node: #N/A, #REF!, #VALUE!, etc.</summary>
internal sealed class ErrorNode(FormulaError error) : FormulaNode
{
    public FormulaError Error => error;
    public override FormulaValue Evaluate(FormulaContext context) => FormulaValue.Error(error);
}

/// <summary>Blank/empty node for omitted function arguments (e.g., IF(A1,,0)).</summary>
internal sealed class BlankNode : FormulaNode
{
    public override FormulaValue Evaluate(FormulaContext context) => FormulaValue.Blank;
}

internal sealed class CellReferenceNode(string reference) : FormulaNode
{
    public string Reference => reference;
    public override FormulaValue Evaluate(FormulaContext context) => context.GetCellValue(reference);
}

internal sealed class RangeReferenceNode(string startRef, string endRef) : FormulaNode
{
    public string StartRef => startRef;
    public string EndRef => endRef;
    public override FormulaValue Evaluate(FormulaContext context) => context.GetRangeValues(startRef, endRef);
}

internal sealed class FunctionCallNode(string functionName, List<FormulaNode> arguments) : FormulaNode
{
    public string FunctionName => functionName;
    public List<FormulaNode> Arguments => arguments;
    public override FormulaValue Evaluate(FormulaContext context) => context.EvaluateFunction(functionName, arguments);
}

/// <summary>Sheet-prefixed cell reference: Sheet1!A1</summary>
internal sealed class SheetCellReferenceNode(string sheetName, string cellReference) : FormulaNode
{
    public string SheetName => sheetName;
    public string CellReference => cellReference;
    public override FormulaValue Evaluate(FormulaContext context)
        => context.GetSheetCellValue(sheetName, cellReference);
}

/// <summary>Sheet-prefixed range reference: Sheet1!A1:B5</summary>
internal sealed class SheetRangeReferenceNode(string sheetName, string startRef, string endRef) : FormulaNode
{
    public string SheetName => sheetName;
    public string StartRef => startRef;
    public string EndRef => endRef;
    public override FormulaValue Evaluate(FormulaContext context)
        => context.GetSheetRangeValues(sheetName, startRef, endRef);
}

/// <summary>Named range reference: SalesTotal</summary>
internal sealed class NamedRangeNode(string name) : FormulaNode
{
    public string Name => name;
    public override FormulaValue Evaluate(FormulaContext context)
        => context.ResolveNamedRange(name);
}

/// <summary>Structured reference: Table1[Column1]</summary>
internal sealed class StructuredReferenceNode(string tableName, string specifier) : FormulaNode
{
    public string TableName => tableName;
    public string Specifier => specifier;
    public override FormulaValue Evaluate(FormulaContext context)
        => context.ResolveStructuredReference(tableName, specifier);
}

/// <summary>3D cell reference: Sheet1:Sheet3!A1</summary>
internal sealed class ThreeDCellReferenceNode(string startSheet, string endSheet, string cellReference) : FormulaNode
{
    public string StartSheet => startSheet;
    public string EndSheet => endSheet;
    public string CellReference => cellReference;
    public override FormulaValue Evaluate(FormulaContext context)
        => context.Resolve3DCell(startSheet, endSheet, cellReference);
}

/// <summary>3D range reference: Sheet1:Sheet3!A1:B5</summary>
internal sealed class ThreeDRangeReferenceNode(string startSheet, string endSheet, string startRef, string endRef) : FormulaNode
{
    public string StartSheet => startSheet;
    public string EndSheet => endSheet;
    public string StartRef => startRef;
    public string EndRef => endRef;
    public override FormulaValue Evaluate(FormulaContext context)
        => context.Resolve3DRange(startSheet, endSheet, startRef, endRef);
}

/// <summary>
/// Implicit intersection operator (@). Reduces array to scalar based on formula position.
/// </summary>
/// <remarks>
/// For implicit intersection to work correctly with ranges, the range's starting position
/// must be known. This node attempts to extract range bounds from the inner node.
/// </remarks>
internal sealed class ImplicitIntersectionNode(FormulaNode inner) : FormulaNode
{
    public FormulaNode Inner => inner;

    public override FormulaValue Evaluate(FormulaContext context)
    {
        var val = inner.Evaluate(context);
        if (!val.IsArray) return val;
        
        // Try to get range bounds from inner node
        var (rangeStartRow, rangeStartCol) = GetRangeBounds(inner);
        
        return val.ArrayVal.ImplicitIntersect(
            context.FormulaRow, context.FormulaCol,
            rangeStartRow, rangeStartCol);
    }

    /// <summary>
    /// Attempts to extract range bounds from a node.
    /// Returns (0, 0) if bounds cannot be determined.
    /// </summary>
    private static (int startRow, int startCol) GetRangeBounds(FormulaNode node)
    {
        return node switch
        {
            RangeReferenceNode range => ParseCellRef(range.StartRef),
            SheetRangeReferenceNode sheetRange => ParseCellRef(sheetRange.StartRef),
            CellReferenceNode cell => ParseCellRef(cell.Reference),
            SheetCellReferenceNode sheetCell => ParseCellRef(sheetCell.CellReference),
            ThreeDRangeReferenceNode threeD => ParseCellRef(threeD.StartRef),
            ThreeDCellReferenceNode threeDCell => ParseCellRef(threeDCell.CellReference),
            // For other node types (e.g., function results), use 0,0 as default
            _ => (0, 0)
        };
    }

    private static (int row, int col) ParseCellRef(string cellRef)
    {
        try
        {
            CellReference.Parse(cellRef, out int row, out int col);
            return (row, col);
        }
        catch
        {
            return (0, 0);
        }
    }
}

internal sealed class BinaryOpNode(FormulaNode left, BinaryOperator op, FormulaNode right) : FormulaNode
{
    public FormulaNode Left => left;
    public FormulaNode Right => right;
    public BinaryOperator Operator => op;

    public override FormulaValue Evaluate(FormulaContext context)
    {
        var l = left.Evaluate(context);
        if (l.IsError) return l;

        var rv = right.Evaluate(context);
        if (rv.IsError) return rv;

        // ── Array Broadcasting ───────────────────────────────────────
        // If either operand is an array, broadcast the operation element-wise
        if (l.IsArray || rv.IsArray)
        {
            var leftArray = l.IsArray ? l.ArrayVal : ArrayValue.FromScalar(l);
            var rightArray = rv.IsArray ? rv.ArrayVal : ArrayValue.FromScalar(rv);

            return FormulaValue.Array(ArrayValue.Broadcast(leftArray, rightArray,
                (a, b) => EvaluateSingleOp(a, b, op)));
        }

        // ── Scalar Operation ─────────────────────────────────────────
        return EvaluateSingleOp(l, rv, op);
    }

    /// <summary>Evaluates binary operation on two scalar values.</summary>
    private static FormulaValue EvaluateSingleOp(FormulaValue l, FormulaValue rv, BinaryOperator op)
    {
        if (l.IsError) return l;
        if (rv.IsError) return rv;

        // Concatenation
        if (op == BinaryOperator.Concatenate)
            return FormulaValue.Text(l.AsText() + rv.AsText());

        // Comparison — uses Excel type ordering
        switch (op)
        {
            case BinaryOperator.Equal: return FormulaValue.Boolean(FormulaValue.AreEqual(l, rv));
            case BinaryOperator.NotEqual: return FormulaValue.Boolean(!FormulaValue.AreEqual(l, rv));
            case BinaryOperator.LessThan: return FormulaValue.Boolean(FormulaValue.Compare(l, rv) < 0);
            case BinaryOperator.LessThanOrEqual: return FormulaValue.Boolean(FormulaValue.Compare(l, rv) <= 0);
            case BinaryOperator.GreaterThan: return FormulaValue.Boolean(FormulaValue.Compare(l, rv) > 0);
            case BinaryOperator.GreaterThanOrEqual: return FormulaValue.Boolean(FormulaValue.Compare(l, rv) >= 0);
        }

        // Arithmetic — coerce to number
        var ln = l.CoerceToNumber();
        if (ln.IsError) return ln;
        var rn = rv.CoerceToNumber();
        if (rn.IsError) return rn;
        double a = ln.NumberValue, b = rn.NumberValue;

        return op switch
        {
            BinaryOperator.Add => FormulaValue.Number(a + b),
            BinaryOperator.Subtract => FormulaValue.Number(a - b),
            BinaryOperator.Multiply => FormulaValue.Number(a * b),
            BinaryOperator.Divide => b == 0 ? FormulaValue.ErrorDiv0 : FormulaValue.Number(a / b),
            BinaryOperator.Power => FormulaValue.Number(Math.Pow(a, b)),
            _ => FormulaValue.ErrorValue,
        };
    }
}

internal sealed class UnaryOpNode(UnaryOperator op, FormulaNode operand) : FormulaNode
{
    public FormulaNode Operand => operand;
    public UnaryOperator Operator => op;

    public override FormulaValue Evaluate(FormulaContext context)
    {
        var v = operand.Evaluate(context);
        if (v.IsError) return v;
        var n = v.CoerceToNumber();
        if (n.IsError) return n;
        double d = n.NumberValue;
        return op switch
        {
            UnaryOperator.Negate => FormulaValue.Number(-d),
            UnaryOperator.Plus => FormulaValue.Number(d),
            UnaryOperator.Percent => FormulaValue.Number(d / 100.0),
            _ => FormulaValue.ErrorValue,
        };
    }
}

// ── Enums ──────────────────────────────────────────────────────────

internal enum BinaryOperator
{
    Add, Subtract, Multiply, Divide, Power, Concatenate,
    Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual
}

internal enum UnaryOperator { Negate, Plus, Percent }

// ── FormulaHelper ──────────────────────────────────────────────────

internal static class FormulaHelper
{
    /// <summary>
    /// Maximum cells to materialize before using lazy iteration.
    /// Ranges larger than this will use LazyRangeIterator to avoid OOM.
    /// </summary>
    private const int LazyThreshold = 10_000;

    /// <summary>Collect numeric values from args. Returns true on success. Error propagates.</summary>
    /// <remarks>
    /// For large ranges (> LazyThreshold cells), uses LazyRangeIterator to avoid
    /// materializing the entire range into memory. This prevents OutOfMemoryException
    /// for formulas like SUM(A:A) which reference 1M+ cells.
    /// </remarks>
    public static bool TryCollectNumbers(List<FormulaNode> args, FormulaContext ctx,
        List<double> numbers, out FormulaValue error)
    {
        error = default;
        foreach (var arg in args)
        {
            // Check if this is a large range that should use lazy iteration
            if (TryCollectNumbersLazy(arg, ctx, numbers, ref error))
                continue;
                
            // Fall back to standard evaluation
            var val = arg.Evaluate(ctx);
            if (!AccumulateNumbers(val, numbers, true, ref error))
                return false;
        }
        return true;
    }

    /// <summary>
    /// Attempts to collect numbers lazily for large range references.
    /// Returns true if handled (either success or error), false to fall back to standard eval.
    /// </summary>
    private static bool TryCollectNumbersLazy(FormulaNode arg, FormulaContext ctx,
        List<double> numbers, ref FormulaValue error)
    {
        // Check for range reference nodes that might be large
        if (arg is RangeReferenceNode rangeNode)
        {
            CellReference.Parse(rangeNode.StartRef, out int sr, out int sc);
            CellReference.Parse(rangeNode.EndRef, out int er, out int ec);
            
            long cellCount = (long)(er - sr + 1) * (ec - sc + 1);
            if (cellCount > LazyThreshold)
            {
                // Use lazy iteration
                var iterator = new LazyRangeIterator(ctx.Worksheet, sr, sc, er, ec);
                foreach (var num in iterator.EnumerateNumbers())
                {
                    numbers.Add(num);
                }
                return true; // Handled
            }
        }
        else if (arg is SheetRangeReferenceNode sheetRangeNode)
        {
            CellReference.Parse(sheetRangeNode.StartRef, out int sr, out int sc);
            CellReference.Parse(sheetRangeNode.EndRef, out int er, out int ec);
            
            long cellCount = (long)(er - sr + 1) * (ec - sc + 1);
            if (cellCount > LazyThreshold)
            {
                // Get the target worksheet
                var workbook = ctx.Worksheet.GetWorkbook();
                if (workbook == null)
                {
                    error = FormulaValue.ErrorRef;
                    return true;
                }
                
                var targetSheet = workbook.Worksheets[sheetRangeNode.SheetName];
                if (targetSheet == null)
                {
                    error = FormulaValue.ErrorRef;
                    return true;
                }
                
                var iterator = new LazyRangeIterator(targetSheet, sr, sc, er, ec);
                foreach (var num in iterator.EnumerateNumbers())
                {
                    numbers.Add(num);
                }
                return true; // Handled
            }
        }
        
        return false; // Not handled, use standard evaluation
    }

    private static bool AccumulateNumbers(FormulaValue val, List<double> numbers,
        bool isDirect, ref FormulaValue error)
    {
        if (val.IsError) { error = val; return false; }
        if (val.IsArray)
        {
            foreach (var item in val.ArrayVal.EnumerateValues())
            {
                if (item.IsError) { error = item; return false; }
                if (item.IsNumber) numbers.Add(item.NumberValue);
                // Skip blanks, text, bools in ranges
            }
        }
        else if (val.IsNumber) numbers.Add(val.NumberValue);
        else if (val.IsBoolean && isDirect) numbers.Add(val.BooleanValue ? 1 : 0);
        else if (val.IsBlank) { /* skip */ }
        else if (val.IsText && isDirect)
        {
            if (double.TryParse(val.TextValue, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double d))
                numbers.Add(d);
            else { error = FormulaValue.ErrorValue; return false; }
        }
        return true;
    }

    /// <summary>Flatten args into individual FormulaValues (expand arrays).</summary>
    public static bool TryFlattenArgs(List<FormulaNode> args, FormulaContext ctx,
        List<FormulaValue> result, out FormulaValue error)
    {
        error = default;
        foreach (var arg in args)
        {
            var val = arg.Evaluate(ctx);
            if (val.IsError) { error = val; return false; }
            if (val.IsArray)
                foreach (var item in val.ArrayVal.EnumerateValues()) result.Add(item);
            else
                result.Add(val);
        }
        return true;
    }

    /// <summary>Evaluate single arg to array (range or single-value wrapped).</summary>
    public static FormulaValue EvalAsArray(FormulaNode node, FormulaContext ctx)
    {
        var val = node.Evaluate(ctx);
        if (val.IsArray || val.IsError) return val;
        var arr = new ArrayValue(1, 1);
        arr[0] = val;
        return FormulaValue.Array(arr);
    }

    /// <summary>Ensure minimum arg count.</summary>
    public static bool CheckArgCount(List<FormulaNode> args, int min, out FormulaValue error, int max = -1)
    {
        error = default;
        if (args.Count < min || (max >= 0 && args.Count > max))
        { error = FormulaValue.ErrorValue; return false; }
        return true;
    }

    /// <summary>Evaluate arg and coerce to double. Error-propagating.</summary>
    public static bool TryEvalDouble(FormulaNode node, FormulaContext ctx, out double result, out FormulaValue error)
    {
        error = default; result = 0;
        var v = node.Evaluate(ctx);
        if (v.IsError) { error = v; return false; }
        var n = v.CoerceToNumber();
        if (n.IsError) { error = n; return false; }
        result = n.NumberValue;
        return true;
    }

    /// <summary>Evaluate arg and get string. Error-propagating.</summary>
    public static bool TryEvalString(FormulaNode node, FormulaContext ctx, out string result, out FormulaValue error)
    {
        error = default; result = "";
        var v = node.Evaluate(ctx);
        if (v.IsError) { error = v; return false; }
        result = v.AsText();
        return true;
    }

    /// <summary>
    /// Matches value against criteria like ">=10", "&lt;&gt;Apple", "A*".
    /// Supports Excel tilde (~) escape: ~* = literal *, ~? = literal ?, ~~ = literal ~
    /// </summary>
    public static bool MatchesCriteria(FormulaValue value, string criteria)
    {
        if (string.IsNullOrEmpty(criteria)) return false;
        string op = "="; string cmp = criteria;
        if (criteria.StartsWith(">=")) { op = ">="; cmp = criteria[2..]; }
        else if (criteria.StartsWith("<=")) { op = "<="; cmp = criteria[2..]; }
        else if (criteria.StartsWith("<>")) { op = "<>"; cmp = criteria[2..]; }
        else if (criteria.StartsWith(">")) { op = ">"; cmp = criteria[1..]; }
        else if (criteria.StartsWith("<")) { op = "<"; cmp = criteria[1..]; }
        else if (criteria.StartsWith("=")) { op = "="; cmp = criteria[1..]; }

        // Numeric comparison
        if (double.TryParse(cmp, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double nc) && value.TryAsDouble(out double nv) && !value.IsText)
        {
            return op switch
            {
                "=" => Math.Abs(nv - nc) < 1e-10, "<>" => Math.Abs(nv - nc) >= 1e-10,
                ">" => nv > nc, ">=" => nv >= nc, "<" => nv < nc, "<=" => nv <= nc, _ => false,
            };
        }
        // Text comparison
        string tv = value.AsText();
        if (op is "=" or "<>")
        {
            // Check if pattern contains unescaped wildcards
            bool hasWildcard = HasUnescapedWildcard(cmp);
            bool m = hasWildcard ? WildcardMatch(tv, cmp)
                : string.Equals(tv, UnescapeTilde(cmp), StringComparison.OrdinalIgnoreCase);
            return op == "=" ? m : !m;
        }
        int c = string.Compare(tv, cmp, StringComparison.OrdinalIgnoreCase);
        return op switch { ">" => c > 0, ">=" => c >= 0, "<" => c < 0, "<=" => c <= 0, _ => false };
    }

    /// <summary>Checks if pattern contains unescaped wildcards (* or ?).</summary>
    private static bool HasUnescapedWildcard(string pattern)
    {
        for (int i = 0; i < pattern.Length; i++)
        {
            if (pattern[i] == '~' && i + 1 < pattern.Length)
            {
                i++; // Skip escaped character
                continue;
            }
            if (pattern[i] == '*' || pattern[i] == '?')
                return true;
        }
        return false;
    }

    /// <summary>Removes tilde escapes from a string (for literal comparison).</summary>
    private static string UnescapeTilde(string s)
    {
        if (!s.Contains('~')) return s;
        var sb = new System.Text.StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '~' && i + 1 < s.Length)
            {
                sb.Append(s[++i]); // Append the escaped character
            }
            else
            {
                sb.Append(s[i]);
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Wildcard match with tilde escape support.
    /// ~* = literal *, ~? = literal ?, ~~ = literal ~
    /// </summary>
    private static bool WildcardMatch(string text, string pattern)
    {
        text = text.ToLowerInvariant();
        pattern = pattern.ToLowerInvariant();
        
        return WildcardMatchInternal(text, 0, pattern, 0);
    }

    private static bool WildcardMatchInternal(string text, int ti, string pattern, int pi)
    {
        while (ti < text.Length || pi < pattern.Length)
        {
            if (pi < pattern.Length)
            {
                char pc = pattern[pi];
                
                // Handle tilde escape
                if (pc == '~' && pi + 1 < pattern.Length)
                {
                    char escaped = pattern[pi + 1];
                    if (ti < text.Length && text[ti] == escaped)
                    {
                        ti++;
                        pi += 2;
                        continue;
                    }
                    return false;
                }
                
                // Handle wildcard *
                if (pc == '*')
                {
                    // Skip consecutive stars
                    while (pi < pattern.Length && pattern[pi] == '*') pi++;
                    
                    // If * is at end of pattern, match rest of text
                    if (pi == pattern.Length) return true;
                    
                    // Try matching from each position
                    for (int i = ti; i <= text.Length; i++)
                    {
                        if (WildcardMatchInternal(text, i, pattern, pi))
                            return true;
                    }
                    return false;
                }
                
                // Handle wildcard ?
                if (pc == '?')
                {
                    if (ti < text.Length)
                    {
                        ti++;
                        pi++;
                        continue;
                    }
                    return false;
                }
                
                // Literal character match
                if (ti < text.Length && text[ti] == pc)
                {
                    ti++;
                    pi++;
                    continue;
                }
                
                return false;
            }
            else
            {
                // Pattern exhausted but text remains
                return false;
            }
        }
        
        return true;
    }
}
