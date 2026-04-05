using System.Globalization;
using System.Text;

namespace TVE.PureDocs.Excel.Formulas;

internal static class TextFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("LEFT", Left, 1, 2); r.Register("RIGHT", Right, 1, 2);
        r.Register("MID", Mid, 3, 3); r.Register("LEN", Len, 1, 1);
        r.Register("UPPER", Upper, 1, 1); r.Register("LOWER", Lower, 1, 1);
        r.Register("PROPER", Proper, 1, 1); r.Register("TRIM", Trim, 1, 1);
        r.Register("CLEAN", Clean, 1, 1); r.Register("SUBSTITUTE", Substitute, 3, 4);
        r.Register("REPLACE", Replace, 4, 4);
        r.Register("CONCATENATE", Concatenate, 1); r.Register("CONCAT", Concatenate, 1);
        r.Register("TEXT", TextFn, 2, 2); r.Register("VALUE", Value, 1, 1);
        r.Register("FIND", Find, 2, 3); r.Register("SEARCH", Search, 2, 3);
        r.Register("REPT", Rept, 2, 2); r.Register("EXACT", Exact, 2, 2);
        r.Register("CHAR", CharFn, 1, 1); r.Register("CODE", Code, 1, 1);
        r.Register("T", T, 1, 1); r.Register("TEXTJOIN", TextJoin, 3);
    }

    private static FormulaValue Left(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        int n = 1; if (a.Count > 1) { if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e; n = (int)d; }
        if (n < 0) return FormulaValue.ErrorValue;
        return FormulaValue.Text(s.Length <= n ? s : s[..n]);
    }

    private static FormulaValue Right(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        int n = 1; if (a.Count > 1) { if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e; n = (int)d; }
        if (n < 0) return FormulaValue.ErrorValue;
        return FormulaValue.Text(s.Length <= n ? s : s[^n..]);
    }

    private static FormulaValue Mid(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double sv, out e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[2], c, out double nv, out e)) return e;
        int start = (int)sv - 1, num = (int)nv;
        if (start < 0 || num < 0) return FormulaValue.ErrorValue;
        if (start >= s.Length) return FormulaValue.EmptyString;
        return FormulaValue.Text(s.Substring(start, Math.Min(num, s.Length - start)));
    }

    private static FormulaValue Len(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return FormulaValue.Number(s.Length);
    }

    private static FormulaValue Upper(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return FormulaValue.Text(s.ToUpperInvariant());
    }

    private static FormulaValue Lower(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return FormulaValue.Text(s.ToLowerInvariant());
    }

    private static FormulaValue Proper(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        var sb = new StringBuilder(s.Length); bool cap = true;
        foreach (char ch in s)
        {
            sb.Append(cap ? char.ToUpper(ch) : char.ToLower(ch));
            cap = !char.IsLetter(ch);
        }
        return FormulaValue.Text(sb.ToString());
    }

    private static FormulaValue Trim(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return FormulaValue.Text(string.Join(' ', s.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
    }

    private static FormulaValue Clean(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return FormulaValue.Text(new string(s.Where(ch => ch >= 32).ToArray()));
    }

    private static FormulaValue Substitute(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        if (!FormulaHelper.TryEvalString(a[1], c, out string old, out e)) return e;
        if (!FormulaHelper.TryEvalString(a[2], c, out string nw, out e)) return e;
        if (a.Count > 3)
        {
            if (!FormulaHelper.TryEvalDouble(a[3], c, out double iv, out e)) return e;
            int inst = (int)iv, found = 0, pos = 0;
            while (pos < s.Length)
            {
                int idx = s.IndexOf(old, pos, StringComparison.Ordinal);
                if (idx < 0) break;
                if (++found == inst)
                    return FormulaValue.Text(s[..idx] + nw + s[(idx + old.Length)..]);
                pos = idx + old.Length;
            }
            return FormulaValue.Text(s);
        }
        return FormulaValue.Text(s.Replace(old, nw));
    }

    private static FormulaValue Replace(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double sv, out e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[2], c, out double nv, out e)) return e;
        if (!FormulaHelper.TryEvalString(a[3], c, out string nw, out e)) return e;
        int start = (int)sv - 1, num = (int)nv;
        if (start < 0) start = 0;
        if (start >= s.Length) return FormulaValue.Text(s + nw);
        num = Math.Min(num, s.Length - start);
        return FormulaValue.Text(s[..start] + nw + s[(start + num)..]);
    }

    private static FormulaValue Concatenate(List<FormulaNode> a, FormulaContext c)
    {
        var sb = new StringBuilder();
        foreach (var arg in a)
        {
            var v = arg.Evaluate(c);
            if (v.IsError) return v;
            sb.Append(v.AsText());
        }
        return FormulaValue.Text(sb.ToString());
    }

    private static FormulaValue TextFn(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (!FormulaHelper.TryEvalString(a[1], c, out string fmt, out e)) return e;
        string netFmt = MapExcelFormat(fmt);
        return FormulaValue.Text(v.ToString(netFmt, CultureInfo.InvariantCulture));
    }

    private static FormulaValue Value(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double d)
            ? FormulaValue.Number(d) : FormulaValue.ErrorValue;
    }

    private static FormulaValue Find(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string find, out var e)) return e;
        if (!FormulaHelper.TryEvalString(a[1], c, out string text, out e)) return e;
        int start = 0; if (a.Count > 2) { if (!FormulaHelper.TryEvalDouble(a[2], c, out double sv, out e)) return e; start = (int)sv - 1; }
        int idx = text.IndexOf(find, start, StringComparison.Ordinal);
        return idx < 0 ? FormulaValue.ErrorValue : FormulaValue.Number(idx + 1);
    }

    private static FormulaValue Search(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string find, out var e)) return e;
        if (!FormulaHelper.TryEvalString(a[1], c, out string text, out e)) return e;
        int start = 0; if (a.Count > 2) { if (!FormulaHelper.TryEvalDouble(a[2], c, out double sv, out e)) return e; start = (int)sv - 1; }
        int idx = text.IndexOf(find, start, StringComparison.OrdinalIgnoreCase);
        return idx < 0 ? FormulaValue.ErrorValue : FormulaValue.Number(idx + 1);
    }

    private static FormulaValue Rept(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double nv, out e)) return e;
        int n = (int)nv; if (n < 0) return FormulaValue.ErrorValue;
        return FormulaValue.Text(string.Concat(Enumerable.Repeat(s, n)));
    }

    private static FormulaValue Exact(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s1, out var e)) return e;
        if (!FormulaHelper.TryEvalString(a[1], c, out string s2, out e)) return e;
        return FormulaValue.Boolean(string.Equals(s1, s2, StringComparison.Ordinal));
    }

    private static FormulaValue CharFn(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        int code = (int)v; if (code < 1 || code > 255) return FormulaValue.ErrorValue;
        return FormulaValue.Text(((char)code).ToString());
    }

    private static FormulaValue Code(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return s.Length == 0 ? FormulaValue.ErrorValue : FormulaValue.Number(s[0]);
    }

    private static FormulaValue T(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        return v.IsText ? v : FormulaValue.EmptyString;
    }

    /// <summary>
    /// TEXTJOIN - Joins text strings with a delimiter.
    /// </summary>
    /// <remarks>
    /// Excel limits cell text to 32,767 characters. This function returns #VALUE!
    /// if the result would exceed this limit.
    /// </remarks>
    private static FormulaValue TextJoin(List<FormulaNode> a, FormulaContext c)
    {
        const int MaxCellTextLength = 32767; // Excel cell text limit
        
        if (!FormulaHelper.TryEvalString(a[0], c, out string delim, out var e)) return e;
        var boolVal = a[1].Evaluate(c);
        if (boolVal.IsError) return boolVal;
        bool ignoreEmpty = boolVal.CoerceToBool().BooleanValue;
        var parts = new List<string>();
        
        for (int i = 2; i < a.Count; i++)
        {
            var v = a[i].Evaluate(c);
            if (v.IsError) return v;
            if (v.IsArray)
            {
                foreach (var item in v.ArrayVal.EnumerateValues())
                {
                    if (!ignoreEmpty || (!item.IsBlank && !(item.IsText && item.TextValue.Length == 0)))
                        parts.Add(item.AsText());
                }
            }
            else if (!ignoreEmpty || (!v.IsBlank && !(v.IsText && v.TextValue.Length == 0)))
            {
                parts.Add(v.AsText());
            }
        }
        
        // Calculate result length before joining to avoid memory allocation
        int resultLength = parts.Count > 0 
            ? parts.Sum(p => p.Length) + delim.Length * (parts.Count - 1)
            : 0;
            
        if (resultLength > MaxCellTextLength)
            return FormulaValue.ErrorValue; // Excel returns #VALUE! when result too long
        
        return FormulaValue.Text(string.Join(delim, parts));
    }

    private static string MapExcelFormat(string fmt) => fmt switch
    {
        "0" => "0", "0.00" => "0.00", "#,##0" => "#,##0", "#,##0.00" => "#,##0.00",
        "0%" => "0%", "0.00%" => "0.00%", "#,##0;-#,##0" => "#,##0;-#,##0",
        _ => fmt.Replace('#', '0'),
    };
}
