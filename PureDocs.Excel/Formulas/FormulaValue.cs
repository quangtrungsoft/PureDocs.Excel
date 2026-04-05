using System.Globalization;
using System.Runtime.InteropServices;

namespace TVE.PureDocs.Excel.Formulas;

/// <summary>Error codes matching Excel error types.</summary>
public enum FormulaError : byte
{
    None = 0, Null, Div0, Value, Ref, Name, Num, NA, Calc, Spill
}

/// <summary>Kind of value stored in a FormulaValue.</summary>
public enum FormulaValueKind : byte
{
    Blank, Number, Text, Boolean, Error, Array
}

/// <summary>
/// Discriminated union for formula values. Stack-allocated, zero-GC for Number/Boolean/Blank/Error.
/// Replaces object? to eliminate boxing and enable error propagation without exceptions.
/// </summary>
[StructLayout(LayoutKind.Auto)]
public readonly struct FormulaValue : IEquatable<FormulaValue>
{
    private readonly FormulaValueKind _kind;
    private readonly FormulaError _error;
    private readonly double _number;
    private readonly object? _object; // string or ArrayValue

    private FormulaValue(FormulaValueKind kind, double number = 0, object? obj = null, FormulaError error = FormulaError.None)
    {
        _kind = kind; _number = number; _object = obj; _error = error;
    }

    // ── Cached singletons ───────────────────────────────────────────
    public static readonly FormulaValue Blank = new(FormulaValueKind.Blank);
    public static readonly FormulaValue Zero = new(FormulaValueKind.Number, 0);
    public static readonly FormulaValue One = new(FormulaValueKind.Number, 1);
    public static readonly FormulaValue True = new(FormulaValueKind.Boolean, 1);
    public static readonly FormulaValue False = new(FormulaValueKind.Boolean, 0);
    public static readonly FormulaValue EmptyString = new(FormulaValueKind.Text, 0, "");
    public static readonly FormulaValue ErrorDiv0 = new(FormulaValueKind.Error, error: FormulaError.Div0);
    public static readonly FormulaValue ErrorValue = new(FormulaValueKind.Error, error: FormulaError.Value);
    public static readonly FormulaValue ErrorRef = new(FormulaValueKind.Error, error: FormulaError.Ref);
    public static readonly FormulaValue ErrorName = new(FormulaValueKind.Error, error: FormulaError.Name);
    public static readonly FormulaValue ErrorNA = new(FormulaValueKind.Error, error: FormulaError.NA);
    public static readonly FormulaValue ErrorNum = new(FormulaValueKind.Error, error: FormulaError.Num);
    public static readonly FormulaValue ErrorNull = new(FormulaValueKind.Error, error: FormulaError.Null);

    // ── Factories ───────────────────────────────────────────────────
    public static FormulaValue Number(double v) => new(FormulaValueKind.Number, v);
    public static FormulaValue Text(string v) => new(FormulaValueKind.Text, 0, v ?? "");
    public static FormulaValue Boolean(bool v) => v ? True : False;
    public static FormulaValue Error(FormulaError e) => new(FormulaValueKind.Error, error: e);
    public static FormulaValue Array(ArrayValue a) => new(FormulaValueKind.Array, 0, a);

    // ── Type checks ─────────────────────────────────────────────────
    public FormulaValueKind Kind => _kind;
    public bool IsBlank => _kind == FormulaValueKind.Blank;
    public bool IsNumber => _kind == FormulaValueKind.Number;
    public bool IsText => _kind == FormulaValueKind.Text;
    public bool IsBoolean => _kind == FormulaValueKind.Boolean;
    public bool IsError => _kind == FormulaValueKind.Error;
    public bool IsArray => _kind == FormulaValueKind.Array;
    public bool IsNumeric => _kind is FormulaValueKind.Number or FormulaValueKind.Boolean or FormulaValueKind.Blank;
    public FormulaError ErrorCode => _error;

    // ── Raw access ──────────────────────────────────────────────────
    public double NumberValue => _number;
    public string TextValue => _object as string ?? "";
    public bool BooleanValue => _number != 0;
    public ArrayValue ArrayVal => _object as ArrayValue ?? ArrayValue.Empty;

    // ── Coercion (Excel semantics, NO exceptions) ───────────────────

    /// <summary>Coerce to number. Returns Error on failure, never throws.</summary>
    public FormulaValue CoerceToNumber()
    {
        return _kind switch
        {
            FormulaValueKind.Number => this,
            FormulaValueKind.Blank => Zero,
            FormulaValueKind.Boolean => Number(_number),
            FormulaValueKind.Error => this,
            FormulaValueKind.Text => double.TryParse((string)_object!, NumberStyles.Any,
                CultureInfo.InvariantCulture, out double d) ? Number(d) : ErrorValue,
            _ => ErrorValue,
        };
    }

    /// <summary>Try to get as double without throwing.</summary>
    public bool TryAsDouble(out double result)
    {
        switch (_kind)
        {
            case FormulaValueKind.Number: result = _number; return true;
            case FormulaValueKind.Blank: result = 0; return true;
            case FormulaValueKind.Boolean: result = _number; return true;
            case FormulaValueKind.Text:
                return double.TryParse((string)_object!, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out result);
            default: result = 0; return false;
        }
    }

    /// <summary>Coerce to string. Never throws.</summary>
    public string AsText()
    {
        return _kind switch
        {
            FormulaValueKind.Text => (string)_object!,
            FormulaValueKind.Number => _number.ToString(CultureInfo.InvariantCulture),
            FormulaValueKind.Boolean => _number != 0 ? "TRUE" : "FALSE",
            FormulaValueKind.Blank => "",
            FormulaValueKind.Error => ErrorToString(_error),
            _ => "",
        };
    }

    /// <summary>Coerce to bool. Returns Error on failure.</summary>
    public FormulaValue CoerceToBool()
    {
        return _kind switch
        {
            FormulaValueKind.Boolean => this,
            FormulaValueKind.Number => Boolean(_number != 0),
            FormulaValueKind.Blank => False,
            FormulaValueKind.Error => this,
            FormulaValueKind.Text => string.Equals((string)_object!, "TRUE", StringComparison.OrdinalIgnoreCase) ? True
                : string.Equals((string)_object!, "FALSE", StringComparison.OrdinalIgnoreCase) ? False
                : ErrorValue,
            _ => ErrorValue,
        };
    }

    // ── Backward compatibility ──────────────────────────────────────

    public object? ToObject() => _kind switch
    {
        FormulaValueKind.Number => _number,
        FormulaValueKind.Text => (string)_object!,
        FormulaValueKind.Boolean => _number != 0,
        FormulaValueKind.Blank => null,
        FormulaValueKind.Error => ErrorToString(_error),
        FormulaValueKind.Array => ((ArrayValue)_object!).ToObjectArray(),
        _ => null,
    };

    public static FormulaValue FromObject(object? value)
    {
        if (value == null) return Blank;
        if (value is double d) return Number(d);
        if (value is int i) return Number(i);
        if (value is long l) return Number(l);
        if (value is float f) return Number(f);
        if (value is decimal dec) return Number((double)dec);
        if (value is bool b) return Boolean(b);
        if (value is string s) return string.IsNullOrEmpty(s) ? Blank : Text(s);
        if (value is DateTime dt) return Number(dt.ToOADate());
        return Text(value.ToString() ?? "");
    }

    // ── Comparison (Excel semantics) ────────────────────────────────

    /// <summary>
    /// Excel-style equality comparison for formula evaluation.
    /// Rules: blank=0, blank="", case-insensitive text, cross-type=false.
    /// </summary>
    /// <remarks>
    /// <para>
    /// For Number comparison, uses epsilon of 1e-10 to handle floating-point
    /// representation errors (e.g., 0.1 + 0.2 = 0.3 returns TRUE).
    /// This matches Excel's display-level comparison behavior.
    /// </para>
    /// <para>
    /// Note: This is different from Equals() which uses bit-exact comparison
    /// for Dictionary/HashSet key purposes.
    /// </para>
    /// </remarks>
    public static bool AreEqual(FormulaValue a, FormulaValue b)
    {
        if (a.IsBlank && b.IsBlank) return true;
        if (a.IsBlank) return b.IsNumber ? Math.Abs(b._number) < 1e-10
            : b.IsText ? string.IsNullOrEmpty(b.TextValue) : b.IsBoolean && !b.BooleanValue;
        if (b.IsBlank) return a.IsNumber ? Math.Abs(a._number) < 1e-10
            : a.IsText ? string.IsNullOrEmpty(a.TextValue) : a.IsBoolean && !a.BooleanValue;
        if (a._kind != b._kind) return false;
        return a._kind switch
        {
            FormulaValueKind.Number => Math.Abs(a._number - b._number) < 1e-10,
            FormulaValueKind.Boolean => a._number == b._number,
            FormulaValueKind.Text => string.Equals(a.TextValue, b.TextValue, StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
    }

    /// <summary>Excel comparison ordering: Number/Blank &lt; Text &lt; Boolean.</summary>
    public static int Compare(FormulaValue a, FormulaValue b)
    {
        int ta = SortOrder(a), tb = SortOrder(b);
        if (ta != tb) return ta.CompareTo(tb);
        return ta switch
        {
            0 => (a.IsBlank ? 0.0 : a._number).CompareTo(b.IsBlank ? 0.0 : b._number),
            1 => string.Compare(a.TextValue, b.TextValue, StringComparison.OrdinalIgnoreCase),
            2 => a._number.CompareTo(b._number),
            _ => 0,
        };
    }

    private static int SortOrder(FormulaValue v) => v.Kind switch
    {
        FormulaValueKind.Blank or FormulaValueKind.Number => 0,
        FormulaValueKind.Text => 1,
        FormulaValueKind.Boolean => 2,
        _ => -1,
    };

    // ── Error helpers ───────────────────────────────────────────────

    public static string ErrorToString(FormulaError e) => e switch
    {
        FormulaError.Null => "#NULL!", FormulaError.Div0 => "#DIV/0!",
        FormulaError.Value => "#VALUE!", FormulaError.Ref => "#REF!",
        FormulaError.Name => "#NAME?", FormulaError.Num => "#NUM!",
        FormulaError.NA => "#N/A", FormulaError.Calc => "#CALC!",
        FormulaError.Spill => "#SPILL!", _ => "#ERROR!",
    };

    public static FormulaError ErrorFromString(string s) => s switch
    {
        "#NULL!" => FormulaError.Null, "#DIV/0!" => FormulaError.Div0,
        "#VALUE!" => FormulaError.Value, "#REF!" => FormulaError.Ref,
        "#NAME?" => FormulaError.Name, "#NUM!" => FormulaError.Num,
        "#N/A" => FormulaError.NA, _ => FormulaError.Value,
    };

    // ── IEquatable ──────────────────────────────────────────────────
    
    /// <summary>
    /// Bit-exact equality for use as Dictionary/HashSet key.
    /// Uses BitConverter.DoubleToInt64Bits for reliable double comparison.
    /// Note: This is different from AreEqual() which uses Excel semantics with epsilon.
    /// </summary>
    public bool Equals(FormulaValue o) => 
        _kind == o._kind 
        && BitConverter.DoubleToInt64Bits(_number) == BitConverter.DoubleToInt64Bits(o._number)
        && _error == o._error 
        && Equals(_object, o._object);
    
    public override bool Equals(object? obj) => obj is FormulaValue o && Equals(o);
    
    public override int GetHashCode() => 
        HashCode.Combine(_kind, BitConverter.DoubleToInt64Bits(_number), _object, _error);
    
    public override string ToString() => AsText();
    public static bool operator ==(FormulaValue l, FormulaValue r) => l.Equals(r);
    public static bool operator !=(FormulaValue l, FormulaValue r) => !l.Equals(r);
}
