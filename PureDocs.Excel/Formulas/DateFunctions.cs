namespace TVE.PureDocs.Excel.Formulas;

internal static class DateFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("NOW", Now, 0, 0, isVolatile: true);
        r.Register("TODAY", Today, 0, 0, isVolatile: true);
        r.Register("DATE", Date, 3, 3); r.Register("YEAR", Year, 1, 1);
        r.Register("MONTH", Month, 1, 1); r.Register("DAY", Day, 1, 1);
        r.Register("HOUR", Hour, 1, 1); r.Register("MINUTE", Minute, 1, 1);
        r.Register("SECOND", Second, 1, 1);
        r.Register("DATEVALUE", DateValue, 1, 1); r.Register("DAYS", Days, 2, 2);
        r.Register("EDATE", Edate, 2, 2); r.Register("EOMONTH", Eomonth, 2, 2);
        r.Register("WEEKDAY", Weekday, 1, 2); r.Register("WEEKNUM", Weeknum, 1, 2);
        r.Register("NETWORKDAYS", Networkdays, 2, 3); r.Register("DATEDIF", Datedif, 3, 3);
    }

    private static DateTime ToDate(double serial) => DateTime.FromOADate(serial);

    private static bool TryEvalDate(FormulaNode node, FormulaContext c, out DateTime dt, out FormulaValue error)
    {
        error = default; dt = default;
        if (!FormulaHelper.TryEvalDouble(node, c, out double d, out error)) return false;
        try { dt = ToDate(d); return true; }
        catch { error = FormulaValue.ErrorValue; return false; }
    }

    private static FormulaValue Now(List<FormulaNode> a, FormulaContext c) => FormulaValue.Number(DateTime.Now.ToOADate());
    private static FormulaValue Today(List<FormulaNode> a, FormulaContext c) => FormulaValue.Number(DateTime.Today.ToOADate());

    private static FormulaValue Date(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double y, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double m, out e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[2], c, out double d, out e)) return e;
        try { return FormulaValue.Number(new DateTime((int)y, (int)m, (int)d).ToOADate()); }
        catch { return FormulaValue.ErrorValue; }
    }

    private static FormulaValue Year(List<FormulaNode> a, FormulaContext c)
    { if (!TryEvalDate(a[0], c, out var dt, out var e)) return e; return FormulaValue.Number(dt.Year); }

    private static FormulaValue Month(List<FormulaNode> a, FormulaContext c)
    { if (!TryEvalDate(a[0], c, out var dt, out var e)) return e; return FormulaValue.Number(dt.Month); }

    private static FormulaValue Day(List<FormulaNode> a, FormulaContext c)
    { if (!TryEvalDate(a[0], c, out var dt, out var e)) return e; return FormulaValue.Number(dt.Day); }

    private static FormulaValue Hour(List<FormulaNode> a, FormulaContext c)
    { if (!TryEvalDate(a[0], c, out var dt, out var e)) return e; return FormulaValue.Number(dt.Hour); }

    private static FormulaValue Minute(List<FormulaNode> a, FormulaContext c)
    { if (!TryEvalDate(a[0], c, out var dt, out var e)) return e; return FormulaValue.Number(dt.Minute); }

    private static FormulaValue Second(List<FormulaNode> a, FormulaContext c)
    { if (!TryEvalDate(a[0], c, out var dt, out var e)) return e; return FormulaValue.Number(dt.Second); }

    private static FormulaValue DateValue(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string s, out var e)) return e;
        return DateTime.TryParse(s, out DateTime dt) ? FormulaValue.Number(dt.Date.ToOADate()) : FormulaValue.ErrorValue;
    }

    private static FormulaValue Days(List<FormulaNode> a, FormulaContext c)
    {
        if (!TryEvalDate(a[0], c, out var d1, out var e)) return e;
        if (!TryEvalDate(a[1], c, out var d2, out e)) return e;
        return FormulaValue.Number((d1 - d2).Days);
    }

    private static FormulaValue Edate(List<FormulaNode> a, FormulaContext c)
    {
        if (!TryEvalDate(a[0], c, out var dt, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double m, out e)) return e;
        try { return FormulaValue.Number(dt.AddMonths((int)m).ToOADate()); }
        catch { return FormulaValue.ErrorNum; }
    }

    private static FormulaValue Eomonth(List<FormulaNode> a, FormulaContext c)
    {
        if (!TryEvalDate(a[0], c, out var dt, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double m, out e)) return e;
        try
        {
            var d = dt.AddMonths((int)m);
            return FormulaValue.Number(new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month)).ToOADate());
        }
        catch { return FormulaValue.ErrorNum; }
    }

    private static FormulaValue Weekday(List<FormulaNode> a, FormulaContext c)
    {
        if (!TryEvalDate(a[0], c, out var dt, out var e)) return e;
        int type = 1; if (a.Count > 1) { if (!FormulaHelper.TryEvalDouble(a[1], c, out double t, out e)) return e; type = (int)t; }
        int dow = (int)dt.DayOfWeek;
        return type switch
        {
            1 => FormulaValue.Number(dow == 0 ? 1 : dow + 1),
            2 => FormulaValue.Number(dow == 0 ? 7 : dow),
            3 => FormulaValue.Number(dow == 0 ? 6 : dow - 1),
            _ => FormulaValue.ErrorNum,
        };
    }

    private static FormulaValue Weeknum(List<FormulaNode> a, FormulaContext c)
    {
        if (!TryEvalDate(a[0], c, out var dt, out var e)) return e;
        var cal = System.Globalization.CultureInfo.InvariantCulture.Calendar;
        int type = 1; if (a.Count > 1) { if (!FormulaHelper.TryEvalDouble(a[1], c, out double t, out e)) return e; type = (int)t; }
        var rule = type == 2 ? System.Globalization.CalendarWeekRule.FirstDay : System.Globalization.CalendarWeekRule.FirstDay;
        var firstDay = type == 2 ? DayOfWeek.Monday : DayOfWeek.Sunday;
        return FormulaValue.Number(cal.GetWeekOfYear(dt, rule, firstDay));
    }

    private static FormulaValue Networkdays(List<FormulaNode> a, FormulaContext c)
    {
        if (!TryEvalDate(a[0], c, out var start, out var e)) return e;
        if (!TryEvalDate(a[1], c, out var end, out e)) return e;
        var holidays = new HashSet<DateTime>();
        if (a.Count > 2)
        {
            var hv = a[2].Evaluate(c);
            if (hv.IsError) return hv;
            if (hv.IsArray)
            {
                foreach (var item in hv.ArrayVal.EnumerateValues())
                {
                    if (item.IsBlank) continue; // Skip blank cells
                    if (item.TryAsDouble(out double d))
                    {
                        try { holidays.Add(ToDate(d).Date); }
                        catch { return FormulaValue.ErrorNum; } // Invalid date returns #NUM!
                    }
                    else if (!item.IsBlank)
                    {
                        return FormulaValue.ErrorValue; // Non-numeric, non-blank returns #VALUE!
                    }
                }
            }
            else if (hv.TryAsDouble(out double d))
            {
                try { holidays.Add(ToDate(d).Date); }
                catch { return FormulaValue.ErrorNum; } // Invalid date returns #NUM!
            }
            else if (!hv.IsBlank)
            {
                return FormulaValue.ErrorValue; // Non-numeric, non-blank returns #VALUE!
            }
        }
        int dir = start <= end ? 1 : -1, count = 0;
        for (var dt = start.Date; dir > 0 ? dt <= end.Date : dt >= end.Date; dt = dt.AddDays(dir))
        {
            if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday && !holidays.Contains(dt))
                count++;
        }
        return FormulaValue.Number(count * dir);
    }

    private static FormulaValue Datedif(List<FormulaNode> a, FormulaContext c)
    {
        if (!TryEvalDate(a[0], c, out var start, out var e)) return e;
        if (!TryEvalDate(a[1], c, out var end, out e)) return e;
        if (!FormulaHelper.TryEvalString(a[2], c, out string unit, out e)) return e;
        if (start > end) return FormulaValue.ErrorNum;
        return unit.ToUpperInvariant() switch
        {
            "Y" => FormulaValue.Number(end.Year - start.Year - (end.Month < start.Month || (end.Month == start.Month && end.Day < start.Day) ? 1 : 0)),
            "M" => FormulaValue.Number((end.Year - start.Year) * 12 + end.Month - start.Month - (end.Day < start.Day ? 1 : 0)),
            "D" => FormulaValue.Number((end - start).Days),
            "MD" => FormulaValue.Number(end.Day >= start.Day ? end.Day - start.Day : DateTime.DaysInMonth(start.Year, start.Month) - start.Day + end.Day),
            "YM" => FormulaValue.Number(end.Month >= start.Month ? end.Month - start.Month - (end.Day < start.Day ? 1 : 0) : 12 - start.Month + end.Month - (end.Day < start.Day ? 1 : 0)),
            "YD" => FormulaValue.Number((new DateTime(end.Year, end.Month, end.Day) - new DateTime(end.Year, start.Month, start.Day)).Days),
            _ => FormulaValue.ErrorNum,
        };
    }
}
