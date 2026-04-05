namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Walks an AST to collect all cell and range references that a formula depends on.
/// V2: uses RangeSet (area-based) instead of expanding ranges to individual cells.
/// 
/// Memory improvement: SUM(A:A) → 1 RangeArea (16 bytes) vs 1M CellAddress (~20MB).
/// </summary>
internal static class DependencyCollector
{
    /// <summary>
    /// Extracts all cell/range dependencies as a compact RangeSet.
    /// </summary>
    public static RangeSet CollectRangeSet(FormulaNode node, int sheetIndex = -1)
    {
        var set = new RangeSet();
        VisitCompact(node, set, sheetIndex);
        return set;
    }

    /// <summary>
    /// Legacy: Extracts all cell addresses (expands ranges to cells).
    /// WARNING: Can be expensive for large ranges. Use CollectRangeSet when possible.
    /// </summary>
    public static HashSet<CellAddress> Collect(FormulaNode node, int sheetIndex = -1)
    {
        return CollectRangeSet(node, sheetIndex).ExpandAll();
    }

    private static void VisitCompact(FormulaNode node, RangeSet set, int sheetIndex)
    {
        switch (node)
        {
            case CellReferenceNode cr:
                set.AddCell(CellAddress.FromReference(cr.Reference, sheetIndex));
                break;

            case RangeReferenceNode rr:
                // Area-based: store as one RangeArea (compact)
                CellReference.Parse(rr.StartRef, out int sr, out int sc);
                CellReference.Parse(rr.EndRef, out int er, out int ec);
                set.AddRange(sr, sc, er, ec);
                break;

            case BinaryOpNode bin:
                VisitCompact(bin.Left, set, sheetIndex);
                VisitCompact(bin.Right, set, sheetIndex);
                break;

            case UnaryOpNode un:
                VisitCompact(un.Operand, set, sheetIndex);
                break;

            case FunctionCallNode fn:
                foreach (var arg in fn.Arguments)
                    VisitCompact(arg, set, sheetIndex);
                break;

            case SheetCellReferenceNode scr:
                set.AddCell(CellAddress.FromReference(scr.CellReference, sheetIndex));
                break;

            case SheetRangeReferenceNode srr:
                CellReference.Parse(srr.StartRef, out int ssr, out int ssc);
                CellReference.Parse(srr.EndRef, out int ser, out int sec);
                set.AddRange(ssr, ssc, ser, sec);
                break;

            // NamedRangeNode: resolved at runtime, not tracked statically
            // Literals have no dependencies
        }
    }

    /// <summary>
    /// Checks if the AST contains any volatile function calls (NOW, RAND, etc.).
    /// </summary>
    public static bool ContainsVolatile(FormulaNode node)
    {
        return node switch
        {
            FunctionCallNode fn => FunctionRegistry.Default.IsVolatile(fn.FunctionName)
                || fn.Arguments.Any(a => ContainsVolatile(a)),
            BinaryOpNode bin => ContainsVolatile(bin.Left) || ContainsVolatile(bin.Right),
            UnaryOpNode un => ContainsVolatile(un.Operand),
            _ => false,
        };
    }
}
