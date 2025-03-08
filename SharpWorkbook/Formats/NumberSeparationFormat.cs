using SharpWorkbook.Abstractions;
using SharpWorkbook.Enumerations;

namespace SharpWorkbook.Formats;

public class NumberSeparationFormat : IFormatCell
{
    public static IFormatCell Instance => new NumberSeparationFormat();

    public string Format { get => "#,##0.00"; }

    public CellDataType DataType { get => CellDataType.Number; }
}
