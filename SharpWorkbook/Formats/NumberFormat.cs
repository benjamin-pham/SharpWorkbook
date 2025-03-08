using SharpWorkbook.Abstractions;
using SharpWorkbook.Enumerations;

namespace SharpWorkbook.Formats;

internal class NumberFormat : IFormatCell
{
    public static IFormatCell Instance => new NumberFormat();

    public string Format { get => string.Empty; }

    public CellDataType DataType => CellDataType.Number;
}
