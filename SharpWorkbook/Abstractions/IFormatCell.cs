using SharpWorkbook.Enumerations;

namespace SharpWorkbook.Abstractions;

public interface IFormatCell
{
    static abstract IFormatCell Instance { get; }
    public string Format { get; }
    public CellDataType DataType { get; }
}
