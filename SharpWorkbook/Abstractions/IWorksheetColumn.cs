using SharpWorkbook.Abstractions;

namespace SharpWorkbook;

public interface IWorksheetColumn
{
    string Name { get; }
    double? Width { get; }
    IFormatCell? FormatCell { get; }
}
