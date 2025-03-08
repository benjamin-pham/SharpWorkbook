using SharpWorkbook.Abstractions;
using SharpWorkbook.Formats;

namespace SharpWorkbook.ColumnCustoms;

public record WorksheetColumnOrdinal : IWorksheetColumn
{
    public string Name { get; private set; } = null!;

    public double? Width { get; private set; }

    public IFormatCell? FormatCell { get; set; }

    public static IWorksheetColumn Create(string name, double? width = null)
    {
        return new WorksheetColumnOrdinal()
        {
            Name = name,
            Width = width,
            FormatCell = NumberFormat.Instance
        };
    }
}
