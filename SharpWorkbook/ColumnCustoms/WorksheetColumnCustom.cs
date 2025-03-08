using SharpWorkbook.Abstractions;

namespace SharpWorkbook.ColumnCustoms;

public record WorksheetColumnCustom : IWorksheetColumn
{
    public string Name { get; protected set; } = null!;
    public double? Width { get; protected set; }
    public IFormatCell? FormatCell { get; protected set; }
    public static IWorksheetColumn Create(string name, double? width, IFormatCell? formatCell)
    {
        return new WorksheetColumnCustom()
        {
            Name = name,
            Width = width,
            FormatCell = formatCell
        };
    }
}
