using SharpWorkbook.Abstractions;

namespace SharpWorkbook.ColumnCustoms;

internal class WorksheetColumnAttribute : Attribute
{
    public string Name { get; private set; }
    /// <summary>
    /// null value is width auto
    /// </summary>
    public double? Width { get; set; }
    public IFormatCell? FormatCell { get; set; }
    public WorksheetColumnAttribute(string name)
    {
        Name = name;
    }
    public IWorksheetColumn GetWorksheetColumn()
    {
        return WorksheetColumnCustom.Create(Name, Width, FormatCell);
    }
}
