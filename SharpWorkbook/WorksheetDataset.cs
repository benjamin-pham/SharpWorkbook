using SharpWorkbook.ColumnCustoms;
using SharpWorkbook.Settings;
using System.Reflection;

namespace SharpWorkbook;

public class WorksheetDataset<T>
{
    public string SheetName { get; private set; }

    public IList<T> Data { get; private set; } = null!;

    public int RowsCount { get; private set; }

    public int ColumnsCount { get; private set; }

    public object[,] Data2d { get; private set; } = null!;

    public WorksheetSettings? Settings { get; private set; }

    public IList<IWorksheetColumn>? Columns { get; private set; }

    private int _addHeader => Settings!.HasHeader ? 1 : 0;

    private int _addOrdinal => Settings!.HasOrdinal ? 1 : 0;

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    private WorksheetDataset() { }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

    private WorksheetDataset(
        string sheetName,
        IList<T> data,
        WorksheetSettings? settings = null,
        IList<IWorksheetColumn>? columns = null)
    {
        EnsureSettingsIsNotNull(ref settings);
        Settings = settings;
        EnsureColumnsIsNotNull(ref columns);
        SheetName = sheetName;
        Data = data;
        Columns = columns;
        ColumnsCount = Columns!.Count;
        RowsCount = Data.Count + _addHeader;
        Data2d = new object[RowsCount, ColumnsCount];
    }

    public static WorksheetDataset<T> Create(
        string sheetName,
        IList<T> data,
        WorksheetSettings? settings = null,
        IList<IWorksheetColumn>? columns = null)
    {
        var dataset = new WorksheetDataset<T>(sheetName, data, settings, columns);

        dataset.BuildData2d();

        return dataset;
    }

    private void EnsureSettingsIsNotNull(ref WorksheetSettings? settings)
    {
        if (settings is null)
            settings = WorksheetSettings.Default;

        if (settings.HeaderStyle is null)
            settings.HeaderStyle = SheetStyle.DefaultHeader;

        if (settings.BodyStyle is null)
            settings.BodyStyle = SheetStyle.DefaultBody;

        if (settings.OrdinalStyle is null)
            settings.OrdinalStyle = SheetStyle.DefaultOrdinal;
    }

    private void EnsureColumnsIsNotNull(ref IList<IWorksheetColumn>? columns)
    {
        List<IWorksheetColumn> tmp = new List<IWorksheetColumn>();

        if (columns is null || !columns.Any())
        {
            tmp = typeof(T)
            .GetProperties()
            .Where(x => x.GetCustomAttribute<ColumnIgnoreAttribute>() == null)
            .Select(x => WorksheetColumnCustom.Create(x.Name, null, null))
            .ToList();
        }

        if (Settings!.HasOrdinal)
            tmp.Insert(0, WorksheetColumnOrdinal.Create("heh"));

        columns = tmp;
    }

    private void BuildData2d()
    {
        for (int i = 0; i < RowsCount; i++)
        {
            for (int j = 0; j < ColumnsCount; j++)
            {
                Data2d[i, j] = GetCellValue(i, j);
            }
        }
    }

    private object GetCellValue(int row, int column)
    {
        if (row == 0 && _addHeader > 0)
        {
            //return column == 0 && _addOrdinal > 0
            //    ? "_"
            //    : Columns![column - _addOrdinal].Name;
            return Columns![column].Name;
        }

        if (column == 0 && _addOrdinal > 0)
        {
            return row;
        }

        var propertyName = Columns![column].Name;
        return Data[row - _addHeader]?
            .GetType()
            .GetProperty(propertyName)?
            .GetValue(Data[row - _addHeader]) ?? string.Empty;
    }
}
