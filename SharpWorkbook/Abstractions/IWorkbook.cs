using SharpWorkbook.Settings;

namespace SharpWorkbook.Abstractions;

public interface IWorkbook
{
    void AddSheet<T>(
        string sheetName,
        IList<T> data,
        IList<IWorksheetColumn>? columns = null,
        WorksheetSettings? settings = null)
        where T : class;
    void AddSheet<T>(WorksheetDataset<T> dataset) where T : class;
    void RemoveSheet(string sheetName);
    void SaveAs(string file);
    byte[] GetBytes();
}
