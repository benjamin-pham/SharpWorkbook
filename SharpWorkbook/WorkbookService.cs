using SharpWorkbook.Abstractions;
using SharpWorkbook.Infrastructure;

namespace SharpWorkbook;

public class WorkbookService
{
    public IWorkbook CreateWorkbook()
    {
        return new Workbook();
    }   
}
