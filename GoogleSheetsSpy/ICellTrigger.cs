using OfficeOpenXml;

namespace GoogleSheetsSpy
{
    public interface ICellTrigger
    {
        bool IsTriggered(ExcelRange cell);
    }
}
