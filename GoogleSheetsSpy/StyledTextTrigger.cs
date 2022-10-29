using OfficeOpenXml;

namespace GoogleSheetsSpy
{
    public class StyledTextTrigger : TextTrigger
    {
        public string BackgroundColor { get; set; }

        public override bool IsTriggered(ExcelRange cell)
        {
            return cell.Style.Fill.BackgroundColor.Rgb == BackgroundColor && base.IsTriggered(cell);
        }
    }
}
