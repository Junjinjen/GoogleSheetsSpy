using OfficeOpenXml;
using System;

namespace GoogleSheetsSpy
{
    public class TextTrigger : ICellTrigger
    {
        public string Text { get; set; }

        public virtual bool IsTriggered(ExcelRange cell)
        {
            var text = cell.Text;
            return !string.IsNullOrEmpty(text) && Text.Contains(text, StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
