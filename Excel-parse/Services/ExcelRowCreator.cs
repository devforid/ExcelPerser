using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excel_parse.Services
{
    public class ExcelRowCreator
    {
        private IRow _row;
        private ICreationHelper _creationHelper;

        public ExcelRowCreator(IRow row, ICreationHelper creationHelper)
        {
            this._creationHelper = creationHelper;
            this._row = row;
        }

        public void CreateCell(int cellNumber, string value)
        {
            var cell = this._row.CreateCell(cellNumber);
            cell.SetCellValue(this._creationHelper.CreateRichTextString(value));
        }
    }
}
