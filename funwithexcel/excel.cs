using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _Excel = Microsoft.Office.Interop.Excel;

namespace funwithexcel
{
    class Excel
    {
        string path = "";
        Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        Range range;
        Range valueFound;
        Range rowSelect;

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path, ReadOnly:true);
            ws = wb.Worksheets[Sheet];
        }

        public string ReadCell(int i, int j)
        {
            double value = ws.Cells[i, j].Value2;
            return value.ToString();
        }

        public string FindValue(string value)
        {
            SetRange(1, 1, 1, LastRow());
            valueFound = ws.Cells.Find(value);
            return ValueRowandCol();
        }

        public int LastRow()
        {
            int lastRow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell).Row;
            return lastRow;
        }

        public int LastCol()
        {
            int lastCol = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell).Column;
            return lastCol;
        }

        private string ValueRowandCol()
        {
            string row = valueFound.Row.ToString();
            string col = valueFound.Column.ToString();

            return "Row " + row + " Col " + col;
        }

        public void SetRange(int col1, int row1, int col2, int row2)
        {
            range = ws.Range[ws.Cells[col1, row1], ws.Cells[col2, row2]];
        }

        public void SelectRow(int row)
        {
            rowSelect = ws.Range[ws.Cells[1, row], ws.Cells[LastCol(), row]];
        }

    }
}