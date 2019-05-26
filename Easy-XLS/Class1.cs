using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Excel_Sharp
{
    
    public class Class1
    {
        
        public static void WriteCells(IXLWorksheet worksheet,int Row,int Cell,string str="")
        {
            worksheet.Row(Row).Cell(Cell).Value = str;
        }
        public static string ReadCells(IXLWorksheet worksheet, int Row, int Cell)
        {
            return worksheet.Row(Row).Cell(Cell).Value.ToString();
        }


    }
}
