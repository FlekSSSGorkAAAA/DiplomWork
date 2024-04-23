using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace InclusiveEducationApp.Classes
{
    internal class ExcelHelperClass
    {
        public ExcelHelperClass(string pathFile)
        {
            Excel.Application app = new Excel.Application();
            Workbook wb = app.Workbooks.Open(pathFile);
        }
    }
}
