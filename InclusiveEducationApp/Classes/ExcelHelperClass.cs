using Microsoft.Win32;
using System.IO;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace InclusiveEducationApp.Classes
{
    internal class ExcelHelperClass
    {
        public ExcelHelperClass(string pathFile)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.ShowDialog();

            using (FileStream fs = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;
                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>();
                    int a = 10;
                    //MessageBox.Show("Row count = {0}" + rows.LongCount().ToString());
                    //MessageBox.Show("Cell count = {0}" + cells.LongCount().ToString());
                    // One way: go through each cell in the sheet
                    //foreach (Cell cell in cells)
                    //{
                    //    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                    //    {
                    //        int ssid = int.Parse(cell.CellValue.Text);
                    //        string str = sst.ChildElements[ssid].InnerText;
                    //        MessageBox.Show("Shared string {0}: {1}" + ssid.ToString() + str.ToString());
                    //    }
                    //    else if (cell.CellValue != null)
                    //    {
                    //        MessageBox.Show("Cell contents: {0}" + cell.CellValue.Text);
                    //    }
                    //}
                    // Or... via each row
                    foreach (Row row in rows)
                    {
                        foreach (Cell c in row.Elements<Cell>())
                        {
                            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                            {
                                int ssid = int.Parse(c.CellValue.Text);
                                string str = sst.ChildElements[ssid].InnerText;
                                MessageBox.Show("Shared string {0}: {1}" + ssid + str);
                            }
                            else if (c.CellValue != null)
                            {
                                MessageBox.Show("Cell contents: {0}", c.CellValue.Text);
                            }
                        }
                    }
                }
            }
        }
    }
}
