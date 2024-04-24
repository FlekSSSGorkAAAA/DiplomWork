using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System;

namespace InclusiveEducationApp.Classes
{
    internal class ExcelHelperClass
    {
        public List<ModelExcel> filledCells;
        public ExcelHelperClass(string pathFile)
        {

            //OpenFileDialog ofd = new OpenFileDialog();
            //ofd.ShowDialog();

            using (FileStream fs = new FileStream(pathFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook workbook;
                workbook = new XSSFWorkbook(pathFile);
                fs.Close();

                int sheetsCounter = workbook.NumberOfSheets;

                if (sheetsCounter > 1)
                {
                    for (int listCounter = 0; listCounter < sheetsCounter; listCounter++) // цикл по каждому листу
                    {
                        ISheet sheet = workbook.GetSheetAt(listCounter);

                        var firstRow = sheet.GetRow(0);
                        var secondRow = sheet.GetRow(1);
                        var thirdRow = sheet.GetRow(2);


                        for (int i = 3; i < sheet.LastRowNum; i++)
                        {
                            var row = sheet.GetRow(i);
                            
                            for(int j = 0; j < row.LastCellNum; j++)
                            {
                                int countValueStudent = Convert.ToInt32(row.GetCell(j).NumericCellValue);
                                
                                if (countValueStudent != 0)
                                {
                                    var valueOVZ = firstRow.GetCell(j).StringCellValue;
                                    var category = secondRow.GetCell(j).StringCellValue;
                                    var valueOfPerson = thirdRow.GetCell(j).StringCellValue;
                                    var collegeName = row.GetCell(0).StringCellValue;
                                    var specializationName = row.GetCell(1).StringCellValue;

                                    filledCells.Add(new ModelExcel(
                                    {
                                        Category = category,
                                        CollegeName = collegeName,
                                        Specialization = specializationName,
                                        StudentCount = countValueStudent,
                                        StudentStatus = valueOfPerson,
                                        SubCategory = valueOVZ
                                    }))
                                }
                                
                            }                         

                            //ICell cellNumb = row.GetCell(59);

                            int a = 10;
                        }
                    }
                }

                //using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                //{
                //    WorkbookPart workbookPart = doc.WorkbookPart;
                //    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                //    SharedStringTable sst = sstpart.SharedStringTable;
                //    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                //    Worksheet sheet = worksheetPart.Worksheet;


                //    var cells = sheet.Descendants<Cell>();
                //    var rows = sheet.Descendants<Row>();

                //    var a = rows.LongCount();

                //    //MessageBox.Show("Row count = {0}" + rows.LongCount().ToString());
                //    //MessageBox.Show("Cell count = {0}" + cells.LongCount().ToString());
                //    // One way: go through each cell in the sheet
                //    //foreach (Cell cell in cells)
                //    //{
                //    //    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                //    //    {
                //    //        int ssid = int.Parse(cell.CellValue.Text);
                //    //        string str = sst.ChildElements[ssid].InnerText;
                //    //        MessageBox.Show("Shared string {0}: {1}" + ssid.ToString() + str.ToString());
                //    //    }
                //    //    else if (cell.CellValue != null)
                //    //    {
                //    //        MessageBox.Show("Cell contents: {0}" + cell.CellValue.Text);
                //    //    }
                //    //}
                //    // Or... via each row

                //    foreach (Row row in rows)
                //    {
                //        foreach (Cell c in row.Elements<Cell>())
                //        {
                //            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                //            {
                //                int ssid = int.Parse(c.CellValue.Text);
                //                string str = sst.ChildElements[ssid].InnerText;
                //                MessageBox.Show("Shared string {0}: {1}" + ssid + str);
                //            }
                //            else if (c.CellValue != null)
                //            {
                //                MessageBox.Show("Cell contents: {0}", c.CellValue.Text);
                //            }
                //        }
                //    }
                //}

            }
        }
    }
}
