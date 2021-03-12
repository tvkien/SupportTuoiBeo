using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SupportTuoiBeo.Data.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTuoiBeo.Excel
{
    public static class ExportExtension
    {
        public static void Export<T>(string path, string sheetName, List<T> entities)
        {
            using var spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();
            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                Name = sheetName,
                SheetId = 1 //sheet Id, anything but unique
            };

            sheets.Append(sheet);
            sheetData.Append(CreateHeaderRowForExcel());

            for (int i = 0; i < entities.Count; i++)
            {
                Row row = new Row();

                //loop column by column depending on the entity properties
                for (int column = 0; column < typeof(T).GetProperties().Count(); column++)
                {
                    var value = typeof(T).GetProperties()[column].GetValue(entities[i]).ToString();
                    var cell = CreateCell(value);
                    row.AppendChild(cell);
                }

                sheetData.Append(row);
            }

            worksheetPart.Worksheet = new Worksheet(sheetData);

            spreadsheet.WorkbookPart.Workbook.AppendChild(sheets);
            workbookPart.Workbook.Save();

            spreadsheet.Close();
        }

        private static Row CreateHeaderRowForExcel()
        {
            Row workRow = new Row();
            workRow.AppendChild(CreateCell("STT", 2U));
            workRow.AppendChild(CreateCell("Ma KH", 2U));
            workRow.AppendChild(CreateCell("Tinh", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 01", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 02", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 03", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 04", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 05", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 06", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 07", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 08", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 09", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 10", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 11", 2U));
            workRow.AppendChild(CreateCell("Doanh Thu Thang 12", 2U));
            workRow.AppendChild(CreateCell("Total 2020", 2U));
            return workRow;
        }

        private static Cell CreateCell(string text)
        {
            return new Cell
            {
                StyleIndex = 1U,
                DataType = ResolveCellDataTypeOnValue(text),
                CellValue = new CellValue(text)
            };
        }

        private static Cell CreateCell(string text, uint styleIndex)
        {
            return new Cell
            {
                StyleIndex = styleIndex,
                DataType = ResolveCellDataTypeOnValue(text),
                CellValue = new CellValue(text)
            };
        }

        private static EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            if (int.TryParse(text, out _) || double.TryParse(text, out _))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }

        public static List<UserDetails> GetUserDetails(
            string fileName,
            string sheetName,
            Data.EnumThang thang)
        {
            var userDetails = new List<UserDetails>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;
                SharedStringTablePart sstpart = wbPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sst = sstpart.SharedStringTable;
                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                var cells = wsPart.Worksheet.Descendants<Cell>();
                var rows = wsPart.Worksheet.Descendants<Row>();

                Console.WriteLine("Row count = {0}", rows.LongCount());
                Console.WriteLine("Cell count = {0}", cells.LongCount());

                foreach (Row row in rows)
                {
                    var index = 1;
                    var userDetail = new UserDetails();
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        string value = null;

                        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(c.CellValue.Text);
                            value = sst.ChildElements[ssid].InnerText;
                        }
                        else if (c.CellValue != null)
                        {
                            value = c.CellValue.Text;
                        }

                        Display(value, index, userDetail);
                        index++;
                    }

                    userDetail.Thang = thang;

                    userDetails.Add(userDetail);
                }
            }

            return userDetails;
        }


        private static void Display(string value, int index, UserDetails userDetail)
        {
            if (index == 1)
            {
                if (string.IsNullOrEmpty(value))
                {
                    userDetail.Ngay = 0;
                }
                else
                {
                    userDetail.Ngay = int.Parse(value);
                    Console.WriteLine("Ngay: " + value);
                }
            }

            if (index == 2)
            {
                userDetail.MaKH = value?.ToLower();
                Console.WriteLine("Ma KH: " + value);
            }

            if (index == 3)
            {
                userDetail.Tinh = value;
                Console.WriteLine("Tinh: " + value);
            }

            if (index == 4)
            {
                if (string.IsNullOrEmpty(value))
                {
                    userDetail.TienThanhToan = 0;
                }
                else
                {
                    var doubleValue = double.Parse(value);
                    userDetail.TienThanhToan = Convert.ToInt64(doubleValue);
                    Console.WriteLine("Thanh Toan: " + value);
                }
            }
        }

        private static string CreateCellReference(int column)
        {
            string result = string.Empty;
            //First is A
            //After Z, is AA
            //After ZZ, is AAA
            char firstRef = 'A';
            uint firstIndex = (uint)firstRef;
            while (column > 0)
            {
                int mod = (column - 1) % 26;
                result += (char)(firstIndex + mod);
                column = (column - mod) / 26;
            }

            return result;
        }
    }
}