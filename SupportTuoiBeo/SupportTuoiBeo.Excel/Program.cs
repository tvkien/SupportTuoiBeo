using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SupportTuoiBeo.Data.Contexts;
using SupportTuoiBeo.Data.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SupportTuoiBeo.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            const string fileName = @"D:\Kien - Copy.xlsx";

            var userDetails = GetUserDetails(fileName, "tháng 12", Data.EnumThang.Thang12);

            InsertDatabase(userDetails);

            Console.ReadKey();
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
            if(index == 1)
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
                userDetail.MaKH = value;
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

        private static void InsertDatabase(List<UserDetails> userDetails)
        {
            using (var db = new ApplicationDbContext())
            {
                db.UserDetails.AddRange(userDetails);
                db.SaveChanges();
            }
        }
    }
}
