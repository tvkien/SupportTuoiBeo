using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;

namespace SyncfusionTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            try
            {
                using ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                using FileStream stream = new FileStream(@"D:\ConsoleTest\Syncfusion\Test.xlsx", FileMode.Open, FileAccess.Read);

                IWorkbook workbook = application.Workbooks.Open(stream);
                IWorksheet worksheet = workbook.Worksheets[1];

                var lstString = new List<string>();

                foreach(var a in worksheet.Range)
                {
                    if(a.Comment != null)
                    {
                        lstString.Add(a.Comment.Text);
                    }
                }

                var count = lstString.Count;

                Console.WriteLine($"Find {count} comment.");

                if(count > 0)
                {
                    foreach(var cmt in lstString)
                    {
                        Console.WriteLine(cmt);
                    }
                }

                //var abc12 = worksheet.Range["B2"].Comment;

                //var text = abc12.Text;

                //var abcd = worksheet.Cells.spe

                //application.DefaultVersion = ExcelVersion.Excel2013;
                //Create a workbook
                //IWorkbook workbook = application.Workbooks.Create(1);
                //IWorksheet worksheet = workbook.Worksheets[0];
                //The new workbook will have 5 worksheets
                //IWorkbook workbook = application.Workbooks.Create(5);
                //Creating a Sheet
                //IWorksheet sheet = workbook.Worksheets.Create();
                //Creating a Sheet with name “Sample”
                //IWorksheet namedSheet = workbook.Worksheets.Create("Sample123");

                //IComment comment = namedSheet.Range["A1"].AddComment();
                //var abc = namedSheet.Range["A1"].Comment;
                //comment.Text = "This is my comment";

                //Saving the workbook as stream
                //workbook.SaveAs(stream);
                //stream.Dispose();
            }
            catch (Exception ex)
            {

            }

            Console.ReadKey();
        }
    }
}
