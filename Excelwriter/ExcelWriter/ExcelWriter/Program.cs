using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                var headerRow = new List<string[]>()
                {
                    new string[] { "ID", "First Name", "Last Name", "DOB" }
                };
                // Determine the header range (e.g. A1:D1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                // Popular header row data
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                //------------------------------------------write data------------------------------------------------------------
                DataTable dt = new DataTable();
                dt.Clear();
                dt.Columns.Add("ID");
                dt.Columns.Add("Fname");
                dt.Columns.Add("Lname");
                dt.Columns.Add("Dob");
                DataRow _ravi = dt.NewRow();
                _ravi["ID"] = "001";
                _ravi["Fname"] = "Kani";
                _ravi["Lname"] = "Gun";
                _ravi["Dob"] = "500";
                DataRow _ravi1 = dt.NewRow();
                _ravi1["ID"] = "002";
                _ravi1["Fname"] = "Yhn";
                _ravi1["Lname"] = "ravi";
                _ravi1["Dob"] = "600";
                dt.Rows.Add(_ravi);
                dt.Rows.Add(_ravi1);
                worksheet.Cells.LoadFromDataTable(dt,true);
                //------------------------------------------------------------------------------------------------
                FileInfo excelFile = new FileInfo(@"C:\test\test.xlsx");
                excel.SaveAs(excelFile);
            }

            Console.ReadLine();
        }
    }
}
