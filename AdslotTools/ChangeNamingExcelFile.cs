using System;
using System.IO;
using Syncfusion.XlsIO;
using System.Collections.Generic;
using System.Diagnostics;

namespace ConsoleApplication8
{
    public class ChangeNamingExcelFile
    {
        public ChangeNamingExcelFile(string directory = @"C:\GitHub\symphony\Src\Application.MainModule.UnitTests\Resources\SymphonyImportTemplate\Success")
        {
            var files = Directory.GetFiles(directory);

            foreach (var e in files)
            {
                try
                {
                    Debug.WriteLine(e);
                    var ExcelEngine1 = new ExcelEngine();
                    var Application1 = ExcelEngine1.Excel;
                    var WorkBook1 = Application1.Workbooks.Open(e, ExcelOpenType.Automatic);

                    WorkBook1.Names["ClientPrepayDiscount"].Name = "ClientPrePayDiscount";
                    WorkBook1.Names["ClientPrepayDiscountType"].Name = "ClientPrePayDiscountType";

                    WorkBook1.Save();

                    WorkBook1.Close();
                    ExcelEngine1.Dispose();
                    Debug.WriteLine("DONE--------> " + e);
                }
                catch (Exception ex)
                {
                }
            }
        }
    }
}
