using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Configuration;

namespace ExcelSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            String inputxls = ConfigurationManager.AppSettings.Get("InputPath");
            String outPath = ConfigurationManager.AppSettings.Get("OutputDir");
            String SheetName = ConfigurationManager.AppSettings.Get("SheetName");
            String OutputFormat = ConfigurationManager.AppSettings.Get("OutputFormat");

            var templatexls = new FileInfo(inputxls);
            using (var package = new ExcelPackage(templatexls))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets[SheetName];

                int iColCnt = ws.Dimension.End.Column ;
                int iRowCnt = ws.Dimension.End.Row ;

                var dataheader = new string[iColCnt+1];
                for (var headercol = 1; headercol <= iColCnt; headercol++)
                {
                    dataheader[headercol] = ws.Cells[1, headercol].Value.ValidateForNull();
                }

                for (int row = 2; row <= iRowCnt; row++)
                {
                    Console.Write("\rProcessing row :" + row);
                    var strbuilder = new StringBuilder();
                    for (var counter = 1; counter <= iColCnt; counter++)
                    {
                        var datacontent = ws.Cells[row, counter].Value.ValidateForNull();
                        strbuilder.AppendLine(dataheader[counter]);
                        strbuilder.AppendLine("-------------");
                        strbuilder.AppendLine(datacontent);
                    }
                    var outfile = Path.Combine(outPath, row + "." + OutputFormat);
                    File.WriteAllText(outfile, strbuilder.ToString());
               }
            }
        }

    }
    public static class Extension
    {
        /// <summary>
        /// Checks for null value, if null return empty string
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ValidateForNull(this Object obj)
        {
            //ternary operator. Go by stmt - 1 if cond is true and stmt-2 if its false
            return obj != null ? obj.ToString() : string.Empty;

        }
    }
}
