using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace pwdCR
{
    class Program
    {
        public const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public const string SafeCharacterSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_-abcdefghijklmnopqrstuvwxyz1234567890!?.#$*%:;@^~";

        static void Main(string[] args)
        {
            Random rnd = new Random();
            XLWorkbook wb = new XLWorkbook();

            for (int j = 0; j < 21; j++)
            {
                try
                {
                    using (DataTable dt = new DataTable())
                    {
                        dt.Columns.Add("Buchstaben");
                        dt.Columns.Add("Lit. 1");
                        dt.Columns.Add("Lit. 2");

                        var max = SafeCharacterSet.Length;

                        for (int i = 0; i < 26; i += 2)
                        {
                            dt.Rows.Add(new object[3]
                            {
                                $"{Alphabet[i]}, {Alphabet[i + 1]}",
                                $"{SafeCharacterSet[rnd.Next(max)]}{SafeCharacterSet[rnd.Next(max)]}",
                                $"{SafeCharacterSet[rnd.Next(max)]}{SafeCharacterSet[rnd.Next(max)]}"

                            });
                        }

                        var ws = wb.Worksheets.Add(dt, $"Sheet {j + 1}");
                        ws.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Columns(1, 1).Width = 20;
                        ws.Columns(2, 3).Width = 12;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error occured while creatiing cheet N. {j + 1}:");
                    Console.WriteLine(ex.ToString());
                }
            }


            try
            {
                wb.SaveAs("pwdCR.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occured while saving excel file:");
                Console.WriteLine(ex.ToString());
            }

            Console.WriteLine();
            Console.WriteLine("Done.");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
