using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using ClosedXML.Excel;
using System.Data;

namespace ClosedXML
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //lendo arquivo
            var wb = new XLWorkbook(@"C:\Users\elian\source\repos\ClosedXML\ClosedXML\CNAE.xlsx");
            var planilha = wb.Worksheet(1);

            int i = 1;
            while (true)
            {
                var codigo = planilha.Cell(i, 2).Value.ToString();
                var descricao = planilha.Cell(i, 3).Value.ToString();

                if (String.IsNullOrEmpty(descricao))
                    break;

                Console.Write("Codigo: " + codigo.PadRight(10));
                Console.WriteLine("Descricao: " + descricao);

                i++;
            }

            Console.ReadLine();
            
            //Gravando arquivo xlsx
            //https://www.fourthbottle.com/2017/07/create-excel-files-in-csharp-using-open-xml.html

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Country");
            dt.Rows.Add("Venkatesh", "India");
            dt.Rows.Add("Santhosh", "USA");
            dt.Rows.Add("Venkat Sai", "Dubai");
            dt.Rows.Add("Venkat Teja", "Pakistan");
            ds.Tables.Add(dt);

            XLWorkbook wb2 = new XLWorkbook();
                wb2.Worksheets.Add(ds.Tables[0], ds.Tables[0].TableName);

            string AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\\2019.xlsx";

            wb2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            wb2.Style.Font.Bold = true;
            wb2.SaveAs(@"C:\Users\elian\teste.xlsx",false);
        }
    }
}
