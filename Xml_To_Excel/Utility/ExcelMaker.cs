using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Xml2CSharp;
using System.Globalization;

namespace Xml_To_Excel.Utility
{
    public class ExcelData
    {
        DateTime _date;
        public string Date
        {
            get { return _date.ToString("mm.yy"); }
            set { _date = DateTime.Parse(value, null, DateTimeStyles.RoundtripKind); }
        }
        public string[,] Excel { get; set; }
    }
    public class ExcelMaker
    {
        async void ToMakeExcel(Task<IEnumerable<Bill>> xmls, Task<object[,]> excel)
        {
            #region XmlToArray
            ExcelData ExcelData = new ExcelData();
            var _xmls = await xmls;
            var excelData = _xmls.Select(xml =>
            {
                ExcelData.Date = xml.Title.B_start;
                var temp = xml.Ch_details.Charges_d.Charge_d.Select(d =>
                    new { call = d.C_num, tot = d.C_tot }).ToArray();

                string[,] arrXml = new string[temp.Length, 2];
                for (int i = 0; i < arrXml.GetUpperBound(0); i++)
                {
                    arrXml[i, 0] = temp[i].call;
                    arrXml[i, 1] = temp[i].tot;
                }
                ExcelData.Excel = arrXml;
                return ExcelData;
            }).ToList();
            #endregion


            #region ToExcel
            Excel.Application xlApp = new Excel.Application();

            for (int i = 4; i < excelData.Count; i++)
            {
                var sortedArray = Sorts(excel.Result, excelData[i].Excel, i);
                Excel.Workbook xlWb;
                Excel.Worksheet xlSht;
                var abs = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

                //Книга.
                xlWb = xlApp.Workbooks.Add(System.Reflection.Missing.Value);
                //Таблица.
                xlSht = (Excel.Worksheet)xlWb.Sheets[1];

                xlSht.Cells[$"{abs[i]}2"] = excelData[i].Date;
                xlSht.Range["A1"]
                    .Resize[sortedArray.GetUpperBound(0), sortedArray.GetUpperBound(1)]
                    .Value = sortedArray; //выгрузка массива на лист Excel начиная с А1
                xlSht.Columns["B:Z"].AutoFit();


                xlWb.Close(true);//закрываем файл и сохраняем изменения, если не сохранять, то false   
            }

            xlApp.Quit(); //закрываем Excel
            #endregion
        }
        public static object[,] Sorts(object[,] arrExcel, string[,] arrXml, int n)
        {
            for (int i = 1; i <= arrExcel.GetLength(0); i++)
                for (int j = 0; j < arrXml.GetLength(0); j++)
                    if (arrExcel[i, 2]?.ToString().Trim() == arrXml[j, 0]?.Trim().Substring(2))
                    {
                        arrExcel[i, n] = arrXml[j, 1];
                    }

            return arrExcel;
        }
    }
}
