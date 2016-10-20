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
        public DateTime Date
        {
            get { return _date; }
            set { _date = value; }
        }
        public string[,] Excel { get; set; }
    }

    public interface IExcelMaker
    {
        void ToMakeExcel(Task<IEnumerable<Bill>> xmls, Task<object[,]> excel);
    }
    public class ExcelMaker : IExcelMaker
    {
        public async void ToMakeExcel(Task<IEnumerable<Bill>> xmls, Task<object[,]> excel)
        => await Task.Run( async() =>
        {
            #region XmlToArray
            var _xmls = await xmls;
            var xmlData = _xmls.Select(xml =>
            {
                ExcelData ExcelData = new ExcelData();
                ExcelData.Date = DateTime.Parse(xml.Title.B_start, null, DateTimeStyles.RoundtripKind);
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
            }).OrderBy(xml=>xml.Date).ToList();
            #endregion


            #region ToExcel
            var excalResult = excel.Result;

            for (int i = 4; i < xmlData.Count + 4; i++)
            {
                excalResult = Sorts(excalResult, xmlData[i - 4].Excel, i);
            }

            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWb;
            Excel.Worksheet xlSht;
            //var abs = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            xlWb = xlApp.Workbooks.Add(System.Reflection.Missing.Value);
            xlSht = (Excel.Worksheet)xlWb.Sheets[1];

            xlSht.Range["A1"]
                .Resize[excalResult.GetUpperBound(0), excalResult.GetUpperBound(1)]
                .Value = excalResult; 
            var excelDate = xmlData.Select(xml => xml.Date.ToString("MMM.yyyy")).ToArray();          
            xlSht.Range["D2"].Resize[1, excelDate.Length].Value = excelDate;
            xlSht.Columns["B:Z"].AutoFit();

            xlWb.Close(true);
            xlApp.Quit();
            #endregion
        });
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
